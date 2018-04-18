using ExcelUtil;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Encrypter;

namespace EncrypterApp
{
    class Program
    {
        static string GetPasswordFromConsole()
        {
            ConsoleKeyInfo key;
            string password = "";
            do
            {
                key = Console.ReadKey(true);
                // Backspace Should Not Work
                if (key.Key != ConsoleKey.Backspace && key.Key != ConsoleKey.Enter)
                {
                    password += key.KeyChar;
                    Console.Write("*");
                }
                else
                {
                    if (key.Key == ConsoleKey.Backspace && password.Length > 0)
                    {
                        password = password.Substring(0, (password.Length - 1));
                        Console.Write("\b \b");
                    }
                }
            }
            // Stops Receving Keys Once Enter is Pressed
            while (key.Key != ConsoleKey.Enter);

            Console.WriteLine();
            return password;
        }

        static void Main(string[] args)
        {
            string firstPassword = "";
            string secondPassword = "";
            do
            {
                if(firstPassword != secondPassword)
                    Console.WriteLine("Passwords do not match!");

                Console.Write("Enter your password: ");
                firstPassword = GetPasswordFromConsole();
                Console.Write("Confirm your password: ");
                secondPassword = GetPasswordFromConsole();
            }
            while (firstPassword != secondPassword);

#if DEBUG
            Console.WriteLine("The Password You entered is : " + firstPassword);
            Console.WriteLine("The Hash : " + HashUtil.Encrypt(firstPassword));
#endif
            try
            {
                // open config
                string configPath = Directory.GetCurrentDirectory() + @"\CatalogPrinter.config";
                if (!File.Exists(configPath))
                    throw new Exception($"Config file " + configPath + " not found!");
                ExeConfigurationFileMap configMap = new ExeConfigurationFileMap();
                configMap.ExeConfigFilename = configPath;
                Configuration config = ConfigurationManager.OpenMappedExeConfiguration(configMap, ConfigurationUserLevel.None);
                var appSettings = config.GetSection("appSettings") as AppSettingsSection;
                // get config values
                string oldHash = appSettings.Settings["password"].Value;
                string masterCatalog = appSettings.Settings["masterCatalog"].Value;

                // encrypt new password
                string encryptedPassword = HashUtil.Encrypt(firstPassword);

                // try opening catalog with old password and change the password
                if (!File.Exists(masterCatalog))
                    throw new Exception($"Workbook " + masterCatalog + " not found!");
                Console.WriteLine("Changing the password...");
                string oldPassword = HashUtil.Decrypt(oldHash);
                Workbook wb = ExcelUtility.GetWorkbook(masterCatalog, oldPassword);
                if (wb == null)
                    throw new Exception($"Wrong password for workbook " + masterCatalog + "!");
                wb.Password = firstPassword;
                ExcelUtility.CloseWorkbook(wb, true);

                // write new encrypted password to config
                appSettings.Settings["password"].Value = encryptedPassword;
                config.Save(ConfigurationSaveMode.Modified);

                Console.WriteLine("Success! Press Enter to close the application.");

            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine("Error! Could not change the password!");
            }

            Console.ReadLine();
        }
    }
}
