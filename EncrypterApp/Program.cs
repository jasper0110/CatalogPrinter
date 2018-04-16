using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XMLUtil;

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
            Console.WriteLine("The Hash : " + Encrypter.HashUtil.Encrypt(firstPassword));
#endif
            string configPath = ConfigurationManager.AppSettings["ConfigPath"];
            XMLUtility.WriteToXml(configPath, new KeyValuePair<string, string>("password", Encrypter.HashUtil.Encrypt(firstPassword)));
            Console.WriteLine("Press Enter to close the application...");
            Console.ReadLine();
        }
    }
}
