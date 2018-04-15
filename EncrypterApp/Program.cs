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
        static void Main(string[] args)
        {
            ConsoleKeyInfo key;

//            string user = "";
//            Console.Write("Enter your username: ");

//            do
//            {
//                key = Console.ReadKey(true);
//                // Backspace Should Not Work
//                if (key.Key != ConsoleKey.Backspace && key.Key != ConsoleKey.Enter)
//                {
//                    user += key.KeyChar;
//                    Console.Write(key.KeyChar);
//                }
//                else
//                {
//                    if (key.Key == ConsoleKey.Backspace && user.Length > 0)
//                    {
//                        user = user.Substring(0, (user.Length - 1));
//                        Console.Write("\b \b");
//                    }
//                }
//            }
//            // Stops Receving Keys Once Enter is Pressed
//            while (key.Key != ConsoleKey.Enter);

//            Console.WriteLine();
//#if DEBUG
//            Console.WriteLine("The User You entered is : " + user);
//#endif

            string pass = "";
            Console.Write("Enter your password: ");

            do
            {
                key = Console.ReadKey(true);
                // Backspace Should Not Work
                if (key.Key != ConsoleKey.Backspace && key.Key != ConsoleKey.Enter)
                {
                    pass += key.KeyChar;
                    Console.Write("*");
                }
                else
                {
                    if (key.Key == ConsoleKey.Backspace && pass.Length > 0)
                    {
                        pass = pass.Substring(0, (pass.Length - 1));
                        Console.Write("\b \b");
                    }
                }
            }
            // Stops Receving Keys Once Enter is Pressed
            while (key.Key != ConsoleKey.Enter);

            Console.WriteLine();
#if DEBUG
            Console.WriteLine("The Password You entered is : " + pass);
            Console.WriteLine("The Hash : " + Encrypter.HashUtil.Encrypt(pass));
#endif
            string configPath = ConfigurationManager.AppSettings["ConfigPath"];
            XMLUtility.WriteToXml(configPath, new KeyValuePair<string, string>("password", Encrypter.HashUtil.Encrypt(pass)));
            Console.ReadLine();
        }
    }
}
