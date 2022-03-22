using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Document.NET;
using Xceed.Words.NET;
using HtmlAgilityPack;
using System.Threading;
using System.Globalization;
using System.IO;

namespace URG_Console
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("uk-UA", false);
            Thread.CurrentThread.CurrentUICulture = new CultureInfo("uk-UA", false);

            clearLicenseMsg();

            string hardcodedPath = @"C:\Users\Neo\Desktop\New folder";

            Console.WriteLine("=======Ukrinform Report Generator=======");
            Console.WriteLine("Select folder location mode:\n");
            Console.WriteLine("1. Use hardcoded location");
            Console.WriteLine($"({hardcodedPath})");
            Console.WriteLine("2. Manually specify folder path");

            int consoleInput = -1;
            bool isValidInput = false;

            do
            {
                Console.Write("> ");
                string consoleRead = Console.ReadLine();
                isValidInput = int.TryParse(consoleRead, out consoleInput) && (consoleInput >= 1 && consoleInput <= 2);

                if (!isValidInput)
                {
                    Console.WriteLine("Invalid input! Try again.");
                }

            } while (!isValidInput);

            if (consoleInput == 1)
            {
                Console.WriteLine();
                ReportGenerator report = new ReportGenerator(hardcodedPath);
            }
            else if (consoleInput == 2)
            {
                string folderPathInput = String.Empty;
                Console.WriteLine("Input full folder path: ");
                Console.WriteLine("(For faster access copy and paste full folder path from Explorer)");
                Console.Write("> ");
                folderPathInput = Console.ReadLine();

                Console.WriteLine();
                ReportGenerator report = new ReportGenerator(folderPathInput);
            }
            else
                throw new ArgumentOutOfRangeException("Invalid menu input was provided!");

        }

        public static void clearLicenseMsg()
        {
            // Doing this to trigger DocX free license message in console to clear it
            try
            {
                DocX.Load((string)null);
            }
            catch (Exception ex) { Console.Clear(); }

        }

    }

}
