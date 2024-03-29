﻿using System;
using System.Globalization;
using System.Text;
using System.Threading;
using Xceed.Words.NET;

namespace URG_Console
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("uk-UA", false);
            Thread.CurrentThread.CurrentUICulture = new CultureInfo("uk-UA", false);

            ClearLicenseMsg();

            string hardcodedPath = @"C:\Users\Neo\Desktop\Weekly";

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

        public static void ClearLicenseMsg()
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
