using System;
using System.Globalization;
using System.Text;
using System.Threading;
using System.IO;
using Microsoft.Extensions.Configuration;
using Xceed.Words.NET;

namespace URG_Console
{
    class Program
    {
        private static string DefaultPath;
        private static readonly DayOfWeek ReportEndDay = DayOfWeek.Sunday;

        static void Main(string[] args)
        {
            InitializeEnvironment();
            ClearLicenseMsg();
            RunReportGenerator();
        }

        private static void InitializeEnvironment()
        {
            Console.OutputEncoding = Encoding.UTF8;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("uk-UA", false);
            Thread.CurrentThread.CurrentUICulture = new CultureInfo("uk-UA", false);

            var configuration = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .Build();

            DefaultPath = string.IsNullOrWhiteSpace(configuration["DefaultPath"]) ? "Not specified in appsettings.json" : configuration["DefaultPath"];
        }

        private static void RunReportGenerator()
        {
            Console.WriteLine("=======Ukrinform Report Generator=======");
            (DateTime reportStartDate, DateTime reportEndDate) = SelectOperatingWeek();
            string folderPath = SelectFolderLocation();

            Console.WriteLine();
            _ = new ReportGenerator(folderPath, reportStartDate, reportEndDate, ReportEndDay);
        }

        private static (DateTime startDate, DateTime endDate) SelectOperatingWeek()
        {
            DateTime today = DateTime.Today;
            DateTime currentWeekEnd = GetCurrentWeekEndDay(today, ReportEndDay);
            DateTime currentWeekStart = currentWeekEnd.AddDays(-6);
            DateTime previousWeekEnd = currentWeekEnd.AddDays(-7);
            DateTime previousWeekStart = previousWeekEnd.AddDays(-6);

            Console.WriteLine("Select operating week:");
            Console.WriteLine($"1. Current week ({currentWeekStart:dd.MM.yyyy} - {currentWeekEnd:dd.MM.yyyy})");
            Console.WriteLine($"2. Previous week ({previousWeekStart:dd.MM.yyyy} - {previousWeekEnd:dd.MM.yyyy})");

            while (true)
            {
                Console.Write("> ");
                if (int.TryParse(Console.ReadLine(), out int choice))
                {
                    switch (choice)
                    {
                        case 1:
                            return (currentWeekStart, currentWeekEnd);
                        case 2:
                            return (previousWeekStart, previousWeekEnd);
                        default:
                            Console.WriteLine("Invalid input! Try again");
                            break;
                    }
                }
                else
                {
                    Console.WriteLine("Invalid input! Try again");
                }
            }
        }

        private static DateTime GetCurrentWeekEndDay(DateTime start, DayOfWeek endDay)
        {
            int daysUntilEndDay = ((int)endDay - (int)start.DayOfWeek + 7) % 7;
            return start.AddDays(daysUntilEndDay);
        }

        private static string SelectFolderLocation()
        {
            string folderPath;
            bool isValidPath;

            do
            {
                DisplayFolderOptions();
                int choice = GetValidUserChoice();
                folderPath = choice == 1 ? DefaultPath : GetUserSpecifiedPath();
                isValidPath = Directory.Exists(folderPath);

                if (!isValidPath)
                {
                    Console.WriteLine("Invalid directory path! Try again");
                }
            } while (!isValidPath);

            return folderPath;
        }

        private static void DisplayFolderOptions()
        {
            Console.WriteLine("\nSelect folder location:");
            Console.WriteLine($"1. Use default location \n({DefaultPath})");
            Console.WriteLine("2. Manually specify folder path");
        }

        private static int GetValidUserChoice()
        {
            while (true)
            {
                Console.Write("> ");
                if (int.TryParse(Console.ReadLine(), out int choice) && (choice == 1 || choice == 2))
                {
                    return choice;
                }
                Console.WriteLine("Invalid input! Try again");
            }
        }

        private static string GetUserSpecifiedPath()
        {
            Console.WriteLine();
            Console.WriteLine("Enter the folder path:");
            Console.Write("> ");
            return Console.ReadLine();
        }

        private static void ClearLicenseMsg()
        {
            try
            {
                DocX.Load((string)null);
            }
            catch { 
                Console.Clear(); 
            }
        }
    }
}