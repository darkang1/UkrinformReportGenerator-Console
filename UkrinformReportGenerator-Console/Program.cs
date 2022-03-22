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

            ReportGenerator report = new ReportGenerator(@"C:\Users\Neo\Desktop\New folder");


            //var doc = DocX.Load(@"C:\Users\Neo\source\repos\URG_Console\URG_Console\bin\Debug\netcoreapp3.1\NoFiles\US_Biden_G7.docx");

            //Console.WriteLine(doc.Paragraphs[2].Text);
            //int i = 0;
            //foreach(var par in doc.Paragraphs)
            //{
            //    Console.WriteLine($"Paragraph [{i}]: {par.Text}");
            //    i++;
            //}
            //i = 1;


            //while (string.IsNullOrWhiteSpace(doc.Paragraphs[i].Text))
            //    i++;

            //Console.WriteLine("\nFound header: " + doc.Paragraphs[i].Text + "i = " + i);


            //List<int> xx = new List<int>() { 1,2,3};
            //Console.WriteLine(xx[0]);


            //DateTime dt = DateTime.Today;
            //Console.WriteLine(dt.DayOfWeek);
            //var firstdayOfThisWeek = dt.FirstDayOfWeek().ToShortDateString();
            //Console.WriteLine(firstdayOfThisWeek);
            //Console.WriteLine(DateTime.Now.LastDayOfWeek());


            ////Temporary current filepath of the programm
            //string currDir = Directory.GetCurrentDirectory();

            //var wordFiles = DocsParser.getDocsFilePath();

            //foreach (var file in wordFiles)
            //{
            //    Console.WriteLine(file);
            //    Console.WriteLine(" * ********");
            //}

            //var allLinks = DocsParser.parseAllDocsLinks(wordFiles);

            //foreach (var link in allLinks)
            //{
            //    Console.WriteLine(link);
            //}
            //Console.WriteLine($"\nAmount of links: {allLinks.Length}");




            //var lastWeekStart = thisWeekStart.AddDays(-7);
            //var lastWeekEnd = thisWeekStart.AddSeconds(-1);
            //var thisMonthStart = baseDate.AddDays(1 - baseDate.Day);
            //var thisMonthEnd = thisMonthStart.AddMonths(1).AddSeconds(-1);
            //var lastMonthStart = thisMonthStart.AddMonths(-1);
            //var lastMonthEnd = thisMonthStart.AddSeconds(-1);


            //Console.WriteLine(today);
            //Console.WriteLine(thisWeekStart);
            //Console.WriteLine(thisWeekEnd);
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









///////////////////////////////////////////////

//// Testing opening files from parsed directory
//for(int i = 0; i < wordFiles.Count(); i++)
//{
//    Process.Start(@"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE", wordFiles[i]);

//    Console.WriteLine("Next?");
//    Console.ReadKey();

//    killProcess("WINWORD");


//}

///////////////////////////////////////////////