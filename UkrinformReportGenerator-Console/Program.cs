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
