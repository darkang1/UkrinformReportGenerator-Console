using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace URG_Console
{
    public class ReportGenerator
    {
        private enum Month
        {
            січня = 1,
            лютого = 2,
            березня = 3,
            квітня = 4,
            травня = 5,
            червня = 6,
            липня = 7,
            серпня = 8,
            вересня = 9,
            жовтня = 10,
            листопада = 11,
            грудня = 12
        }

        // Setting variables with default values
        private int _weekStartDay { get; set; } = 0;
        private int _weekEndDay { get; set; } = 0;

        private Month _currMonthEnum { get; set; } = Month.січня;

        private int _currYear { get; set; } = 2000;

        private string _header { get; set; } = "[Header not set]" + Environment.NewLine;

        private int _unsuccessfulConnections { get; set; } = 0;

        private string _folderPath { get; set; } = String.Empty;

        private List<WebParser> _parsedArticles = new List<WebParser>();

        public ReportGenerator(string folderPath)
        {

            _folderPath = folderPath;

            // Setting current date, including week beginnig and endning dates for the report header and filename
            SetCurrDate();
            SetHeader();

            // Adding all parsed articles to an array
            Dictionary<string, string> fileLinks = new Dictionary<string, string>(DocsParser.GetLinks(folderPath));
            _parsedArticles = WebParser.ParseArticles(fileLinks);

            // Sorting parsed articles
            // For some ridiculous, unknown reason sometimes one function sorts properly, sometimes the other
            // So switch sorting function when needed
            SortParsedArticlesByLINQ();
            //sortParsedArticlesByDelegate();

            DisplayParsedArtiles();

            // Generating final Word report with all obtained and processed data
            GenerateMSWordReport();
        }

        private void SetCurrDate()
        {
            Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("uk-UA");

            DateTime todaysDate = DateTime.Today;

            string thisWeekStart = todaysDate.FirstDayOfWeek().ToShortDateString();
            string thisWeekEnd = todaysDate.LastDayOfWeek().ToShortDateString();

            TrimAndSetDate(thisWeekStart, thisWeekEnd);

        }

        private void TrimAndSetDate(string startDay, string endDay)
        {
            // WORKS ONLY IF CULTURE SET TO UKRAINE (uk-UA)!!!!!!!!!
            Thread.CurrentThread.CurrentCulture = new CultureInfo("uk-UA", false);

            try
            {
                // Look for the first '.' occurence
                int startIndex = startDay.IndexOf('.');
                int endIndex = endDay.IndexOf('.');

                // Getting substring from 0 index to the first '.' occurence to get day
                _weekStartDay = int.Parse(startDay.Substring(0, startIndex));
                _weekEndDay = int.Parse(endDay.Substring(0, endIndex));

                // Getting subsrting betwee the first '.' and second '.' occurence to get month 
                int monthIndex = int.Parse(startDay.Substring(startDay.IndexOf('.') + 1, startDay.IndexOf('.', startDay.IndexOf('.'))));
                _currMonthEnum = (Month)monthIndex;

                // Removing substring from the beginning to the second '.' occurence to get the year
                _currYear = int.Parse(endDay.Remove(0, (endDay.IndexOf('.', endDay.IndexOf('.') + 1)) + 1));

            }
            catch (Exception ex)
            {
                Console.WriteLine("[ReportGenerator] " + ex.Message);
                Console.WriteLine("Error occured during setting date. Need to set date manually in the report");
            }

        }

        private void SetHeader()
        {
            try
            {
                if (_weekStartDay < _weekEndDay)
                    _header = "Автор - Ярослав Довгопол" + Environment.NewLine + $"Публікації, що вийшли в період з {_weekStartDay} по {_weekEndDay} {_currMonthEnum} {_currYear} року" + Environment.NewLine;
                else if (_weekStartDay > _weekEndDay)
                {

                    Month nextMonth = (int)_currMonthEnum >= 12 ? Month.січня : (_currMonthEnum + 1);
                    _header = "Автор - Ярослав Довгопол" + Environment.NewLine + $"Публікації, що вийшли в період з {_weekStartDay} {_currMonthEnum} по {_weekEndDay} {nextMonth} {_currYear} року" + Environment.NewLine;
                }
                else
                    throw new ArgumentException("[ReportGenerator] Beginning week day and ending week day cannot be the same!");
            }
            catch (ArgumentException ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine("Need to set header manually in the report");
            }
            catch (Exception ex)
            {
                Console.WriteLine("[ReportGenerator] Unhandled exception occured: ");
                Console.WriteLine(ex.ToString());
                Console.WriteLine("Need to set header manually in the report");
            }


        }


        private void SortParsedArticlesByLINQ()
        {
            // Sorting list using LINQ query
            if (_parsedArticles != null)
            {
                _parsedArticles.OrderBy(article => article.ArticleDate).ThenBy(article => article.ArticleDate);
            }

        }

        private void SortParsedArticlesByDelegate()
        {
            // Sorting list using anonymos method type delegate
            if (_parsedArticles != null)
            {
                _parsedArticles.Sort(delegate (WebParser article1, WebParser article2)
                {
                    return article1.ArticleDate.CompareTo(article2.ArticleDate);
                });
            }

        }


        internal void DisplayParsedArtiles()
        {
            if (_parsedArticles == null || _parsedArticles.Count == 0)
            {
                Console.WriteLine("[ReportGenerator] No parsed articles to display!");
                //Environment.Exit(-100);
            }
            else
            {
                try
                {
                    Console.WriteLine("\n*****Parsed Articles*****\n");
                    for (int i = 0; i < _parsedArticles.Count; i++)
                    {
                        Console.WriteLine("=====================");

                        if (_parsedArticles[i].ArticleHeader == "Unknown")
                            Console.WriteLine($"{i + 1}. " + Path.GetFileName(_parsedArticles[i].ArticleFilePath));
                        else
                            Console.WriteLine($"{i + 1}. {_parsedArticles[i].ArticleHeader}");

                        Console.WriteLine($"Date: {_parsedArticles[i].ArticleDate}");
                        Console.WriteLine($"Article type: {_parsedArticles[i].ArticleType}");
                        Console.WriteLine($"Amount of chars: {_parsedArticles[i].ArticleChars}");
                        Console.WriteLine($"Exclusive: {_parsedArticles[i].ArticleExclusive}");
                        Console.WriteLine($"Link: {_parsedArticles[i].ArticleLink}");
                        Console.WriteLine("=====================\n");

                        if (_parsedArticles[i].ArticleChars <= 0)
                            _unsuccessfulConnections++;

                    }
                    Console.WriteLine($"Total of unsuccessful connections: {_unsuccessfulConnections}");

                }
                catch (Exception ex)
                {
                    Console.WriteLine("[ReportGenerator] " + ex.Message);
                    Console.WriteLine("Skipping...");
                }

            }

        }

        internal void DisplayHeader()
        {
            Console.WriteLine(_header);
        }

        internal void DisplayDate()
        {
            Console.WriteLine($"Current week start day: {_weekStartDay}");
            Console.WriteLine($"Current week end day: {_weekEndDay}");
            Console.WriteLine($"Current month: {_currMonthEnum}");
            Console.WriteLine($"Current year: {_currYear}");
        }

        private int GetCurrentMonthWeekNumber()
        {
            DateTime date = DateTime.Today;
            DateTime firstMonthDay = new DateTime(date.Year, date.Month, 1);
            DateTime firstMonthMonday = firstMonthDay.AddDays((DayOfWeek.Monday + 7 - firstMonthDay.DayOfWeek) % 7);
            // Decided to go with Wednesday as starting point for counting month week number 
            DateTime firstMonthWednesday = firstMonthDay.AddDays((DayOfWeek.Wednesday + 7 - firstMonthDay.DayOfWeek) % 7);

            if (firstMonthWednesday > date)
            {
                firstMonthDay = firstMonthDay.AddMonths(-1);
                firstMonthWednesday = firstMonthDay.AddDays((DayOfWeek.Wednesday + 7 - firstMonthDay.DayOfWeek) % 7);
            }
            int weekNum = (int)Math.Round((date - firstMonthWednesday).Days / 7.0) + 1;
            return weekNum;
        }

        private string ArabicToRoman(int num)
        {
            // Using this function to represent week number as a roman number in report filename, thus range 1-5 should be more than enough
            Dictionary<int, string> romanNums = new Dictionary<int, string>(){
                {1, "I" }, {2, "II"}, {3, "III"}, {4, "IV"}, {5, "V"}};

            string answer = String.Empty;

            try
            {
                if (romanNums.ContainsKey(num))
                {
                    romanNums.TryGetValue(num, out answer);
                    return answer;
                }

                else
                    throw new ArgumentOutOfRangeException("Invalid number was passed to ArabicToRoman number convertor!" + Environment.NewLine + "Acceptable range is [1-5]");
            }
            catch (ArgumentOutOfRangeException ex)
            {
                Console.WriteLine("[ReportGenerator] " + ex.Message);
                return "0";
            }

        }


        private void GenerateMSWordReport()
        {
            // Initializing parsedArticles just in case if forgotted, to not to crash the application on this stage
            if (_parsedArticles == null)
                _parsedArticles = new List<WebParser>();
            if (_parsedArticles.Count == 0)
                Console.WriteLine("\n[ReportGenerator] No parsed articles can be loaded!" + Environment.NewLine + "Generating empty report...");

            // Generating month date for the filename
            string fileMonthDate = (int)_currMonthEnum < 10 ? "0" + ((int)_currMonthEnum).ToString() : ((int)_currMonthEnum).ToString();

            // Creating a proper filename which contains full path
            string fp = Path.GetFullPath(_folderPath);
            string fn = Path.GetFileName($@"\AUTO_Dovgopol_{_currYear}_{fileMonthDate}_{ArabicToRoman(GetCurrentMonthWeekNumber())}={_parsedArticles.Count}.docx");
            string fileName = Path.Combine(fp, fn);

            try
            {
                var doc = DocX.Create(fileName);
                doc.SetDefaultFont(new Xceed.Document.NET.Font("Arial"), 12);
                doc.InsertParagraph(_header);

                //Create Table with 'n' rows and 5 columns (columns are always the same). 
                int rows = _parsedArticles.Count + 1; // Because our first row is taken with column names
                const int cols = 5;
                Table t = doc.AddTable(rows, cols);
                t.Alignment = Alignment.center;

                //Setting Headers
                //To convert to Inches (in) : pt*72 
                t.Rows[0].Height = 0.39 * 72;
                t.Rows[0].Cells[0].Paragraphs.First().Append($"№з/п");
                t.Rows[0].Cells[1].Paragraphs.First().Append($"Дата");
                t.Rows[0].Cells[2].Paragraphs.First().Append($"Заголовок");
                t.Rows[0].Cells[3].Paragraphs.First().Append($"Жанр");
                t.Rows[0].Cells[4].Paragraphs.First().Append($"Кільк. знаків");

                // Creating list of hyperlinks from all found links in the files
                Hyperlink[] hyperlinks = new Hyperlink[_parsedArticles.Count];

                for (int i = 0; i < hyperlinks.Length; i++)
                {
                    if (_parsedArticles[i] != null)
                    {
                        if (_parsedArticles[i].ArticleHeader == "Unknown")
                        {
                            hyperlinks[i] = doc.AddHyperlink("!!! No Link !!!: " + Path.GetFileName(_parsedArticles[i].ArticleFilePath), new Uri(_parsedArticles[i].ArticleLink));
                        }
                        else
                        {
                            hyperlinks[i] = doc.AddHyperlink(_parsedArticles[i].ArticleHeader, new Uri(_parsedArticles[i].ArticleLink));
                        }

                    }

                }

                // Creating a numbering list to have our rows properly number formatted
                var numberedList = doc.AddList(listText: "", listType: ListItemType.Numbered);

                //Fill cells by adding text.  
                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < cols; j++)
                    {

                        // Aligning cells vertically to center left
                        t.Rows[i].Cells[j].VerticalAlignment = VerticalAlignment.Center;

                        // Aligning first and last column text to the center
                        if (j == 0 || j == 4)
                            t.Rows[i].Cells[j].Paragraphs.Last().Alignment = Alignment.center;

                        // Setting needed column width
                        if (j == 0)
                            t.Rows[i].Cells[j].Width = 0.54 * 72;

                        else if (j == 1)
                            t.Rows[i].Cells[j].Width = 1.21 * 72;

                        else if (j == 2)
                            t.Rows[i].Cells[j].Width = 3.25 * 72; //2.97

                        else if (j == 3)
                            t.Rows[i].Cells[j].Width = 1.19 * 72; //1.29

                        else if (j == 4)
                            t.Rows[i].Cells[j].Width = 0.82 * 72;

                        else
                            throw new IndexOutOfRangeException();

                        // Numbering columns
                        if (j == 0 && i > 0)
                        {
                            t.Rows[i].Cells[j].RemoveParagraphAt(0); // Removing space before numbered
                            t.Rows[i].Cells[j].InsertList(numberedList);
                            t.Rows[i].Cells[j].Paragraphs.Last().Alignment = Alignment.center; // Aligning again, since apparently it resets alignment after inserting numbered list
                        }

                        // Adding some text...
                        if (i > 0)
                        {
                            // Adding date
                            if (j == 1)
                            {
                                // From here and now on we are doing 'i-1' since we start on i=1 to insert data, whereas our parsedArticles and hyperlinks begin with i=0
                                if (i - 1 < _parsedArticles.Count)
                                    t.Rows[i].Cells[j].Paragraphs.Last().Append(_parsedArticles[i - 1].ArticleDate);
                            }

                            // Adding title
                            else if (j == 2)
                            {
                                if (i - 1 < hyperlinks.Length) // Checking that amount of links <= amount of rows. If there are more rows, skip them
                                    if (hyperlinks[i - 1] != null)
                                        t.Rows[i].Cells[j].Paragraphs.Last().AppendHyperlink(hyperlinks[i - 1]).
                                                                                     //Setting color
                                                                                     Color(Color.Blue).
                                                                                     // Setting text underline
                                                                                     UnderlineStyle(UnderlineStyle.singleLine);
                            }

                            // Adding article type
                            else if (j == 3)
                            {
                                if (i - 1 < _parsedArticles.Count)
                                    t.Rows[i].Cells[j].Paragraphs.Last().Append(_parsedArticles[i - 1].ArticleType);
                                if (_parsedArticles[i - 1].ArticleType == "Коментар")
                                    t.Rows[i].Cells[j].Paragraphs.Last().Highlight(Highlight.yellow);
                                if (_parsedArticles[i - 1].ArticleExclusive == true)
                                    t.Rows[i].Cells[j].Paragraphs.Last().InsertParagraphAfterSelf("Ексклюзив").Highlight(Highlight.yellow);
                            }

                            // Adding amount of chars
                            else if (j == 4)
                            {
                                if (i - 1 < _parsedArticles.Count)
                                    t.Rows[i].Cells[j].Paragraphs.Last().Append(_parsedArticles[i - 1].ArticleChars.ToString());
                            }
                        }
                    }

                }

                doc.InsertTable(t);
                doc.Save();

            }
            catch (IOException ex)
            {
                KillProcess("WINWORD");
                Console.WriteLine(ex.Message);
                Console.WriteLine("[DocsParser] *** Generated reports cannot be open during the runtime. ***");

            }
            catch (Exception ex)
            {
                Console.WriteLine($"[DocsParser] File: {Path.GetFileName(fileName)}");
                Console.WriteLine("[DocsParser] Unhandled exception occured: ");
                Console.WriteLine(ex.ToString());
            }

            // MS Word can't handle filepath without double-quoting, so adding extra double quotes
            Process.Start(@"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE", '"' + fileName + '"');
        }

        private static void KillProcess(string processName)
        {
            foreach (var process in Process.GetProcessesByName(processName))
            {
                process.Kill();
            }
        }
    }
}
