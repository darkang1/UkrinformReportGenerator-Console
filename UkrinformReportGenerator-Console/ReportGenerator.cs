using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace URG_Console
{
    public class ReportGenerator
    {
        #region Enums and Properties
        private enum Month
        {
            січня = 1, лютого = 2, березня = 3, квітня = 4, травня = 5, червня = 6,
            липня = 7, серпня = 8, вересня = 9, жовтня = 10, листопада = 11, грудня = 12
        }

        private string FolderPath { get; }
        private DateTime ReportStartDate { get; }
        private DateTime ReportEndDate { get; }
        private DayOfWeek ReportEndDay { get; }
        private List<WebParser> ParsedArticles { get; set; } = new List<WebParser>();
        private string Header { get; set; } = "[Header not set]" + Environment.NewLine;
        private int UnsuccessfulConnections { get; set; } = 0;

        private int WeekStartDay { get; set; }
        private int WeekEndDay { get; set; }
        private Month CurrentMonthEnum { get; set; }
        private int CurrentYear { get; set; }
        private bool IsReportSpanningTwoMonths { get; set; }
        private bool IsReportSpanningTwoYears { get; set; }
        private Month NextMonth { get; set; }
        private int NextYear { get; set; }
        #endregion

        #region Constructor and Initialization
        public ReportGenerator(string folderPath, DateTime reportStartDate, DateTime reportEndDate, DayOfWeek reportEndDay)
        {
            FolderPath = folderPath;
            ReportStartDate = reportStartDate;
            ReportEndDate = reportEndDate;
            ReportEndDay = reportEndDay;

            InitializeReportData();
            ProcessArticles();
            GenerateReports();
        }

        private void InitializeReportData()
        {
            SetCurrentDate();
            SetHeader();
        }

        private void ProcessArticles()
        {
            List<ArticleSource> articleSources = DocsParser.ParseDocumentsInDirectory(FolderPath);
            ParsedArticles = WebParser.ParseArticles(articleSources);
            SortParsedArticles();
            DisplayParsedArticles();
        }

        private void GenerateReports()
        {
            GenerateMSWordReport();
            //GenerateMSWordReport_Simplified();
        }
        #endregion

        #region Date and Header Management
        private void SetCurrentDate()
        {
            WeekStartDay = ReportStartDate.Day;
            WeekEndDay = ReportEndDate.Day;
            CurrentMonthEnum = (Month)ReportStartDate.Month;
            CurrentYear = ReportStartDate.Year;

            IsReportSpanningTwoMonths = ReportStartDate.Month != ReportEndDate.Month;
            IsReportSpanningTwoYears = ReportStartDate.Year != ReportEndDate.Year;
            NextMonth = IsReportSpanningTwoMonths ? (Month)ReportEndDate.Month : CurrentMonthEnum;
            NextYear = IsReportSpanningTwoYears ? ReportEndDate.Year : CurrentYear;
        }

        private void SetHeader()
        {
            try
            {
                string dateRange;
                if (IsReportSpanningTwoYears)
                {
                    dateRange = $"з {WeekStartDay} {CurrentMonthEnum} {CurrentYear} року по {WeekEndDay} {NextMonth} {NextYear} року";
                }
                else if (IsReportSpanningTwoMonths)
                {
                    dateRange = $"з {WeekStartDay} {CurrentMonthEnum} по {WeekEndDay} {NextMonth} {CurrentYear} року";
                }
                else
                {
                    dateRange = $"з {WeekStartDay} по {WeekEndDay} {CurrentMonthEnum} {CurrentYear} року";
                }

                Header = $"Автор - Ярослав Довгопол{Environment.NewLine}Публікації, що вийшли в період {dateRange}{Environment.NewLine}";
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ReportGenerator] Error setting header: {ex.Message}");
                Console.WriteLine("Need to set header manually in the report");
                Header = "[Header not set]" + Environment.NewLine;
            }
        }
        #endregion

        #region Article Processing and Sorting
        private void SortParsedArticles()
        {
            ParsedArticles = ParsedArticles?.OrderBy(article =>
            {
                DateTime.TryParse(article.ArticleDate, out DateTime dt);
                return dt;
            }).ToList();
        }
        #endregion

        #region Display Methods
        public void DisplayParsedArticles()
        {
            if (ParsedArticles == null || ParsedArticles.Count == 0)
            {
                Console.WriteLine("[ReportGenerator] No parsed articles to display!");
                return;
            }

            try
            {
                Console.WriteLine("\n*****Parsed Articles*****\n");
                for (int i = 0; i < ParsedArticles.Count; i++)
                {
                    Console.WriteLine("=====================");
                    Console.WriteLine(ParsedArticles[i].ArticleHeader == "Unknown"
                        ? $"{i + 1}. {Path.GetFileName(ParsedArticles[i].ArticleFilePath)}"
                        : $"{i + 1}. {ParsedArticles[i].ArticleHeader}");
                    Console.WriteLine($"Date: {ParsedArticles[i].ArticleDate}");
                    Console.WriteLine($"Article type: {ParsedArticles[i].ArticleType}");
                    Console.WriteLine($"Amount of chars: {ParsedArticles[i].ArticleChars}");
                    Console.WriteLine($"Exclusive: {ParsedArticles[i].ArticleExclusive}");
                    Console.WriteLine($"Link: {ParsedArticles[i].ArticleLink}");
                    Console.WriteLine("=====================\n");

                    if (ParsedArticles[i].ArticleChars <= 0)
                        UnsuccessfulConnections++;
                }
                Console.WriteLine($"Total of unsuccessful connections: {UnsuccessfulConnections}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("[ReportGenerator] " + ex.Message);
                Console.WriteLine("Skipping...");
            }
        }
        #endregion

        #region Utility Methods
        private int GetCurrentWeekNumber()
        {
            DateTime firstDayOfMonth = new(ReportEndDate.Year, ReportEndDate.Month, 1);
            DateTime firstReportDayOfMonth = GetFirstReportDayOfMonth(firstDayOfMonth);

            if (firstReportDayOfMonth > ReportEndDate)
            {
                return 1; // This is effectively the first week of the new month
            }

            int weekNum = (int)Math.Ceiling((ReportEndDate - firstReportDayOfMonth).TotalDays / 7) + 1;
            return Math.Max(1, Math.Min(5, weekNum));
        }

        private DateTime GetFirstReportDayOfMonth(DateTime firstDayOfMonth)
        {
            return firstDayOfMonth.AddDays((7 + (int)ReportEndDay - (int)firstDayOfMonth.DayOfWeek) % 7);
        }

        private static string ArabicToRoman(int num)
        {
            if (num < 1 || num > 5)
            {
                Console.WriteLine($"[ReportGenerator] Warning: Invalid number {num} passed to ArabicToRoman. Using 'I' as default.");
                return "I";
            }

            string[] romanNumerals = { "I", "II", "III", "IV", "V" };
            return romanNumerals[num - 1];
        }

        private static void KillProcess(string processName)
        {
            foreach (var process in Process.GetProcessesByName(processName))
            {
                process.Kill();
            }
        }
        #endregion

        #region Report Generation
        private void GenerateMSWordReport()
        {
            ParsedArticles ??= new List<WebParser>();
            if (ParsedArticles.Count == 0)
                Console.WriteLine("\n[ReportGenerator] No parsed articles can be loaded!" + Environment.NewLine + "Generating empty report...");

            string fileMonthDate = (int)CurrentMonthEnum < 10 ? "0" + ((int)CurrentMonthEnum).ToString() : ((int)CurrentMonthEnum).ToString();
            string folderPath = Path.GetFullPath(FolderPath);
            string fileName = Path.GetFileName($@"\AUTO_Dovgopol_{CurrentYear}_{fileMonthDate}_{ArabicToRoman(GetCurrentWeekNumber())}={ParsedArticles.Count}.docx");
            string reportFilePath = Path.Combine(folderPath, fileName);

            try
            {
                var doc = DocX.Create(reportFilePath);
                doc.SetDefaultFont(new Xceed.Document.NET.Font("Arial"), 12);
                doc.InsertParagraph(Header);

                int rows = ParsedArticles.Count + 1;
                const int cols = 5;
                Table t = doc.AddTable(rows, cols);
                t.Alignment = Alignment.center;

                SetupTableHeaders(t);
                Hyperlink[] hyperlinks = CreateHyperlinks(doc);
                var numberedList = doc.AddList(listText: "", listType: ListItemType.Numbered);

                FillTableWithData(t, hyperlinks, numberedList);

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
                Console.WriteLine($"[DocsParser] File: {Path.GetFileName(reportFilePath)}");
                Console.WriteLine("[DocsParser] Unhandled exception occurred: ");
                Console.WriteLine(ex.ToString());
            }

            Process.Start(@"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE", '"' + reportFilePath + '"');
        }

        private static void SetupTableHeaders(Table t)
        {
            t.Rows[0].Height = 0.39 * 72;
            t.Rows[0].Cells[0].Paragraphs.First().Append("№з/п");
            t.Rows[0].Cells[1].Paragraphs.First().Append("Дата");
            t.Rows[0].Cells[2].Paragraphs.First().Append("Заголовок");
            t.Rows[0].Cells[3].Paragraphs.First().Append("Жанр");
            t.Rows[0].Cells[4].Paragraphs.First().Append("Кільк. знаків");
        }

        private Hyperlink[] CreateHyperlinks(DocX doc)
        {
            Hyperlink[] hyperlinks = new Hyperlink[ParsedArticles.Count];
            for (int i = 0; i < hyperlinks.Length; i++)
            {
                if (ParsedArticles[i] != null)
                {
                    try
                    {
                        string linkText = ParsedArticles[i].ArticleHeader == "Unknown"
                            ? Path.GetFileName(ParsedArticles[i].ArticleFilePath)
                            : ParsedArticles[i].ArticleHeader;

                        if (string.IsNullOrWhiteSpace(ParsedArticles[i].ArticleLink))
                        {
                            // Create a file URI for the local file
                            string fullPath = Path.GetFullPath(ParsedArticles[i].ArticleFilePath);
                            Uri fileUri = new(fullPath);
                            hyperlinks[i] = doc.AddHyperlink("!!! NO LINK FOUND !!! " + linkText + " (Click To Open)", fileUri);
                        }
                        else
                        {
                            // Create hyperlink with valid URL
                            hyperlinks[i] = doc.AddHyperlink(linkText, new Uri(ParsedArticles[i].ArticleLink));
                        }
                    }
                    catch (UriFormatException ex)
                    {
                        Console.WriteLine($"[ReportGenerator] Invalid URI for article {i + 1}: {ex.Message}");
                        // Attempt to create a file URI as a fallback
                        try
                        {
                            string fullPath = Path.GetFullPath(ParsedArticles[i].ArticleFilePath);
                            Uri fileUri = new(fullPath);
                            hyperlinks[i] = doc.AddHyperlink($"Local File for article {i + 1}", fileUri);
                        }
                        catch
                        {
                            hyperlinks[i] = doc.AddHyperlink($"Invalid URL for article {i + 1}", new Uri("about:blank"));
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[ReportGenerator] Unexpected error creating hyperlink for article {i + 1}: {ex.Message}");
                        hyperlinks[i] = doc.AddHyperlink($"Error in article {i + 1}", new Uri("about:blank"));
                    }
                }
            }
            return hyperlinks;
        }

        private void FillTableWithData(Table t, Hyperlink[] hyperlinks, List numberedList)
        {
            for (int i = 0; i < t.RowCount; i++)
            {
                for (int j = 0; j < t.ColumnCount; j++)
                {
                    SetupTableCell(t, i, j, numberedList);
                    if (i > 0)
                    {
                        FillTableCell(t, i, j, hyperlinks);
                    }
                }
            }
        }

        private static void SetupTableCell(Table t, int i, int j, List numberedList)
        {
            t.Rows[i].Cells[j].VerticalAlignment = VerticalAlignment.Center;

            if (j == 0 || j == 4)
                t.Rows[i].Cells[j].Paragraphs.Last().Alignment = Alignment.center;

            SetCellWidth(t.Rows[i].Cells[j], j);

            if (j == 0 && i > 0)
            {
                t.Rows[i].Cells[j].RemoveParagraphAt(0);
                t.Rows[i].Cells[j].InsertList(numberedList);
                t.Rows[i].Cells[j].Paragraphs.Last().Alignment = Alignment.center;
            }
        }

        private static void SetCellWidth(Cell cell, int columnIndex)
        {
            switch (columnIndex)
            {
                case 0: cell.Width = 0.54 * 72; break;
                case 1: cell.Width = 1.21 * 72; break;
                case 2: cell.Width = 3.25 * 72; break;
                case 3: cell.Width = 1.19 * 72; break;
                case 4: cell.Width = 0.82 * 72; break;
                default: throw new IndexOutOfRangeException();
            }
        }

        private void FillTableCell(Table t, int i, int j, Hyperlink[] hyperlinks)
        {
            int articleIndex = i - 1;
            switch (j)
            {
                case 1:
                    if (articleIndex < ParsedArticles.Count)
                        t.Rows[i].Cells[j].Paragraphs.Last().Append(ParsedArticles[articleIndex].ArticleDate);
                    break;
                case 2:
                    if (articleIndex < hyperlinks.Length && hyperlinks[articleIndex] != null)
                    {
                        AddHyperlinkToCell(t.Rows[i].Cells[j], hyperlinks[articleIndex]);
                    }
                    break;
                case 3:
                    if (articleIndex < ParsedArticles.Count)
                    {
                        AddArticleTypeToCell(t.Rows[i].Cells[j], ParsedArticles[articleIndex]);
                    }
                    break;
                case 4:
                    if (articleIndex < ParsedArticles.Count)
                        t.Rows[i].Cells[j].Paragraphs.Last().Append(ParsedArticles[articleIndex].ArticleChars.ToString());
                    break;
            }
        }

        private static void AddHyperlinkToCell(Cell cell, Hyperlink hyperlink)
        {
            cell.Paragraphs.Last().AppendHyperlink(hyperlink)
                .Color(Color.Blue)
                .UnderlineStyle(UnderlineStyle.singleLine);
        }

        private static void AddArticleTypeToCell(Cell cell, WebParser article)
        {
            cell.Paragraphs.Last().Append(article.ArticleType);
            if (article.ArticleType == "Коментар")
                cell.Paragraphs.Last().Highlight(Highlight.yellow);
            if (article.ArticleExclusive)
                cell.Paragraphs.Last().InsertParagraphAfterSelf("Ексклюзив").Highlight(Highlight.yellow);
        }
        #endregion

        #region Simplified Report Generation (Deprecated)
        private void GenerateMSWordReport_Simplified()
        {
            ParsedArticles ??= new List<WebParser>();
            if (ParsedArticles.Count == 0)
                Console.WriteLine("\n[ReportGenerator] No parsed articles can be loaded!" + Environment.NewLine + "Generating empty report...");

            string fileMonthDate = (int)CurrentMonthEnum < 10 ? "0" + ((int)CurrentMonthEnum).ToString() : ((int)CurrentMonthEnum).ToString();
            string fp = Path.GetFullPath(FolderPath);
            string fn = Path.GetFileName($@"\AUTO_Dovgopol_{CurrentYear}_{fileMonthDate}_{ArabicToRoman(GetCurrentWeekNumber())}={ParsedArticles.Count}_Simplified.docx");
            string fileName = Path.Combine(fp, fn);

            try
            {
                var doc = DocX.Create(fileName);
                doc.SetDefaultFont(new Xceed.Document.NET.Font("Arial"), 12);
                doc.InsertParagraph(Header);

                int rows = ParsedArticles.Count + 1;
                const int cols = 2;
                Table t = doc.AddTable(rows, cols);
                t.Alignment = Alignment.center;

                SetupSimplifiedTableHeaders(t);
                Hyperlink[] hyperlinks = CreateHyperlinks(doc);
                var numberedList = doc.AddList(listText: "", listType: ListItemType.Numbered);

                FillSimplifiedTableWithData(t, hyperlinks, numberedList);

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
                Console.WriteLine("[DocsParser] Unhandled exception occurred: ");
                Console.WriteLine(ex.ToString());
            }

            Process.Start(@"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE", '"' + fileName + '"');
        }

        private static void SetupSimplifiedTableHeaders(Table t)
        {
            t.Rows[0].Height = 0.39 * 72;
            t.Rows[0].Cells[0].Paragraphs.First().Append("№з/п");
            t.Rows[0].Cells[1].Paragraphs.First().Append("Заголовок");
        }

        private void FillSimplifiedTableWithData(Table t, Hyperlink[] hyperlinks, List numberedList)
        {
            for (int i = 0; i < t.RowCount; i++)
            {
                for (int j = 0; j < t.ColumnCount; j++)
                {
                    SetupSimplifiedTableCell(t, i, j, numberedList);
                    if (i > 0)
                    {
                        FillSimplifiedTableCell(t, i, j, hyperlinks);
                    }
                }
            }
        }

        private static void SetupSimplifiedTableCell(Table t, int i, int j, List numberedList)
        {
            t.Rows[i].Cells[j].VerticalAlignment = VerticalAlignment.Center;

            if (j == 0)
            {
                t.Rows[i].Cells[j].Paragraphs.Last().Alignment = Alignment.center;
                t.Rows[i].Cells[j].Width = 0.54 * 72;

                if (i > 0)
                {
                    t.Rows[i].Cells[j].RemoveParagraphAt(0);
                    t.Rows[i].Cells[j].InsertList(numberedList);
                    t.Rows[i].Cells[j].Paragraphs.Last().Alignment = Alignment.center;
                }
            }
            else if (j == 1)
            {
                t.Rows[i].Cells[j].Width = 7.96 * 72;
            }
        }

        private void FillSimplifiedTableCell(Table t, int i, int j, Hyperlink[] hyperlinks)
        {
            if (j == 1)
            {
                int articleIndex = i - 1;
                if (articleIndex < hyperlinks.Length && hyperlinks[articleIndex] != null)
                {
                    AddSimplifiedHyperlinkToCell(t.Rows[i].Cells[j], hyperlinks[articleIndex], ParsedArticles[articleIndex]);
                }
            }
        }

        private static void AddSimplifiedHyperlinkToCell(Cell cell, Hyperlink hyperlink, WebParser article)
        {
            string originalHyperlinkText = hyperlink.Text;
            cell.Paragraphs.Last().Append(originalHyperlinkText);

            Hyperlink extendedHyperLink = hyperlink;
            extendedHyperLink.Text = hyperlink.Uri.ToString();

            cell.Paragraphs.Last().AppendLine().AppendHyperlink(extendedHyperLink)
                .Color(Color.Blue)
                .UnderlineStyle(UnderlineStyle.singleLine);

            if (article.ArticleType == "Коментар")
                cell.Paragraphs.Last().Append(" " + article.ArticleType.ToUpper()).Bold();
        }
        #endregion
    }
}