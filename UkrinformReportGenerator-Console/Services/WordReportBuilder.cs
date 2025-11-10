using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using URG_Console.Configuration;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace URG_Console.Services;

/// <summary>
/// Service for building Word document reports.
/// </summary>
public class WordReportBuilder
{
    private readonly AppSettings _appSettings;

    /// <summary>
    /// Initializes a new instance of the <see cref="WordReportBuilder"/> class.
    /// </summary>
    /// <param name="appSettings">The application settings containing configuration.</param>
    /// <exception cref="ArgumentNullException">Thrown when appSettings is null.</exception>
    public WordReportBuilder(AppSettings appSettings)
    {
        _appSettings = appSettings ?? throw new ArgumentNullException(nameof(appSettings));
    }

    /// <summary>
    /// Generates a Word document report with the provided articles and metadata.
    /// </summary>
    /// <param name="folderPath">The folder path where the report will be saved.</param>
    /// <param name="parsedArticles">The list of parsed articles to include in the report.</param>
    /// <param name="header">The header text to include at the top of the report.</param>
    /// <param name="dateInfo">The date information for file naming.</param>
    /// <returns>A task that represents the asynchronous operation.</returns>
    public async Task GenerateReportAsync(
        string folderPath,
        List<WebParser> parsedArticles,
        string header,
        ReportDateInfo dateInfo)
    {
        parsedArticles ??= new List<WebParser>();
        if (parsedArticles.Count == 0)
            Console.WriteLine($"\n{Constants.ConsolePrefixWordReportBuilder} No parsed articles can be loaded!" + Environment.NewLine + "Generating empty report...");

        string fileMonthDate = dateInfo.FileMonth < 10 ? "0" + dateInfo.FileMonth.ToString() : dateInfo.FileMonth.ToString();
        string fullFolderPath = await Task.Run(() => Path.GetFullPath(folderPath));
        string romanWeek = ArabicToRoman(dateInfo.WeekNumber);
        string fileName = $"AUTO_Dovgopol_{dateInfo.FileYear}_{fileMonthDate}_{romanWeek}={parsedArticles.Count}.docx";
        string reportFilePath = Path.Combine(fullFolderPath, fileName);
        string wordExecutablePath = _appSettings.WordExecutablePath;

        try
        {
            await Task.Run(() =>
            {
                var doc = DocX.Create(reportFilePath);
                doc.SetDefaultFont(new Xceed.Document.NET.Font("Arial"), 12);
                doc.InsertParagraph(header);

                int rows = parsedArticles.Count + 1;
                const int cols = 5;
                Table t = doc.AddTable(rows, cols);
                t.Alignment = Alignment.center;

                SetupTableHeaders(t);
                Hyperlink[] hyperlinks = CreateHyperlinks(doc, parsedArticles);
                var numberedList = doc.AddList(listText: "", listType: ListItemType.Numbered);

                FillTableWithData(t, hyperlinks, numberedList, parsedArticles);

                doc.InsertTable(t);
                doc.Save();
            });
        }
        catch (IOException ex)
        {
            KillProcess(Constants.WordProcessName);
            Console.WriteLine(ex.Message);
            Console.WriteLine($"{Constants.ConsolePrefixWordReportBuilder} *** Generated reports cannot be open during the runtime. ***");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"{Constants.ConsolePrefixWordReportBuilder} File: {Path.GetFileName(reportFilePath)}");
            Console.WriteLine($"{Constants.ConsolePrefixWordReportBuilder} Unhandled exception occurred: ");
            Console.WriteLine(ex.ToString());
        }

        bool fileExists = await Task.Run(() => !string.IsNullOrWhiteSpace(wordExecutablePath) && File.Exists(wordExecutablePath));
        if (fileExists)
        {
            await Task.Run(() => Process.Start(wordExecutablePath, '"' + reportFilePath + '"'));
        }
        else
        {
            Console.WriteLine($"{Constants.ConsolePrefixWordReportBuilder} Warning: Word executable not found at '{wordExecutablePath}'. Please open the report manually.");
        }
    }

    /// <summary>
    /// Sets up the table headers with column names.
    /// </summary>
    /// <param name="t">The table to set up headers for.</param>
    private static void SetupTableHeaders(Table t)
    {
        t.Rows[0].Height = Constants.TableHeaderRowHeight;
        t.Rows[0].Cells[0].Paragraphs.First().Append("№з/п");
        t.Rows[0].Cells[1].Paragraphs.First().Append("Дата");
        t.Rows[0].Cells[2].Paragraphs.First().Append("Заголовок");
        t.Rows[0].Cells[3].Paragraphs.First().Append("Жанр");
        t.Rows[0].Cells[4].Paragraphs.First().Append("Кільк. знаків");
    }

    /// <summary>
    /// Creates hyperlinks for all articles in the document.
    /// </summary>
    /// <param name="doc">The Word document to add hyperlinks to.</param>
    /// <param name="parsedArticles">The list of parsed articles.</param>
    /// <returns>An array of hyperlinks created for the articles.</returns>
    private static Hyperlink[] CreateHyperlinks(DocX doc, List<WebParser> parsedArticles)
    {
        Hyperlink[] hyperlinks = new Hyperlink[parsedArticles.Count];
        for (int i = 0; i < hyperlinks.Length; i++)
        {
            if (parsedArticles[i] != null)
            {
                try
                {
                    string linkText = parsedArticles[i].ArticleHeader == Constants.DefaultHeader
                        ? Path.GetFileName(parsedArticles[i].ArticleFilePath)
                        : parsedArticles[i].ArticleHeader;

                    if (string.IsNullOrWhiteSpace(parsedArticles[i].ArticleLink))
                    {
                        string fullPath = Path.GetFullPath(parsedArticles[i].ArticleFilePath);
                        Uri fileUri = new(fullPath);
                        hyperlinks[i] = doc.AddHyperlink("!!! NO LINK FOUND !!! " + linkText + " (Click To Open)", fileUri);
                    }
                    else
                    {
                        hyperlinks[i] = doc.AddHyperlink(linkText, new Uri(parsedArticles[i].ArticleLink));
                    }
                }
                catch (UriFormatException ex)
                {
                    Console.WriteLine($"{Constants.ConsolePrefixWordReportBuilder} Invalid URI for article {i + 1}: {ex.Message}");
                    try
                    {
                        string fullPath = Path.GetFullPath(parsedArticles[i].ArticleFilePath);
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
                    Console.WriteLine($"{Constants.ConsolePrefixWordReportBuilder} Unexpected error creating hyperlink for article {i + 1}: {ex.Message}");
                    hyperlinks[i] = doc.AddHyperlink($"Error in article {i + 1}", new Uri("about:blank"));
                }
            }
        }
        return hyperlinks;
    }

    /// <summary>
    /// Fills the table with article data.
    /// </summary>
    /// <param name="t">The table to fill.</param>
    /// <param name="hyperlinks">The array of hyperlinks for articles.</param>
    /// <param name="numberedList">The numbered list for row numbering.</param>
    /// <param name="parsedArticles">The list of parsed articles.</param>
    private static void FillTableWithData(Table t, Hyperlink[] hyperlinks, List numberedList, List<WebParser> parsedArticles)
    {
        for (int i = 0; i < t.RowCount; i++)
        {
            for (int j = 0; j < t.ColumnCount; j++)
            {
                SetupTableCell(t, i, j, numberedList);
                if (i > 0)
                {
                    FillTableCell(t, i, j, hyperlinks, parsedArticles);
                }
            }
        }
    }

    /// <summary>
    /// Sets up a table cell with proper alignment and formatting.
    /// </summary>
    /// <param name="t">The table containing the cell.</param>
    /// <param name="i">The row index.</param>
    /// <param name="j">The column index.</param>
    /// <param name="numberedList">The numbered list for the first column.</param>
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

    /// <summary>
    /// Sets the width of a table cell based on its column index.
    /// </summary>
    /// <param name="cell">The cell to set the width for.</param>
    /// <param name="columnIndex">The column index (0-4).</param>
    /// <exception cref="IndexOutOfRangeException">Thrown when columnIndex is out of range.</exception>
    private static void SetCellWidth(Cell cell, int columnIndex)
    {
        cell.Width = columnIndex switch
        {
            0 => Constants.TableCellWidthColumn0,
            1 => Constants.TableCellWidthColumn1,
            2 => Constants.TableCellWidthColumn2,
            3 => Constants.TableCellWidthColumn3,
            4 => Constants.TableCellWidthColumn4,
            _ => throw new IndexOutOfRangeException($"Invalid column index: {columnIndex}")
        };
    }

    /// <summary>
    /// Fills a specific table cell with article data.
    /// </summary>
    /// <param name="t">The table containing the cell.</param>
    /// <param name="i">The row index.</param>
    /// <param name="j">The column index.</param>
    /// <param name="hyperlinks">The array of hyperlinks.</param>
    /// <param name="parsedArticles">The list of parsed articles.</param>
    private static void FillTableCell(Table t, int i, int j, Hyperlink[] hyperlinks, List<WebParser> parsedArticles)
    {
        int articleIndex = i - 1;
        switch (j)
        {
            case 1:
                if (articleIndex < parsedArticles.Count)
                    t.Rows[i].Cells[j].Paragraphs.Last().Append(parsedArticles[articleIndex].ArticleDate);
                break;
            case 2:
                if (articleIndex < hyperlinks.Length && hyperlinks[articleIndex] != null)
                {
                    AddHyperlinkToCell(t.Rows[i].Cells[j], hyperlinks[articleIndex]);
                }
                break;
            case 3:
                if (articleIndex < parsedArticles.Count)
                {
                    AddArticleTypeToCell(t.Rows[i].Cells[j], parsedArticles[articleIndex]);
                }
                break;
            case 4:
                if (articleIndex < parsedArticles.Count)
                    t.Rows[i].Cells[j].Paragraphs.Last().Append(parsedArticles[articleIndex].ArticleChars.ToString());
                break;
        }
    }

    /// <summary>
    /// Adds a hyperlink to a table cell with blue color and underline styling.
    /// </summary>
    /// <param name="cell">The cell to add the hyperlink to.</param>
    /// <param name="hyperlink">The hyperlink to add.</param>
    private static void AddHyperlinkToCell(Cell cell, Hyperlink hyperlink)
    {
        cell.Paragraphs.Last().AppendHyperlink(hyperlink)
            .Color(Xceed.Drawing.Color.Blue)
            .UnderlineStyle(UnderlineStyle.singleLine);
    }

    /// <summary>
    /// Adds article type information to a cell, with highlighting for comments and exclusive articles.
    /// </summary>
    /// <param name="cell">The cell to add the article type to.</param>
    /// <param name="article">The article containing type information.</param>
    private static void AddArticleTypeToCell(Cell cell, WebParser article)
    {
        cell.Paragraphs.Last().Append(article.ArticleType);
        if (article.ArticleType == Constants.ArticleTypeComment)
            cell.Paragraphs.Last().Highlight(Highlight.yellow);
        if (article.ArticleExclusive)
            cell.Paragraphs.Last().InsertParagraphAfterSelf(Constants.ArticleExclusivePrefix).Highlight(Highlight.yellow);
    }

    /// <summary>
    /// Converts an Arabic number (1-5) to its Roman numeral representation.
    /// </summary>
    /// <param name="num">The number to convert (must be between 1 and 5).</param>
    /// <returns>The Roman numeral string (I, II, III, IV, or V).</returns>
    private static string ArabicToRoman(int num)
    {
        if (num < Constants.MinWeekNumber || num > Constants.MaxWeekNumber)
        {
            Console.WriteLine($"{Constants.ConsolePrefixWordReportBuilder} Warning: Invalid number {num} passed to ArabicToRoman. Using 'I' as default.");
            return Constants.RomanNumerals[0];
        }

        return Constants.RomanNumerals[num - 1];
    }

    /// <summary>
    /// Kills all processes with the specified name.
    /// </summary>
    /// <param name="processName">The name of the process to kill (without .exe extension).</param>
    private static void KillProcess(string processName)
    {
        foreach (var process in Process.GetProcessesByName(processName))
        {
            process.Kill();
        }
    }
}

