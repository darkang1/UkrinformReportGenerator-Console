using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace URG_Console.Services;

/// <summary>
/// Service for displaying parsed articles to the console.
/// </summary>
public class ArticleDisplayService
{
    /// <summary>
    /// Displays parsed articles to the console with detailed information.
    /// </summary>
    /// <param name="parsedArticles">The list of parsed articles to display.</param>
    public void DisplayParsedArticles(List<WebParser> parsedArticles)
    {
        if (parsedArticles == null || parsedArticles.Count == 0)
        {
            Console.WriteLine($"{Constants.ConsolePrefixArticleDisplayService} No parsed articles to display!");
            return;
        }

        try
        {
            Console.WriteLine("\n*****Parsed Articles*****\n");
            int unsuccessfulConnections = 0;

            for (int i = 0; i < parsedArticles.Count; i++)
            {
                Console.WriteLine("=====================");
                Console.WriteLine(parsedArticles[i].ArticleHeader == Constants.DefaultHeader
                    ? $"{i + 1}. {Path.GetFileName(parsedArticles[i].ArticleFilePath)}"
                    : $"{i + 1}. {parsedArticles[i].ArticleHeader}");
                Console.WriteLine($"Date: {parsedArticles[i].ArticleDate}");
                Console.WriteLine($"Article type: {parsedArticles[i].ArticleType}");
                Console.WriteLine($"Amount of chars: {parsedArticles[i].ArticleChars}");
                Console.WriteLine($"Exclusive: {parsedArticles[i].ArticleExclusive}");

                // Display file name as clickable hyperlink for modern terminals (Windows Terminal, VS Code, etc.)
                string filePath = parsedArticles[i].ArticleFilePath;
                if (!string.IsNullOrEmpty(filePath) && File.Exists(filePath))
                {
                    // Use OSC 8 escape sequence to create clickable hyperlink with custom display text
                    string fileUri = FormatFileUri(filePath);
                    string fileName = parsedArticles[i].ArticleFileName;
                    string clickableFileName = CreateClickableHyperlink(fileUri, fileName);
                    Console.WriteLine($"Article file name: {clickableFileName}");
                }
                else
                {
                    Console.WriteLine($"Article file name: {parsedArticles[i].ArticleFileName}");
                }

                Console.WriteLine($"Link: {parsedArticles[i].ArticleLink}");
                Console.WriteLine("=====================\n");

                if (parsedArticles[i].ArticleChars <= 0)
                    unsuccessfulConnections++;
            }
            Console.WriteLine($"Total of unsuccessful connections: {unsuccessfulConnections}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"{Constants.ConsolePrefixArticleDisplayService} {ex.Message}");
            Console.WriteLine("Skipping...");
        }
    }

    /// <summary>
    /// Sorts parsed articles by their publication date in ascending order.
    /// </summary>
    /// <param name="parsedArticles">The list of articles to sort.</param>
    /// <returns>A new sorted list of articles. Returns an empty list if input is null.</returns>
    public List<WebParser> SortArticlesByDate(List<WebParser> parsedArticles)
    {
        if (parsedArticles == null)
        {
            return new List<WebParser>();
        }

        return parsedArticles.OrderBy(article =>
        {
            DateTime.TryParse(article.ArticleDate, out DateTime dt);
            return dt;
        }).ToList();
    }

    /// <summary>
    /// Formats a file path as a file:// URI for clickable links in modern terminals.
    /// </summary>
    /// <param name="filePath">The absolute file path to format.</param>
    /// <returns>A file:// URI string that can be clicked in modern terminals.</returns>
    private static string FormatFileUri(string filePath)
    {
        try
        {
            // Convert to absolute path if relative
            string absolutePath = Path.GetFullPath(filePath);

            // Create URI - this handles Windows paths correctly
            Uri fileUri = new Uri(absolutePath);
            return fileUri.AbsoluteUri;
        }
        catch
        {
            // Fallback to original path if URI creation fails
            return filePath;
        }
    }

    /// <summary>
    /// Creates a clickable hyperlink using OSC 8 escape sequence for modern terminals.
    /// The display text will be shown, but clicking it will open the URI.
    /// </summary>
    /// <param name="uri">The URI to link to (e.g., file:// URI).</param>
    /// <param name="displayText">The text to display (e.g., filename).</param>
    /// <returns>A string with escape sequences that creates a clickable hyperlink in supported terminals.</returns>
    private static string CreateClickableHyperlink(string uri, string displayText)
    {
        // OSC 8 escape sequence format: \x1b]8;;<uri>\x1b\\<text>\x1b]8;;\x1b\\
        // This creates a hyperlink where only the display text is shown, but clicking opens the URI
        // Supported by: Windows Terminal, VS Code terminal, iTerm2, and other modern terminals
        return $"\x1b]8;;{uri}\x1b\\{displayText}\x1b]8;;\x1b\\";
    }
}

