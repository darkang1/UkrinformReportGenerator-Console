using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.XPath;
using URG_Console.Interfaces;

namespace URG_Console.Services;

/// <summary>
/// Service for parsing articles from web sources.
/// </summary>
public class ArticleParser : IArticleParser
{
    /// <inheritdoc/>
    public async Task<List<WebParser>> ParseArticlesAsync(List<ArticleSource> articleSources)
    {
        SetCultureToUkrainian();
        ValidateArticleSources(articleSources);

        var articles = new List<WebParser>();
        int processedCount = 0;
        object lockObject = new object();

        // Process articles in parallel for better performance
        var tasks = articleSources.Select(async (articleSource, index) =>
        {
            try
            {
                var article = await ParseSingleArticleAsync(articleSource);
                lock (lockObject)
                {
                    articles.Add(article);
                    processedCount++;
                    Console.WriteLine($"Total processed articles: [{processedCount}/{articleSources.Count}]");
                }
            }
            catch (Exception ex)
            {
                HandleParsingException(ex, articleSource);
                lock (lockObject)
                {
                    articles.Add(CreateErrorArticle(articleSource));
                    processedCount++;
                }
            }
        });

        await Task.WhenAll(tasks);

        return articles;
    }

    /// <summary>
    /// Sets the current thread culture to Ukrainian.
    /// </summary>
    private static void SetCultureToUkrainian()
    {
        Thread.CurrentThread.CurrentCulture = new CultureInfo("uk-UA");
    }

    /// <summary>
    /// Validates that article sources are provided.
    /// </summary>
    /// <param name="articleSources">The list of article sources to validate.</param>
    private static void ValidateArticleSources(List<ArticleSource> articleSources)
    {
        if (articleSources == null || articleSources.Count == 0)
            Console.WriteLine($"\n{Constants.ConsolePrefixArticleParser} No article sources were passed to the parser!");
        else
            Console.WriteLine($"\n{Constants.ConsolePrefixArticleParser} Article sources loaded successfully!");
    }

    /// <summary>
    /// Parses a single article from the provided article source.
    /// </summary>
    /// <param name="articleSource">The article source containing the link and file path.</param>
    /// <returns>A task that represents the asynchronous operation. The task result contains the parsed article.</returns>
    private static async Task<WebParser> ParseSingleArticleAsync(ArticleSource articleSource)
    {
        if (!articleSource.HasLink)
        {
            Console.WriteLine($"{Constants.ConsolePrefixArticleParser} No link found for file: {Path.GetFileName(articleSource.FilePath)}");
            return CreateErrorArticle(articleSource);
        }

        return await Task.Run(() =>
        {
            var web = new HtmlWeb();
            try
            {
                HtmlDocument doc = web.Load(articleSource.Link);

                string newsTitle = ExtractNewsTitle(doc);
                string publishDate = ExtractPublishDate(doc);
                string newsText = ExtractNewsText(doc);
                bool newsExclusive = IsNewsExclusive(doc);

                string fixedDate = RemoveTimestampFromDate(publishDate);

                string finalText = CombineArticleText(newsTitle, fixedDate, newsText);
                int textCharsAmount = CountNonWhiteSpaceChars(finalText);
                string newsType = DetermineArticleType(textCharsAmount);

                return new WebParser(fixedDate, newsTitle, newsType, textCharsAmount, articleSource.Link, newsExclusive, articleSource.FilePath);
            }
            catch (WebException)
            {
                Console.WriteLine($"{Constants.ConsolePrefixArticleParser} Unable to load page for file: {Path.GetFileName(articleSource.FilePath)}");
                return CreateErrorArticle(articleSource);
            }
            catch (XPathException ex)
            {
                Console.WriteLine($"{Constants.ConsolePrefixArticleParser} Error extracting data from: {articleSource.Link}. Error: {ex.Message}");
                return CreateErrorArticle(articleSource);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"{Constants.ConsolePrefixArticleParser} Unexpected error processing file: {Path.GetFileName(articleSource.FilePath)}. Error: {ex.Message}");
                return CreateErrorArticle(articleSource);
            }
        });
    }

    /// <summary>
    /// Extracts the news title from the HTML document.
    /// </summary>
    /// <param name="doc">The HTML document to extract from.</param>
    /// <returns>The extracted and cleaned news title.</returns>
    /// <exception cref="XPathException">Thrown when the title node cannot be found.</exception>
    private static string ExtractNewsTitle(HtmlDocument doc)
    {
        HtmlNode[] newsTitlesArray = doc.DocumentNode.SelectNodes("//h1[@class='newsTitle'] | //div[@class='firstTitle']")?.ToArray()
            ?? throw new XPathException("Title body node is missing. Cannot obtain any text");

        return string.Join("", newsTitlesArray.Select(node => CleanHtmlText(node.InnerText)));
    }

    /// <summary>
    /// Extracts the publication date from the HTML document.
    /// </summary>
    /// <param name="doc">The HTML document to extract from.</param>
    /// <returns>The extracted publication date.</returns>
    /// <exception cref="XPathException">Thrown when the date node cannot be found.</exception>
    private static string ExtractPublishDate(HtmlDocument doc)
    {
        HtmlNode[] publishDateArray = doc.DocumentNode.SelectNodes("//time[@datetime] | //div[@class='firstDate']")?.ToArray()
            ?? throw new XPathException("Date body node is missing. Cannot obtain any text");

        return publishDateArray[0].InnerText.Trim();
    }

    /// <summary>
    /// Extracts the news text content from the HTML document, removing banners.
    /// </summary>
    /// <param name="doc">The HTML document to extract from.</param>
    /// <returns>The extracted and cleaned news text.</returns>
    /// <exception cref="XPathException">Thrown when the text node cannot be found.</exception>
    private static string ExtractNewsText(HtmlDocument doc)
    {
        HtmlNode[] newsText = doc.DocumentNode.SelectNodes("//div[@class='newsText'] | //div[@class='interviewText']")?.ToArray()
            ?? throw new XPathException("Text body node is missing. Cannot obtain any text");

        // Remove useless banners
        foreach (var node in newsText)
        {
            var uselessBanners = node.SelectNodes(".//section[@class='read']");
            if (uselessBanners != null)
            {
                foreach (var banner in uselessBanners)
                {
                    banner.Remove();
                }
            }
        }

        return string.Join("", newsText.Select(node => CleanHtmlText(node.InnerText)));
    }

    /// <summary>
    /// Determines if the article is marked as exclusive.
    /// </summary>
    /// <param name="doc">The HTML document to check.</param>
    /// <returns>True if the article is exclusive, otherwise false.</returns>
    private static bool IsNewsExclusive(HtmlDocument doc)
    {
        return doc.DocumentNode.SelectSingleNode("//div[@class='newsPrefix']")?.InnerText == Constants.ArticleExclusivePrefix;
    }

    /// <summary>
    /// Removes the timestamp portion from a date string.
    /// </summary>
    /// <param name="publishDate">The date string that may contain a timestamp.</param>
    /// <returns>The date string without the timestamp, or the default date if input is null/empty.</returns>
    private static string RemoveTimestampFromDate(string? publishDate)
    {
        if (string.IsNullOrWhiteSpace(publishDate))
        {
            return Constants.DefaultDate;
        }

        int lastSpaceIndex = publishDate.LastIndexOf(" ", StringComparison.Ordinal);
        if (lastSpaceIndex < 0)
        {
            // No space found, return the date as-is
            return publishDate;
        }

        return publishDate.Substring(0, lastSpaceIndex);
    }

    /// <summary>
    /// Combines article components into a single text string.
    /// </summary>
    /// <param name="newsTitle">The news title.</param>
    /// <param name="fixedDate">The publication date.</param>
    /// <param name="newsText">The news text content.</param>
    /// <returns>The combined article text.</returns>
    private static string CombineArticleText(string newsTitle, string fixedDate, string newsText)
    {
        return $"{newsTitle}{Environment.NewLine}{fixedDate}{Environment.NewLine}{newsText}";
    }

    /// <summary>
    /// Determines the article type based on character count.
    /// </summary>
    /// <param name="textCharsAmount">The number of non-whitespace characters in the article.</param>
    /// <returns>The article type classification.</returns>
    private static string DetermineArticleType(int textCharsAmount)
    {
        return textCharsAmount switch
        {
            > 0 and <= Constants.ArticleTypeInfoMessageMaxChars => Constants.ArticleTypeInfoMessage,
            > Constants.ArticleTypeInfoMessageMaxChars and <= Constants.ArticleTypeExtendedInfoMessageMaxChars => Constants.ArticleTypeExtendedInfoMessage,
            >= Constants.ArticleTypeCommentMinChars => Constants.ArticleTypeComment,
            _ => Constants.ErrorArticleType
        };
    }

    /// <summary>
    /// Cleans HTML entities and special characters from text.
    /// </summary>
    /// <param name="text">The text to clean.</param>
    /// <returns>The cleaned text.</returns>
    private static string CleanHtmlText(string text)
    {
        return text.Replace("&ndash;", "-")
                   .Replace("&laquo;", "\"")
                   .Replace("&raquo;", "\"")
                   .Replace("&rsquo;", "'")
                   .Replace("&nbsp;", " ")
                   .Replace("&#039;", "'")
                   .Replace("&amp;", "&")
                   .Trim();
    }

    /// <summary>
    /// Counts the number of non-whitespace characters in a string.
    /// </summary>
    /// <param name="text">The text to count characters in.</param>
    /// <returns>The number of non-whitespace characters.</returns>
    private static int CountNonWhiteSpaceChars(string text)
    {
        return text.Count(c => !char.IsWhiteSpace(c));
    }

    /// <summary>
    /// Handles exceptions that occur during article parsing.
    /// </summary>
    /// <param name="ex">The exception that occurred.</param>
    /// <param name="articleSource">The article source that caused the exception.</param>
    private static void HandleParsingException(Exception ex, ArticleSource articleSource)
    {
        if (ex is HtmlWebException || ex is WebException)
        {
            Console.WriteLine($"{Constants.ConsolePrefixArticleParser} Error processing file: {Path.GetFileName(articleSource.FilePath)}");
        }
        else if (ex is XPathException)
        {
            Console.WriteLine($"{Constants.ConsolePrefixArticleParser} Error extracting data from: {articleSource.Link}");
        }
        else
        {
            Console.WriteLine($"{Constants.ConsolePrefixArticleParser} Unhandled error for file: {Path.GetFileName(articleSource.FilePath)}");
        }
    }

    /// <summary>
    /// Creates an error article when parsing fails.
    /// </summary>
    /// <param name="articleSource">The article source that failed to parse.</param>
    /// <returns>A WebParser instance with default values indicating an error.</returns>
    private static WebParser CreateErrorArticle(ArticleSource articleSource)
    {
        return new WebParser(link: articleSource.Link, filePath: articleSource.FilePath);
    }
}

