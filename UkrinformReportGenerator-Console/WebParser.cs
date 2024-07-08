using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.XPath;
using System.Globalization;
using System.Net;

namespace URG_Console
{
    public class WebParser
    {
        public string ArticleDate { get; private set; } = "**.**.****";
        public string ArticleHeader { get; private set; } = "Unknown";
        public string ArticleType { get; private set; } = "N/A";
        public int ArticleChars { get; private set; } = 0;
        public string ArticleLink { get; private set; } = "";
        public bool ArticleExclusive { get; private set; } = false;
        public string ArticleFilePath { get; private set; } = String.Empty;

        public WebParser(string date = "**.**.****", string header = "Unknown", string type = "N/A", int chars = 0, string link = "", bool exclusive = false, string filePath = "")
        {
            ArticleDate = date;
            ArticleHeader = header;
            ArticleType = type;
            ArticleChars = chars;
            ArticleLink = link;
            ArticleExclusive = exclusive;
            ArticleFilePath = filePath;
        }

        public static List<WebParser> ParseArticles(List<ArticleSource> articleSources)
        {
            SetCultureToUkrainian();
            ValidateArticleSources(articleSources);

            var articles = new List<WebParser>();

            for (int i = 0; i < articleSources.Count; i++)
            {
                try
                {
                    var article = ParseSingleArticle(articleSources[i]);
                    articles.Add(article);
                    Console.WriteLine($"Total processed articles: [{i + 1}/{articleSources.Count}]");
                }
                catch (Exception ex)
                {
                    HandleParsingException(ex, articleSources[i]);
                    articles.Add(CreateErrorArticle(articleSources[i]));
                }
            }

            return articles;
        }

        private static void SetCultureToUkrainian()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("uk-UA");
        }

        private static void ValidateArticleSources(List<ArticleSource> articleSources)
        {
            if (articleSources == null || articleSources.Count == 0)
                Console.WriteLine("\n[WebParser] No article sources were passed to the parser!");
            else
                Console.WriteLine("\n[WebParser] Article sources loaded successfully!");
        }

        private static WebParser ParseSingleArticle(ArticleSource articleSource)
        {
            if (!articleSource.HasLink)
            {
                Console.WriteLine($"[WebParser] No link found for file: {Path.GetFileName(articleSource.FilePath)}");
                return CreateErrorArticle(articleSource);
            }

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
                Console.WriteLine($"[WebParser] Unable to load page for file: {Path.GetFileName(articleSource.FilePath)}");
                return CreateErrorArticle(articleSource);
            }
            catch (XPathException ex)
            {
                Console.WriteLine($"[WebParser] Error extracting data from: {articleSource.Link}. Error: {ex.Message}");
                return CreateErrorArticle(articleSource);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[WebParser] Unexpected error processing file: {Path.GetFileName(articleSource.FilePath)}. Error: {ex.Message}");
                return CreateErrorArticle(articleSource);
            }
        }

        private static string ExtractNewsTitle(HtmlDocument doc)
        {
            HtmlNode[] newsTitlesArray = doc.DocumentNode.SelectNodes("//h1[@class='newsTitle'] | //div[@class='firstTitle']")?.ToArray() 
                ?? throw new XPathException("Title body node is missing. Cannot obtain any text");

            return string.Join("", newsTitlesArray.Select(node => CleanHtmlText(node.InnerText)));
        }

        private static string ExtractPublishDate(HtmlDocument doc)
        {
            HtmlNode[] publishDateArray = doc.DocumentNode.SelectNodes("//time[@datetime] | //div[@class='firstDate']")?.ToArray() 
                ?? throw new XPathException("Date body node is missing. Cannot obtain any text");

            return publishDateArray[0].InnerText.Trim();
        }

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

        private static bool IsNewsExclusive(HtmlDocument doc)
        {
            return doc.DocumentNode.SelectSingleNode("//div[@class='newsPrefix']")?.InnerText == "Ексклюзив";
        }

        private static string RemoveTimestampFromDate(string publishDate)
        {
            return publishDate?.Substring(0, publishDate.LastIndexOf(" "));
        }
        
        private static string CombineArticleText(string newsTitle, string fixedDate, string newsText)
        {
            return $"{newsTitle}{Environment.NewLine}{fixedDate}{Environment.NewLine}{newsText}";
        }

        private static string DetermineArticleType(int textCharsAmount)
        {
            if (textCharsAmount > 0 && textCharsAmount < 1400)
                return "Інф. повідомлення";
            else if (textCharsAmount >= 1400 && textCharsAmount < 5000)
                return "Розш. інф. повідомлення";
            else if (textCharsAmount >= 5000)
                return "Коментар";
            else
                return "Error";
        }

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

        private static int CountNonWhiteSpaceChars(string text)
        {
            return text.Count(c => !char.IsWhiteSpace(c));
        }

        private static void HandleParsingException(Exception ex, ArticleSource articleSource)
        {
            if (ex is HtmlWebException || ex is WebException)
            {
                Console.WriteLine($"[WebParser] Error processing file: {Path.GetFileName(articleSource.FilePath)}");
            }
            else if (ex is XPathException)
            {
                Console.WriteLine($"[WebParser] Error extracting data from: {articleSource.Link}");
            }
            else
            {
                Console.WriteLine($"[WebParser] Unhandled error for file: {Path.GetFileName(articleSource.FilePath)}");
            }
        }

        private static WebParser CreateErrorArticle(ArticleSource articleSource)
        {
            return new WebParser(link: articleSource.Link, filePath: articleSource.FilePath);
        }
    }
}