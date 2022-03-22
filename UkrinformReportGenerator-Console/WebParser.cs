using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Xml.XPath;
using HtmlAgilityPack;
using System.ComponentModel;
using System.IO;

namespace URG_Console
{
    internal class WebParser
    {
        // Internal class variables
        internal string ArticleDate { get; private set; } = "**.**.****";
        internal string ArticleHeader { get; private set; } = "Unknown";
        internal string ArticleType { get; private set; } = "N/A";
        internal int ArticleChars { get; private set; } = 0;
        internal string ArticleLink { get; private set; } = "https://nolink.ukrinform/";
        internal bool ArticleExclusive { get; private set; } = false;
        internal string ArticleFilePath { get; private set; } = String.Empty;

        public WebParser(string date = "**.**.****", string header = "Unknown", string type = "N/A", int chars = 0, string link = "https://nolink.ukrinform/", bool exclusive = false, string filePath = "")
        {
            ArticleDate = date;
            ArticleHeader = header;
            ArticleType = type;
            ArticleChars = chars;
            ArticleLink = link;
            ArticleExclusive = exclusive;
            ArticleFilePath = filePath;
        }

        internal static List<WebParser> ParseArticles(Dictionary<string, string> fileLinks)
        {
            Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("uk-UA");

            if (fileLinks == null || fileLinks.Count == 0)
                Console.WriteLine("\n[WebParser] No links were passed to the parser!");
            else
                Console.WriteLine("[WebParser] Links loaded successfully!");

            WebParser[] articles = new WebParser[fileLinks.Count];

            for (int i = 0; i < fileLinks.Count; i++)
            {
                try
                {

                    HtmlWeb web = new HtmlWeb();

                    if (fileLinks.ElementAt(i).Key.Contains("https://nolink.ukrinform"))
                    {
                        throw new HtmlWebException("A file with no link was passed to WebParser!");
                    }

                    HtmlDocument doc = web.Load(fileLinks.ElementAt(i).Key);

                    //string newsTitle1 = doc.DocumentNode.SelectSingleNode("//h1[@class='newsTitle']")?.InnerText ?? "Unknown";
                    //string publishDate = doc.DocumentNode.SelectSingleNode("//time[@datetime]")?.InnerText.ToString() ?? "01.01.2000 ";

                    string newsTitle = "";
                    HtmlNode[] newsTitlesArray = doc.DocumentNode.SelectNodes("//h1[@class='newsTitle'] | //div[@class='firstTitle']")?.ToArray() ?? throw new XPathException("Title body node is missing. Cannot obtain any text");
                    // Replacing HTML tags in the article's title with proper symbols
                    for (int j = 0; j < newsTitlesArray.Count(); j++)
                        newsTitle += newsTitlesArray[j].InnerText.Replace("&ndash;", "-").Replace("&laquo;", "\"").Replace("&raquo;", "\"").Replace("&rsquo;", "'").Replace("&nbsp;", " ").Replace("&#039;", "'").Replace("&amp;", "&").Trim();

                    string publishDate = "";
                    HtmlNode[] publishDateArray = doc.DocumentNode.SelectNodes("//time[@datetime] | //div[@class='firstDate']")?.ToArray() ?? throw new XPathException("Date body node is missing. Cannot obtain any text");
                    publishDate += publishDateArray[0].InnerText.Trim();

                    HtmlNode[] newsText = doc.DocumentNode.SelectNodes("//div[@class='newsText'] | //div[@class='interviewText']")?.ToArray() ?? throw new XPathException("Text body node is missing. Cannot obtain any text"); // Couldn't find a way to work with null-coalescing operator (??), so simply throw an exception
                    bool newsExclusive = doc.DocumentNode.SelectSingleNode("//div[@class='newsPrefix']")?.InnerText == "Ексклюзив" ? true : false;
                    string newsLink = fileLinks.ElementAt(i).Key;

                    string newsLinkFilePath = fileLinks.ElementAt(i).Value;

                    // Removing timestamp from full publish date
                    string fixedDate = publishDate?.Substring(0, publishDate.LastIndexOf(" "));

                    // Removing useless banners, such as 'Читайте також'
                    var uselessBanners = doc.DocumentNode.SelectNodes("//section[@class='read']");

                    if (uselessBanners != null)
                    {
                        foreach (var item in uselessBanners)
                        {
                            item.Remove();
                        }
                    }

                    string finalText = newsTitle + Environment.NewLine + fixedDate + Environment.NewLine;

                    // Replacing HTML tags in the article's body with proper symbols and saving result to final string
                    for (int j = 0; j < newsText.Count(); j++)
                    {
                        finalText += newsText[j].InnerText.Replace("&ndash;", "-").Replace("&laquo;", "\"").Replace("&raquo;", "\"").Replace("&rsquo;", "'").Replace("&nbsp;", " ").Replace("&#039;", "'").Replace("&amp;", "&");
                    }

                    // Counting number of non-white space chars in text
                    int textCharsAmount = CountNonWhiteSpaceChars(finalText);

                    // Setting article type
                    string newsType = "Undefined";
                    if (textCharsAmount > 0 && textCharsAmount < 1400)
                        newsType = "Інф. повідомлення";
                    else if (textCharsAmount >= 1400 && textCharsAmount < 5000)
                        newsType = "Розш. інф. повідомлення";
                    else if (textCharsAmount >= 5000)
                        newsType = "Коментар";
                    else
                        newsType = "Error";

                    // Creating article object with all parsed data
                    articles[i] = new WebParser(fixedDate, newsTitle, newsType, textCharsAmount, newsLink, newsExclusive, newsLinkFilePath);
                    Console.WriteLine($"Total processed links: [{i + 1}/{fileLinks.Count}]");

                }
                catch (HtmlWebException ex)
                {
                    Console.WriteLine($"[WebParser] File: {Path.GetFileName(fileLinks.ElementAt(i).Value)}");
                    Console.WriteLine(ex.Message);
                    articles[i] = new WebParser(link: fileLinks.ElementAt(i).Key, filePath: fileLinks.ElementAt(i).Value);
                }
                catch (XPathException ex)
                {
                    Console.WriteLine($"[WebParser] At link: {fileLinks.ElementAt(i).Key}");
                    Console.WriteLine(ex.Message);
                    Console.WriteLine("Skipping to the next article...");
                    articles[i] = new WebParser();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("[WebParser] Unhandled exception occured: ");
                    Console.WriteLine(ex.ToString());
                }
            }

            return articles.ToList();

        }

        private static int CountNonWhiteSpaceChars(string text)
        {
            int result = 0;
            foreach (char c in text)
            {
                if (!char.IsWhiteSpace(c))
                {
                    result++;
                }
            }
            return result;
        }



    }
}
