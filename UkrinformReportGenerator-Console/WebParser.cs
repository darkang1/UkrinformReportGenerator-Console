using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Xml.XPath;
using HtmlAgilityPack;

namespace URG_Console
{
    internal class WebParser
    {

        //enum ArticleType
        //{
        //    [Description("Інф. повідомлення")]
        //    InfPovidom = 1,

        //    [Description("Розш. інф. повідомлення")]
        //    RozInfPovidom = 2,

        //    [Description("Коментар")]
        //    Comentar = 3,

        //}

        internal string articleDate { get; private set; } = "01.01.2000";
        internal string articleHeader { get; private set; } = "Unknown";
        internal string articleType { get; private set; } = "N/A";
        internal int articleChars { get; private set; } = 0;
        internal string articleLink { get; private set; } = "https://None";
        internal bool articleExclusive { get; private set; } = false;


        //private WebParser()
        //{
        //    //If link wasn't parsed successfuly, creating a 'default' article
        //}

        private WebParser(string date = "01.01.2000", string header = "Unknown", string type = "N/A", int chars = 0, string link = "https://None", bool exclusive = false)
        {
            articleDate = date;
            articleHeader = header;
            articleType = type;
            articleChars = chars;
            articleLink = link;
            articleExclusive = exclusive;
        }


        internal static List<WebParser> webParser(string[] links)
        {
            Console.OutputEncoding = Encoding.UTF8;
            Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("uk-UA");
           
            if (links == null || links.Length == 0)
                Console.WriteLine("\n[WebParser] No links were passed to the parser!");
            else
                Console.WriteLine("[WebParser] Links loaded successfully!");

            WebParser[] articles = new WebParser[links.Length];

            for (int i = 0; i < links.Length; i++)
            {
                try {
                    
                HtmlWeb web = new HtmlWeb();
                HtmlDocument doc = web.Load(links[i]);
                string newsTitle = doc.DocumentNode.SelectSingleNode("//h1[@class='newsTitle']")?.InnerText ?? "Unknown";
                string publishDate = doc.DocumentNode.SelectSingleNode("//time[@datetime]")?.InnerText.ToString() ?? "01.01.2000 ";
                HtmlNode[] newsText = doc.DocumentNode.SelectNodes("//div[@class='newsText']")?.ToArray() ?? throw new XPathException("Text body token is missing. Cannot obtain any text"); // Couldn't find a way to work with null-coalescing operator (??), so simply throw an exception
                bool newsExclusive = doc.DocumentNode.SelectSingleNode("//div[@class='newsPrefix']")?.InnerText == "Ексклюзив" ? true : false;
                string newsLink = links[i];

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

                // Replacing HTML tags in the article's title with proper symbols
                newsTitle = newsTitle.Replace("&ndash;", "-").Replace("&laquo;", "\"").Replace("&raquo;", "\"").Replace("&rsquo;", "'").Replace("&nbsp;", " ").Replace("&#039;", "'").Replace("&amp;", "&");

                string finalText = newsTitle + Environment.NewLine + fixedDate + Environment.NewLine;

                // Replacing HTML tags in the article's body with proper symbols and saving result to final string
                for (int j = 0; j < newsText.Count(); j++)
                {
                    finalText += newsText[j].InnerText.Replace("&ndash;", "-").Replace("&laquo;", "\"").Replace("&raquo;", "\"").Replace("&rsquo;", "'").Replace("&nbsp;", " ").Replace("&#039;", "'").Replace("&amp;", "&");
                }

                // Counting number of non-white space chars in text
                int textCharsAmount = countNonWhiteSpaceChars(finalText);

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
                articles[i] = new WebParser(fixedDate, newsTitle, newsType, textCharsAmount, newsLink, newsExclusive);
                Console.WriteLine($"Total processed links: [{i + 1}/{links.Length}]");

                }
                catch(HtmlWebException ex)
                {
                    Console.WriteLine("[WebParser] " + ex.Message);
                    articles[i] = new WebParser(link: links[i]);
                }
                catch (XPathException ex)
                {
                    Console.WriteLine($"[WebParser] At link: {links[i]}");
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

        private static int countNonWhiteSpaceChars(string text)
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
