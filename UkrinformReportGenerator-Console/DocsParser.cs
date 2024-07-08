using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xceed.Words.NET;

namespace URG_Console
{
    public class DocsParser
    {
        public static List<ArticleSource> ParseDocumentsInDirectory(string directoryPath)
        {
            string[] filePaths = GetDocsFilePath(directoryPath);
            return ExtractLinksFromDocuments(filePaths);
        }

        private static string[] GetDocsFilePath(string path)
        {
            if (!Directory.Exists(path))
                throw new DirectoryNotFoundException("[DocsParser] Selected directory doesn't exist!");

            var wordFiles = Directory.GetFiles(path, "*.docx")
                .Concat(Directory.GetFiles(path, "*.doc"))
                .Where(file => !file.Contains("~$"))
                .OrderBy(file => new FileInfo(file).LastWriteTimeUtc)
                .ToArray();

            return wordFiles;
        }

        private static List<ArticleSource> ExtractLinksFromDocuments(string[] filePaths)
        {
            if (filePaths == null)
                throw new ArgumentNullException(nameof(filePaths), "Null filePaths array was passed to DocsParser!");

            var articleSources = new List<ArticleSource>();
            foreach (var filePath in filePaths)
            {
                try
                {
                    ProcessFile(filePath, articleSources);
                }
                catch (Exception ex)
                {
                    HandleFileProcessingException(ex, filePath);
                }
            }

            Console.WriteLine($"\n[DocsParser] Successfully extracted article sources from documents: {articleSources.Count}");
            return articleSources;
        }

        private static void ProcessFile(string filePath, List<ArticleSource> articleSources)
        {
            using var doc = DocX.Load(filePath);
            if (!doc.Hyperlinks.Exists(HasUkrinformLinks))
            {
                articleSources.Add(new ArticleSource("", filePath));
                throw new MissingMemberException("No Ukrinform links were found in the file!");
            }

            string fileName = Path.GetFileName(filePath);
            if (fileName.Contains("1+"))
            {
                ProcessMultiArticleFile(doc, filePath, articleSources, fileName);
            }
            else
            {
                ProcessSingleArticleFile(doc, filePath, articleSources);
            }
        }

        private static void ProcessMultiArticleFile(DocX doc, string filePath, List<ArticleSource> articleSources, string fileName)
        {
            int numOfAdditionalArticles = ExtractAdditionalArticlesCount(fileName);
            var ukrinformLinks = doc.Hyperlinks.FindAll(HasUkrinformLinks);

            if (numOfAdditionalArticles + 1 > ukrinformLinks.Count)
            {
                articleSources.Add(new ArticleSource("", filePath));
                throw new IndexOutOfRangeException($"Number of articles declared in the filename [1+{numOfAdditionalArticles}] is invalid!\nYou have less article links in the file than declared!");
            }

            foreach (var link in ukrinformLinks)
            {
                AddArticleSource(articleSources, link.Uri?.ToString(), filePath);
            }
        }

        private static void ProcessSingleArticleFile(DocX doc, string filePath, List<ArticleSource> articleSources)
        {
            var firstLink = doc.Hyperlinks[0];
            string? linkUri = firstLink.Uri?.ToString();

            if (linkUri is null)
            {
                articleSources.Add(new ArticleSource("", filePath));
                throw new NullReferenceException("Null URL was detected when tried to add first document hyperlink to the list!\n(You might have broken saved file. Try to modify, resave it and try again)");
            }

            if (linkUri.Contains("ukrinform"))
            {
                AddArticleSource(articleSources, linkUri, filePath);
            }
            else
            {
                articleSources.Add(new ArticleSource("", filePath));
                throw new MissingMemberException("First hyperlink is not Ukrinform related!");
            }
        }

        private static void AddArticleSource(List<ArticleSource> articleSources, string? link, string filePath)
        {
            var potentialDuplicateArticle = articleSources.FirstOrDefault(x => x.Link == link);

            if (link is null || potentialDuplicateArticle != default(ArticleSource))
            {
                articleSources.Add(new ArticleSource("", filePath));
                throw new InvalidDataException($"Same link has already been seen in another file! ({Path.GetFileName(potentialDuplicateArticle?.FilePath)})");
            }
            articleSources.Add(new ArticleSource(link, filePath));
        }

        private static int ExtractAdditionalArticlesCount(string fileName)
        {
            int startIndex = fileName.IndexOf('+') + 1;
            int endIndex = fileName.IndexOf('_', startIndex);
            return int.Parse(fileName[startIndex..endIndex]);
        }

        private static bool HasUkrinformLinks(Xceed.Document.NET.Hyperlink link)
        {
            return link.Uri?.ToString().Contains("ukrinform") ?? false;
        }

        private static void HandleFileProcessingException(Exception ex, string filePath)
        {
            Console.WriteLine($"[DocsParser] File: {Path.GetFileName(filePath)}");
            Console.WriteLine(ex.Message);
            Console.WriteLine("Skipping file...");
        }
    }
}