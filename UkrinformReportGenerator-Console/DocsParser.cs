using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using Xceed.Words.NET;

namespace URG_Console
{
    internal class DocsParser
    {

        internal static string[] getLinks()
        {
            return parseAllDocsLinks(getDocsFilePath());
        }

        internal static string[] getLinks(string filesPath)
        {
            return parseAllDocsLinks(getDocsFilePath(filesPath));
        }

        private static string[] getDocsFilePath()
        {
            return getDocsFilePath(Directory.GetCurrentDirectory());
        }

        private static string[] getDocsFilePath(string path)
        {
            try
            {
                if (!Directory.Exists(path))
                    throw new DirectoryNotFoundException("[DocsParser] Selected directory doesn't exist!");

                // Adding all .docx files to the list and sorting them by last modification time
                List<string> wordFiles = new List<string>(Directory.GetFiles(path, "*.docx").OrderBy(ex => new FileInfo(ex).LastWriteTimeUtc));

                // Adding all .doc files to the list and sorting them by last modification time
                wordFiles.AddRange(Directory.GetFiles(path, "*.doc").OrderBy(ex => new FileInfo(ex).LastWriteTimeUtc));

                //Removing temporary '~$' files from the list, if there are any
                for (int i = 0; i < wordFiles.Count(); i++)
                {
                    if (wordFiles[i].Contains("~$"))
                    {
                        wordFiles.Remove(wordFiles[i]);
                        i--;
                    }
                }

                return wordFiles.ToArray();
            }
            catch(DirectoryNotFoundException ex)
            {
                Console.WriteLine(ex.Message + Environment.NewLine + "Cannot operate further." + Environment.NewLine + "Exiting...");
                Environment.Exit(-99);
                return new string[0];
            }
            catch(Exception ex)
            {
                Console.WriteLine("[DocsParser] Unknown exception occured: ");
                Console.WriteLine(ex.ToString());
                return new string[0];
            }
  
 
        }


        private static string[] parseAllDocsLinks(string[] files)
        {

            if (files == null)
            {
                Console.WriteLine("Null value was passed to DocsParser");
                throw new ArgumentNullException();
            }

            // Parsing all links and hyperlinks in given docx file
            List<string> allLinks = new List<string>();

            for (int i = 0; i < files.Count(); i++)
            {
                try
                {
                    var doc = DocX.Load(files[i]);

                    // Creating predicate to further check if any Ukrinform links exists in the file
                    Predicate<Xceed.Document.NET.Hyperlink> ukrLinks = hasUkrinformLinks;

                    if(!doc.Hyperlinks.Exists(ukrLinks) || doc.Hyperlinks.Count == 0)
                        throw new MissingMemberException($"No Ukrinform links were found in the file!");

                    // Checking if filename contains "1+[num]" in the beginning.
                    // It means that there are 1+someNum of articles in one file, which need to be parsed
                    string fileName = Path.GetFileName(files[i]);
                    if (fileName != null && fileName.Contains("1+"))
                    {
                        // Checking number between '+' and '_' symbols in the filename
                        // Ex: 1+5_USA.docx
                        int numOfAdditionalArticles = int.Parse(fileName.Substring(fileName.IndexOf('+') + 1, (fileName.IndexOf('_') - 2)));

                        // Creating a hyperlink list of all Ukrinform links found in file
                        var ukrinformLinks = doc.Hyperlinks.FindAll(ukrLinks);
                        
                        if (numOfAdditionalArticles + 1 > ukrinformLinks.Count)
                            throw new IndexOutOfRangeException($"Number of articles declared in the filename [1+{numOfAdditionalArticles}] is invalid!" + Environment.NewLine + "You have less article links in the file than declared!");

                        // Parsing declared number of articles
                        for (int z = 0; z < ukrinformLinks.Count; z++) 
                        {
                                allLinks.Add(ukrinformLinks[z].Uri.ToString());
                        }
                    }

                    else
                    {
                        foreach (var link in doc.Hyperlinks)
                        {
                            try
                            {
                                //Checking the first url in document and if it's related to Ukrinform. If not, then skip
                                if (link.Uri == doc.Hyperlinks[0].Uri) {
                                    if (link.Uri?.ToString().Contains("ukrinform") ?? throw new NullReferenceException("Null URL was detected when tried to add first document hyperlink to the list!" + Environment.NewLine + "(You might have broken saved file. Try to modify, resave it and try again)"))
                                        allLinks.Add(link.Uri.ToString());
                                }

                            }
                            catch (NullReferenceException ex)
                            {
                                // We have another instance of NullReferenceException handler here, since it's possible to have more than 1 NULL link in the file (usually 3).
                                // So not to cycle through all of them we simply break out of the current loop and moving forward to the next file.
                                // Usually having NULL links in file Hyperlinks list allegedly means we have a broken saved file (?) which needs to be resaved in order to fix this issue.
                                Console.WriteLine($"[DocsParser] File: {Path.GetFileName(files[i])}");
                                Console.WriteLine(ex.Message);
                                Console.WriteLine("Skipping file...");
                                break;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"[DocsParser] File: {Path.GetFileName(files[i])}");
                                Console.WriteLine("[DocsParser] Unhandled exception occured: ");
                                Console.WriteLine(ex.ToString());
                                Console.WriteLine("Skipping file...");
                                break;
                            }
                        }
                    }
                   
                }catch(IOException ex)
                {
                    killProcess("WINWORD");
                    Console.WriteLine(ex.Message);
                    Console.WriteLine("[DocsParser] *** Parsed documents cannot be open during the runtime. Run the program again... ***");
                    Environment.Exit(-10);
                    
                }
                catch (MissingMemberException ex)
                {
                    Console.WriteLine($"[DocsParser] File: {Path.GetFileName(files[i])}");
                    Console.WriteLine(ex.Message);
                    Console.WriteLine("Skipping file...");
                }
                catch (IndexOutOfRangeException ex)
                {
                    Console.WriteLine($"[DocsParser] File: {Path.GetFileName(files[i])}");
                    Console.WriteLine(ex.Message);
                    Console.WriteLine("Skipping file...");
                }
                catch (NullReferenceException ex)
                {
                    Console.WriteLine($"[DocsParser] File: {Path.GetFileName(files[i])}");
                    Console.WriteLine(ex.Message);
                    Console.WriteLine("Skipping file...");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[DocsParser] File: {Path.GetFileName(files[i])}");
                    Console.WriteLine("[DocsParser] Unhandled exception occured: ");
                    Console.WriteLine(ex.ToString());
                }
            }

            // Removing any duplicates, if occurred
            var allLinksFiltered = allLinks.Distinct().ToArray();
            Console.WriteLine($"\n[DocsParser] Successfully links parsed from documents: {allLinksFiltered.Length}");

            return allLinksFiltered;

        }

        private static bool hasUkrinformLinks(Xceed.Document.NET.Hyperlink link)
        {
            if (link.Uri != null)
                return link.Uri.ToString().Contains("ukrinform");
            else
                throw new NullReferenceException("Null URL was detected when searched for Ukrinform links!" + Environment.NewLine + "(You might have broken saved file. Try to modify, resave it and try again)");
        }

        private static void killProcess(string processName)
        {
            foreach (var process in Process.GetProcessesByName(processName))
            {
                process.Kill();
            }
        }



    }
}
