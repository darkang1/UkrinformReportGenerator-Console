namespace URG_Console
{
    public class ArticleSource
    {
        public string Link { get; }
        public string FilePath { get; }
        public bool HasLink => !string.IsNullOrEmpty(Link);

        public ArticleSource(string link, string filePath)
        {
            Link = link;
            FilePath = filePath;
        }
    }
}