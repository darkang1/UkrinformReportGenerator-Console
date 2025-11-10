namespace URG_Console;

/// <summary>
/// Represents a source article with its link and file path.
/// </summary>
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