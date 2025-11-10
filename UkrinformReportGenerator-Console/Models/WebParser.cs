using System.IO;

namespace URG_Console;

/// <summary>
/// Represents a parsed article with all its metadata.
/// </summary>
public class WebParser
{
    /// <summary>
    /// Gets the article publication date.
    /// </summary>
    public string ArticleDate { get; private set; } = Constants.DefaultDate;

    /// <summary>
    /// Gets the article header/title.
    /// </summary>
    public string ArticleHeader { get; private set; } = Constants.DefaultHeader;

    /// <summary>
    /// Gets the article type classification.
    /// </summary>
    public string ArticleType { get; private set; } = Constants.DefaultArticleType;

    /// <summary>
    /// Gets the number of non-whitespace characters in the article.
    /// </summary>
    public int ArticleChars { get; private set; } = 0;

    /// <summary>
    /// Gets the article URL link.
    /// </summary>
    public string ArticleLink { get; private set; } = "";

    /// <summary>
    /// Gets a value indicating whether the article is marked as exclusive.
    /// </summary>
    public bool ArticleExclusive { get; private set; } = false;

    /// <summary>
    /// Gets the file path of the source document.
    /// </summary>
    public string ArticleFilePath { get; private set; } = string.Empty;

    /// <summary>
    /// Gets the file name of the source document.
    /// </summary>
    public string ArticleFileName { get; private set; } = string.Empty;

    /// <summary>
    /// Initializes a new instance of the <see cref="WebParser"/> class.
    /// </summary>
    /// <param name="date">The article publication date.</param>
    /// <param name="header">The article header/title.</param>
    /// <param name="type">The article type classification.</param>
    /// <param name="chars">The number of non-whitespace characters.</param>
    /// <param name="link">The article URL link.</param>
    /// <param name="exclusive">Whether the article is marked as exclusive.</param>
    /// <param name="filePath">The file path of the source document.</param>
    public WebParser(string date = Constants.DefaultDate, string header = Constants.DefaultHeader, string type = Constants.DefaultArticleType, int chars = 0, string link = "", bool exclusive = false, string filePath = "")
    {
        ArticleDate = date;
        ArticleHeader = header;
        ArticleType = type;
        ArticleChars = chars;
        ArticleLink = link;
        ArticleExclusive = exclusive;
        ArticleFilePath = filePath;
        ArticleFileName = !string.IsNullOrEmpty(filePath) ? Path.GetFileName(filePath) : string.Empty;
    }
}

