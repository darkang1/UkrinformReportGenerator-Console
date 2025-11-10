using System.Collections.Generic;
using System.Threading.Tasks;

namespace URG_Console.Interfaces;

/// <summary>
/// Interface for parsing articles from web sources.
/// </summary>
public interface IArticleParser
{
    /// <summary>
    /// Parses articles from the provided article sources.
    /// </summary>
    /// <param name="articleSources">List of article sources to parse.</param>
    /// <returns>A task that represents the asynchronous operation. The task result contains a list of parsed articles.</returns>
    Task<List<WebParser>> ParseArticlesAsync(List<ArticleSource> articleSources);
}

