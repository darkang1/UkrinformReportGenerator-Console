using System.Collections.Generic;
using System.Threading.Tasks;

namespace URG_Console.Interfaces;

/// <summary>
/// Interface for parsing documents to extract article sources.
/// </summary>
public interface IDocumentParser
{
    /// <summary>
    /// Parses documents in the specified directory and extracts article sources.
    /// </summary>
    /// <param name="directoryPath">Path to the directory containing documents.</param>
    /// <returns>A task that represents the asynchronous operation. The task result contains a list of article sources extracted from documents.</returns>
    Task<List<ArticleSource>> ParseDocumentsInDirectoryAsync(string directoryPath);
}

