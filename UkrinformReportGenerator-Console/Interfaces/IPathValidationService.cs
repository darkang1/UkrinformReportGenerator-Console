using System.Threading.Tasks;

namespace URG_Console.Interfaces;

/// <summary>
/// Service for validating and selecting folder paths.
/// </summary>
public interface IPathValidationService
{
    /// <summary>
    /// Prompts the user to select a folder location and validates it.
    /// </summary>
    /// <param name="defaultPath">The default path to use if user selects option 1.</param>
    /// <returns>A validated folder path that exists on the file system.</returns>
    Task<string> SelectFolderLocationAsync(string defaultPath);
}

