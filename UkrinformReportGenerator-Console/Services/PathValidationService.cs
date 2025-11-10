using System;
using System.IO;
using System.Threading.Tasks;
using URG_Console.Interfaces;

namespace URG_Console.Services;

/// <summary>
/// Service for validating and selecting folder paths.
/// </summary>
public class PathValidationService : IPathValidationService
{
    /// <inheritdoc/>
    public async Task<string> SelectFolderLocationAsync(string defaultPath)
    {
        string folderPath;
        bool isValidPath;

        do
        {
            DisplayFolderOptions(defaultPath);
            int choice = GetValidUserChoice();
            folderPath = choice == 1 ? defaultPath : GetUserSpecifiedPath();

            if (string.IsNullOrWhiteSpace(folderPath))
            {
                Console.WriteLine("Invalid directory path! Try again");
                isValidPath = false;
            }
            else
            {
                isValidPath = await Task.Run(() => Directory.Exists(folderPath));
                if (!isValidPath)
                {
                    Console.WriteLine("Invalid directory path! Try again");
                }
            }
        } while (!isValidPath);

        return folderPath;
    }

    /// <summary>
    /// Displays options for selecting a folder location.
    /// </summary>
    /// <param name="defaultPath">The default path to display.</param>
    private static void DisplayFolderOptions(string defaultPath)
    {
        Console.WriteLine("\nSelect folder location:");
        Console.WriteLine($"1. Use default location \n({defaultPath})");
        Console.WriteLine("2. Manually specify folder path");
    }

    /// <summary>
    /// Gets a valid user choice (1 or 2) for folder selection.
    /// </summary>
    /// <returns>The user's valid choice.</returns>
    private static int GetValidUserChoice()
    {
        while (true)
        {
            Console.Write("> ");
            string? input = Console.ReadLine();

            if (int.TryParse(input, out int choice) && (choice == 1 || choice == 2))
            {
                return choice;
            }
            Console.WriteLine("Invalid input! Try again");
        }
    }

    /// <summary>
    /// Prompts the user to enter a folder path.
    /// </summary>
    /// <returns>The folder path entered by the user.</returns>
    private static string GetUserSpecifiedPath()
    {
        Console.WriteLine();
        Console.WriteLine("Enter the folder path:");
        Console.Write("> ");
        return Console.ReadLine() ?? string.Empty;
    }
}

