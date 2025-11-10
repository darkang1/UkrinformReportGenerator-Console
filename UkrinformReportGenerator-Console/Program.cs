using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using URG_Console.Configuration;
using URG_Console.Interfaces;
using URG_Console.Services;
using Xceed.Words.NET;

namespace URG_Console;

/// <summary>
/// Main entry point for the Ukrinform Report Generator console application.
/// </summary>
class Program
{
    private static AppSettings? _appSettings;
    private static readonly DayOfWeek ReportEndDay = DayOfWeek.Sunday;

    /// <summary>
    /// Main entry point of the application.
    /// </summary>
    /// <param name="args">Command line arguments (currently unused).</param>
    static async Task Main(string[] args)
    {
        InitializeEnvironment();
        ClearLicenseMsg();
        await RunReportGeneratorAsync();
    }

    /// <summary>
    /// Initializes the application environment including culture settings and configuration.
    /// </summary>
    private static void InitializeEnvironment()
    {
        Console.OutputEncoding = Encoding.UTF8;
        Thread.CurrentThread.CurrentCulture = new CultureInfo("uk-UA", false);
        Thread.CurrentThread.CurrentUICulture = new CultureInfo("uk-UA", false);

        var configuration = new ConfigurationBuilder()
            .SetBasePath(Directory.GetCurrentDirectory())
            .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
            .Build();

        _appSettings = new AppSettings
        {
            DefaultPath = configuration["DefaultPath"] ?? string.Empty,
            WordExecutablePath = configuration["WordExecutablePath"] ?? @"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
        };
    }

    /// <summary>
    /// Runs the report generation process asynchronously.
    /// </summary>
    private static async Task RunReportGeneratorAsync()
    {
        Console.WriteLine("=======Ukrinform Report Generator=======");

        if (_appSettings == null)
        {
            Console.WriteLine($"{Constants.ConsolePrefixProgram} Error: Application settings not initialized!");
            return;
        }

        // Set up dependency injection
        var serviceProvider = BuildServiceProvider(_appSettings, ReportEndDay);
        var weekSelectionService = serviceProvider.GetRequiredService<IWeekSelectionService>();
        var pathValidationService = serviceProvider.GetRequiredService<IPathValidationService>();
        var reportGenerator = serviceProvider.GetRequiredService<IReportGenerator>();

        // Get user input
        (DateTime reportStartDate, DateTime reportEndDate) = weekSelectionService.SelectOperatingWeek(ReportEndDay);
        string folderPath = await pathValidationService.SelectFolderLocationAsync(_appSettings.GetDefaultPathOrDefault());

        Console.WriteLine();

        // Generate report
        await reportGenerator.GenerateReportAsync(folderPath, reportStartDate, reportEndDate, ReportEndDay);
    }

    /// <summary>
    /// Builds and configures the dependency injection service provider.
    /// </summary>
    /// <param name="appSettings">The application settings to register.</param>
    /// <param name="reportEndDay">The day of week that marks the end of the reporting week.</param>
    /// <returns>A configured service provider with all services registered.</returns>
    private static ServiceProvider BuildServiceProvider(AppSettings appSettings, DayOfWeek reportEndDay)
    {
        var services = new ServiceCollection();

        // Register configuration
        services.AddSingleton(appSettings);

        // Register business services
        services.AddSingleton<IDocumentParser, DocumentParser>();
        services.AddSingleton<IArticleParser, ArticleParser>();
        services.AddSingleton(sp => new ReportDateCalculator(reportEndDay));
        services.AddSingleton<ReportHeaderBuilder>();
        services.AddSingleton(sp => new WordReportBuilder(appSettings));
        services.AddSingleton<ArticleDisplayService>();

        // Register user interaction services
        services.AddSingleton<IWeekSelectionService, WeekSelectionService>();
        services.AddSingleton<IPathValidationService, PathValidationService>();

        // Register main orchestrator
        services.AddSingleton<IReportGenerator, ReportGenerator>();

        return services.BuildServiceProvider();
    }

    /// <summary>
    /// Clears the license message that appears when DocX library is first used.
    /// This is a workaround to remove the license message from the console output.
    /// </summary>
    private static void ClearLicenseMsg()
    {
        try
        {
            DocX.Load((string)null!);
        }
        catch (Exception)
        {
            // Expected exception when loading null to trigger license message
            // Silently clear console to remove license message
            Console.Clear();
        }
    }
}