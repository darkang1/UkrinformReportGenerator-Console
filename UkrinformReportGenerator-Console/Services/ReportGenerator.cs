using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using URG_Console.Interfaces;

namespace URG_Console.Services;

/// <summary>
/// Orchestrates the report generation process.
/// </summary>
public class ReportGenerator : IReportGenerator
{
    private readonly IDocumentParser _documentParser;
    private readonly IArticleParser _articleParser;
    private readonly ReportDateCalculator _dateCalculator;
    private readonly ReportHeaderBuilder _headerBuilder;
    private readonly WordReportBuilder _wordReportBuilder;
    private readonly ArticleDisplayService _displayService;

    /// <summary>
    /// Initializes a new instance of the <see cref="ReportGenerator"/> class.
    /// </summary>
    /// <param name="documentParser">The document parser service.</param>
    /// <param name="articleParser">The article parser service.</param>
    /// <param name="dateCalculator">The date calculator service.</param>
    /// <param name="headerBuilder">The header builder service.</param>
    /// <param name="wordReportBuilder">The Word report builder service.</param>
    /// <param name="displayService">The article display service.</param>
    /// <exception cref="ArgumentNullException">Thrown when any of the required dependencies are null.</exception>
    public ReportGenerator(
        IDocumentParser documentParser,
        IArticleParser articleParser,
        ReportDateCalculator dateCalculator,
        ReportHeaderBuilder headerBuilder,
        WordReportBuilder wordReportBuilder,
        ArticleDisplayService displayService)
    {
        _documentParser = documentParser ?? throw new ArgumentNullException(nameof(documentParser));
        _articleParser = articleParser ?? throw new ArgumentNullException(nameof(articleParser));
        _dateCalculator = dateCalculator ?? throw new ArgumentNullException(nameof(dateCalculator));
        _headerBuilder = headerBuilder ?? throw new ArgumentNullException(nameof(headerBuilder));
        _wordReportBuilder = wordReportBuilder ?? throw new ArgumentNullException(nameof(wordReportBuilder));
        _displayService = displayService ?? throw new ArgumentNullException(nameof(displayService));
    }

    /// <inheritdoc/>
    public async Task GenerateReportAsync(string folderPath, DateTime reportStartDate, DateTime reportEndDate, DayOfWeek reportEndDay)
    {
        // Calculate report date information
        ReportDateInfo dateInfo = _dateCalculator.CalculateReportDateInfo(reportStartDate, reportEndDate);

        // Build header
        string header = _headerBuilder.BuildHeader(dateInfo);

        // Parse documents and articles
        List<ArticleSource> articleSources = await _documentParser.ParseDocumentsInDirectoryAsync(folderPath);
        List<WebParser> parsedArticles = await _articleParser.ParseArticlesAsync(articleSources);

        // Sort and display articles
        parsedArticles = _displayService.SortArticlesByDate(parsedArticles);
        _displayService.DisplayParsedArticles(parsedArticles);

        // Generate Word report
        await _wordReportBuilder.GenerateReportAsync(folderPath, parsedArticles, header, dateInfo);
    }
}

