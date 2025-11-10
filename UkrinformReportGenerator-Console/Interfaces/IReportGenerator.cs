using System;
using System.Threading.Tasks;

namespace URG_Console.Interfaces;

/// <summary>
/// Interface for generating reports.
/// </summary>
public interface IReportGenerator
{
    /// <summary>
    /// Generates the report based on the provided parameters.
    /// </summary>
    /// <param name="folderPath">Path to the folder containing source documents.</param>
    /// <param name="reportStartDate">Start date of the report period.</param>
    /// <param name="reportEndDate">End date of the report period.</param>
    /// <param name="reportEndDay">Day of week that marks the end of the reporting week.</param>
    /// <returns>A task that represents the asynchronous operation.</returns>
    Task GenerateReportAsync(string folderPath, DateTime reportStartDate, DateTime reportEndDate, DayOfWeek reportEndDay);
}

