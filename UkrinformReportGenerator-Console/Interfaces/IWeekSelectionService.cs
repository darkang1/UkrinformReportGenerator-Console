using System;

namespace URG_Console.Interfaces;

/// <summary>
/// Service for selecting and calculating operating weeks.
/// </summary>
public interface IWeekSelectionService
{
    /// <summary>
    /// Prompts the user to select an operating week and returns the start and end dates.
    /// </summary>
    /// <param name="reportEndDay">The day of week that marks the end of the reporting week.</param>
    /// <returns>A tuple containing the start date and end date of the selected week.</returns>
    (DateTime startDate, DateTime endDate) SelectOperatingWeek(DayOfWeek reportEndDay);

    /// <summary>
    /// Calculates the end date of the week containing the specified date.
    /// </summary>
    /// <param name="start">The date to calculate from.</param>
    /// <param name="endDay">The day of week that marks the end of the week.</param>
    /// <returns>The end date of the week (on the specified end day).</returns>
    DateTime GetCurrentWeekEndDay(DateTime start, DayOfWeek endDay);
}

