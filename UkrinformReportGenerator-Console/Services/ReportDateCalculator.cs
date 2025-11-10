using System;

namespace URG_Console.Services;

/// <summary>
/// Service for calculating report dates and week numbers based on the 4-day rule.
/// A week belongs to a month if at least 4 days of that week are within that month.
/// </summary>
public class ReportDateCalculator
{
    private enum Month
    {
        січня = 1, лютого = 2, березня = 3, квітня = 4, травня = 5, червня = 6,
        липня = 7, серпня = 8, вересня = 9, жовтня = 10, листопада = 11, грудня = 12
    }

    private readonly DayOfWeek _reportEndDay;

    /// <summary>
    /// Initializes a new instance of the <see cref="ReportDateCalculator"/> class.
    /// </summary>
    /// <param name="reportEndDay">The day of week that marks the end of the reporting week (e.g., Sunday).</param>
    public ReportDateCalculator(DayOfWeek reportEndDay)
    {
        _reportEndDay = reportEndDay;
    }

    /// <summary>
    /// Calculates report metadata including month, year, and week number.
    /// </summary>
    /// <param name="reportStartDate">The start date of the report period.</param>
    /// <param name="reportEndDate">The end date of the report period.</param>
    /// <returns>A <see cref="ReportDateInfo"/> object containing calculated date information.</returns>
    public ReportDateInfo CalculateReportDateInfo(DateTime reportStartDate, DateTime reportEndDate)
    {
        // Normalize ReportEndDate to ensure it's on ReportEndDay (defensive programming)
        DateTime normalizedWeekEndDate = GetWeekEndDate(reportEndDate);
        DateTime normalizedWeekStartDate = normalizedWeekEndDate.AddDays(-6);

        // Determine which month the report week belongs to (for file naming)
        // A week belongs to a month if at least 4 days of that week are within that month
        (DateTime targetMonth, int weekNumber) = DetermineWeekMonthAndNumber(normalizedWeekStartDate, normalizedWeekEndDate);

        return new ReportDateInfo
        {
            WeekStartDay = reportStartDate.Day,
            WeekEndDay = reportEndDate.Day,
            FileMonth = (int)targetMonth.Month,
            FileYear = targetMonth.Year,
            WeekNumber = weekNumber,
            IsReportSpanningTwoMonths = reportStartDate.Month != reportEndDate.Month,
            IsReportSpanningTwoYears = reportStartDate.Year != reportEndDate.Year,
            StartMonth = reportStartDate.Month,
            StartYear = reportStartDate.Year,
            EndMonth = reportEndDate.Month,
            EndYear = reportEndDate.Year
        };
    }

    /// <summary>
    /// Gets the end date of the week containing the given date.
    /// The week ends on the configured report end day (e.g., Sunday).
    /// </summary>
    /// <param name="date">The date to find the week end for.</param>
    /// <returns>The end date of the week (on the report end day).</returns>
    private DateTime GetWeekEndDate(DateTime date)
    {
        int daysUntilEndDay = ((int)_reportEndDay - (int)date.DayOfWeek + 7) % 7;
        if (daysUntilEndDay == 0 && date.DayOfWeek != _reportEndDay)
        {
            daysUntilEndDay = 7;
        }
        return date.AddDays(daysUntilEndDay).Date;
    }

    /// <summary>
    /// Determines which month a week belongs to and calculates the week number within that month.
    /// Uses the 4-day rule: a week belongs to a month if at least 4 days of that week are within that month.
    /// </summary>
    /// <param name="weekStart">The start date of the week.</param>
    /// <param name="weekEnd">The end date of the week.</param>
    /// <returns>A tuple containing the target month (first day of month) and the week number (1-5).</returns>
    /// <exception cref="ArgumentException">Thrown when weekStart is after weekEnd.</exception>
    private (DateTime targetMonth, int weekNumber) DetermineWeekMonthAndNumber(DateTime weekStart, DateTime weekEnd)
    {
        if (weekStart > weekEnd)
        {
            throw new ArgumentException("Week start date must be before or equal to week end date.");
        }

        int daysInStartMonth = 0;
        int daysInEndMonth = 0;
        bool weekSpansTwoMonths = weekStart.Month != weekEnd.Month || weekStart.Year != weekEnd.Year;

        for (DateTime day = weekStart; day <= weekEnd; day = day.AddDays(1))
        {
            if (day.Month == weekStart.Month && day.Year == weekStart.Year)
            {
                daysInStartMonth++;
            }

            // Only count days in end month if the week actually spans two months
            if (weekSpansTwoMonths && day.Month == weekEnd.Month && day.Year == weekEnd.Year)
            {
                daysInEndMonth++;
            }
        }

        DateTime targetMonth;
        if (daysInStartMonth >= 4)
        {
            targetMonth = new DateTime(weekStart.Year, weekStart.Month, 1);
        }
        else if (daysInEndMonth >= 4)
        {
            targetMonth = new DateTime(weekEnd.Year, weekEnd.Month, 1);
        }
        else
        {
            Console.WriteLine($"{Constants.ConsolePrefixReportDateCalculator} Warning: Neither month has 4+ days. Start month: {daysInStartMonth}, End month: {daysInEndMonth}");
            targetMonth = daysInStartMonth >= daysInEndMonth
                ? new DateTime(weekStart.Year, weekStart.Month, 1)
                : new DateTime(weekEnd.Year, weekEnd.Month, 1);
        }

        // Calculate week number within the target month
        DateTime firstDayOfTargetMonth = new DateTime(targetMonth.Year, targetMonth.Month, 1);
        DateTime firstReportDayOfMonth = GetFirstReportDayOfMonth(firstDayOfTargetMonth);

        if (firstReportDayOfMonth > weekEnd)
        {
            return (targetMonth, 1);
        }

        // Iterate through all weeks in the month and count only those that belong to the month
        DateTime lastDayOfTargetMonth = firstDayOfTargetMonth.AddMonths(1).AddDays(-1);
        int weekNumber = 0;
        DateTime currentWeekEnd = firstReportDayOfMonth;

        while (currentWeekEnd <= lastDayOfTargetMonth && currentWeekEnd <= weekEnd)
        {
            DateTime currentWeekStart = currentWeekEnd.AddDays(-6);

            int daysInTargetMonth = 0;
            for (DateTime day = currentWeekStart; day <= currentWeekEnd; day = day.AddDays(1))
            {
                if (day.Month == targetMonth.Month && day.Year == targetMonth.Year)
                {
                    daysInTargetMonth++;
                }
            }

            if (daysInTargetMonth >= 4)
            {
                weekNumber++;

                if (currentWeekEnd >= weekEnd)
                {
                    weekNumber = Math.Max(Constants.MinWeekNumber, Math.Min(Constants.MaxWeekNumber, weekNumber));
                    return (targetMonth, weekNumber);
                }
            }

            currentWeekEnd = currentWeekEnd.AddDays(7);
        }

        // Fallback
        int fallbackWeekNum = (int)Math.Ceiling((weekEnd - firstReportDayOfMonth).TotalDays / 7) + 1;
        fallbackWeekNum = Math.Max(Constants.MinWeekNumber, Math.Min(Constants.MaxWeekNumber, fallbackWeekNum));

        Console.WriteLine($"{Constants.ConsolePrefixReportDateCalculator} Warning: Could not determine week number by counting. Using fallback: {fallbackWeekNum}");
        return (targetMonth, fallbackWeekNum);
    }

    /// <summary>
    /// Gets the first report day (report end day) of the specified month.
    /// </summary>
    /// <param name="firstDayOfMonth">The first day of the month.</param>
    /// <returns>The first occurrence of the report end day in that month.</returns>
    private DateTime GetFirstReportDayOfMonth(DateTime firstDayOfMonth)
    {
        return firstDayOfMonth.AddDays((7 + (int)_reportEndDay - (int)firstDayOfMonth.DayOfWeek) % 7);
    }
}


