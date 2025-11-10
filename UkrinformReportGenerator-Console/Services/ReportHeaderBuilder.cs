using System;

namespace URG_Console.Services;

/// <summary>
/// Service for building report headers.
/// </summary>
public class ReportHeaderBuilder
{
    private enum Month
    {
        січня = 1, лютого = 2, березня = 3, квітня = 4, травня = 5, червня = 6,
        липня = 7, серпня = 8, вересня = 9, жовтня = 10, листопада = 11, грудня = 12
    }

    /// <summary>
    /// Builds the report header string based on date information.
    /// The header uses actual report dates (not week ownership) for accurate date ranges.
    /// </summary>
    /// <param name="dateInfo">The date information containing start/end dates and spanning flags.</param>
    /// <returns>The formatted header string with author name and date range.</returns>
    public string BuildHeader(ReportDateInfo dateInfo)
    {
        try
        {
            Month startMonth = (Month)dateInfo.StartMonth;
            Month endMonth = (Month)dateInfo.EndMonth;

            string dateRange;
            if (dateInfo.IsReportSpanningTwoYears)
            {
                dateRange = $"з {dateInfo.WeekStartDay} {startMonth} {dateInfo.StartYear} року по {dateInfo.WeekEndDay} {endMonth} {dateInfo.EndYear} року";
            }
            else if (dateInfo.IsReportSpanningTwoMonths)
            {
                dateRange = $"з {dateInfo.WeekStartDay} {startMonth} по {dateInfo.WeekEndDay} {endMonth} {dateInfo.StartYear} року";
            }
            else
            {
                dateRange = $"з {dateInfo.WeekStartDay} по {dateInfo.WeekEndDay} {startMonth} {dateInfo.StartYear} року";
            }

            return $"Автор - Ярослав Довгопол{Environment.NewLine}Публікації, що вийшли в період {dateRange}{Environment.NewLine}";
        }
        catch (Exception ex)
        {
            Console.WriteLine($"{Constants.ConsolePrefixReportHeaderBuilder} Error setting header: {ex.Message}");
            Console.WriteLine("Need to set header manually in the report");
            return Constants.DefaultHeaderNotSet + Environment.NewLine;
        }
    }
}

