namespace URG_Console;

/// <summary>
/// Contains calculated report date information used for file naming and header generation.
/// </summary>
public class ReportDateInfo
{
    /// <summary>
    /// Gets or sets the day number of the week start date.
    /// </summary>
    public int WeekStartDay { get; set; }

    /// <summary>
    /// Gets or sets the day number of the week end date.
    /// </summary>
    public int WeekEndDay { get; set; }

    /// <summary>
    /// Gets or sets the month number for file naming (based on week ownership).
    /// </summary>
    public int FileMonth { get; set; }

    /// <summary>
    /// Gets or sets the year for file naming (based on week ownership).
    /// </summary>
    public int FileYear { get; set; }

    /// <summary>
    /// Gets or sets the week number within the file month (1-5).
    /// </summary>
    public int WeekNumber { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether the report spans two months.
    /// </summary>
    public bool IsReportSpanningTwoMonths { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether the report spans two years.
    /// </summary>
    public bool IsReportSpanningTwoYears { get; set; }

    /// <summary>
    /// Gets or sets the start month number (for header generation).
    /// </summary>
    public int StartMonth { get; set; }

    /// <summary>
    /// Gets or sets the start year (for header generation).
    /// </summary>
    public int StartYear { get; set; }

    /// <summary>
    /// Gets or sets the end month number (for header generation).
    /// </summary>
    public int EndMonth { get; set; }

    /// <summary>
    /// Gets or sets the end year (for header generation).
    /// </summary>
    public int EndYear { get; set; }
}

