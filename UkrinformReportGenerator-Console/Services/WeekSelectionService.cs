using System;
using URG_Console.Interfaces;

namespace URG_Console.Services;

/// <summary>
/// Service for selecting and calculating operating weeks.
/// </summary>
public class WeekSelectionService : IWeekSelectionService
{
    /// <inheritdoc/>
    public (DateTime startDate, DateTime endDate) SelectOperatingWeek(DayOfWeek reportEndDay)
    {
        DateTime today = DateTime.Today;
        DateTime currentWeekEnd = GetCurrentWeekEndDay(today, reportEndDay);
        DateTime currentWeekStart = currentWeekEnd.AddDays(-6);
        DateTime previousWeekEnd = currentWeekEnd.AddDays(-7);
        DateTime previousWeekStart = previousWeekEnd.AddDays(-6);

        Console.WriteLine("Select operating week:");
        Console.WriteLine($"1. Current week ({currentWeekStart:dd.MM.yyyy} - {currentWeekEnd:dd.MM.yyyy})");
        Console.WriteLine($"2. Previous week ({previousWeekStart:dd.MM.yyyy} - {previousWeekEnd:dd.MM.yyyy})");

        while (true)
        {
            Console.Write("> ");
            string? input = Console.ReadLine();

            if (int.TryParse(input, out int choice))
            {
                (DateTime startDate, DateTime endDate) result = choice switch
                {
                    1 => (currentWeekStart, currentWeekEnd),
                    2 => (previousWeekStart, previousWeekEnd),
                    _ => default
                };

                if (result.startDate != default || result.endDate != default || choice == 1 || choice == 2)
                {
                    return result;
                }
                Console.WriteLine("Invalid input! Try again");
            }
            else
            {
                Console.WriteLine("Invalid input! Try again");
            }
        }
    }

    /// <inheritdoc/>
    public DateTime GetCurrentWeekEndDay(DateTime start, DayOfWeek endDay)
    {
        int daysUntilEndDay = ((int)endDay - (int)start.DayOfWeek + 7) % 7;
        return start.AddDays(daysUntilEndDay);
    }
}

