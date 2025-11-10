namespace URG_Console.Configuration
{
    /// <summary>
    /// Application settings configuration class.
    /// </summary>
    public class AppSettings
    {
        /// <summary>
        /// Default folder path for report generation.
        /// </summary>
        public string DefaultPath { get; set; } = string.Empty;

        /// <summary>
        /// Path to Microsoft Word executable.
        /// </summary>
        public string WordExecutablePath { get; set; } = @"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE";

        /// <summary>
        /// Gets the default path or a fallback message if not specified.
        /// </summary>
        public string GetDefaultPathOrDefault()
        {
            return string.IsNullOrWhiteSpace(DefaultPath)
                ? "Not specified in appsettings.json"
                : DefaultPath;
        }
    }
}



