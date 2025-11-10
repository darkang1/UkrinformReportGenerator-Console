# Ukrinform Report Generator - Console Application

A .NET 8.0 console application that automatically generates weekly reports from Word documents containing Ukrinform article links. The application parses articles from web sources, extracts metadata, and generates formatted Word document reports.

## Features

- **Automatic Document Parsing**: Extracts article links from Word documents (.docx, .doc) with intelligent header area validation
- **Web Article Parsing**: Fetches and parses article content from Ukrinform website
- **Intelligent Week Calculation**: Uses 4-day rule to determine week ownership for accurate file naming
- **Report Generation**: Creates formatted Word documents with article metadata
- **Multi-Article Support**: Handles documents with multiple articles (1+N pattern)
- **Async I/O Operations**: Non-blocking file and network operations for better performance
- **Dependency Injection**: Clean architecture with testable services

## Requirements

- **.NET 8.0 SDK** or later
- **Microsoft Word** (for opening generated reports)
- **Internet Connection** (for fetching article content)

## Installation

1. Clone the repository:
   ```bash
   git clone <repository-url>
   cd UkrinformReportGenerator-Console
   ```

2. Restore NuGet packages:
   ```bash
   dotnet restore
   ```

3. Build the application:
   ```bash
   dotnet build
   ```

4. Run the application:
   ```bash
   dotnet run
   ```

## Configuration

Create or edit `appsettings.json` in the project root:

```json
{
  "DefaultPath": "C:\\Users\\YourName\\Desktop\\Weekly",
  "WordExecutablePath": "C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE"
}
```

### Configuration Options

- **DefaultPath**: Default folder path where source documents are located
- **WordExecutablePath**: Path to Microsoft Word executable (optional, has default)

## Usage

1. **Start the application**:
   ```bash
   dotnet run
   ```

2. **Select operating week**:
   - Choose between current week or previous week
   - The application calculates week boundaries based on Sunday as the week end day

3. **Select folder location**:
   - Use default path from configuration, or
   - Manually specify folder path containing Word documents

4. **Wait for processing**:
   - Application parses documents
   - Fetches article content from web
   - Displays parsed articles in console (with clickable file names)
   - Generates Word report

5. **Report is generated**:
   - Report file is saved in the selected folder
   - File name format: `AUTO_Dovgopol_YYYY_MM_WEEK=COUNT.docx`
   - Example: `AUTO_Dovgopol_2025_11_II=60.docx`
   - Word automatically opens the generated report

## File Naming Convention

Reports are named using the following format:
```
AUTO_Dovgopol_{YEAR}_{MONTH}_{WEEK}={ARTICLE_COUNT}.docx
```

Where:
- **YEAR**: 4-digit year (e.g., 2025)
- **MONTH**: 2-digit month (01-12)
- **WEEK**: Roman numeral (I, II, III, IV, V) representing week number in the month
- **ARTICLE_COUNT**: Number of articles in the report

### Week Number Calculation

The application uses a **4-day rule** to determine which month a week belongs to:
- A week belongs to a month if **at least 4 days** of that week are within that month
- Week numbers are calculated within the target month, counting only weeks that satisfy the 4-day rule

**Example**: If a week spans Jan 29 - Feb 4:
- 3 days in January, 4 days in February
- Week belongs to **February** (4 days ≥ 4)
- File name uses February month and calculates week number within February

## Project Structure

```
UkrinformReportGenerator-Console/
├── Configuration/
│   └── AppSettings.cs          # Application configuration
├── Interfaces/
│   ├── IArticleParser.cs      # Article parsing interface
│   ├── IDocumentParser.cs     # Document parsing interface
│   ├── IReportGenerator.cs    # Report generation interface
│   ├── IWeekSelectionService.cs
│   └── IPathValidationService.cs
├── Models/
│   ├── ArticleSource.cs       # Article source data model
│   ├── WebParser.cs           # Article data model
│   └── ReportDateInfo.cs      # Report date information model
├── Services/
│   ├── ArticleParser.cs       # Web article parsing service
│   ├── DocumentParser.cs      # Word document parsing service
│   ├── ReportGenerator.cs     # Main orchestrator
│   ├── ReportDateCalculator.cs # Date/week calculation service
│   ├── ReportHeaderBuilder.cs  # Header generation service
│   ├── WordReportBuilder.cs   # Word document generation service
│   ├── ArticleDisplayService.cs
│   ├── WeekSelectionService.cs
│   └── PathValidationService.cs
├── Constants.cs                # Application constants
├── Program.cs                  # Application entry point
└── appsettings.json            # Configuration file
```

## Architecture

The application follows **Clean Architecture** principles:

- **Separation of Concerns**: Each service has a single responsibility
- **Dependency Injection**: All services are injected via constructor
- **Interface-Based Design**: Services implement interfaces for testability
- **Async/Await**: I/O operations are asynchronous for better performance

See [ARCHITECTURE.md](ARCHITECTURE.md) for detailed architecture documentation.

## Article Types

Articles are classified based on character count (non-whitespace):

- **Інф. повідомлення** (Information Message): 1-1,399 characters
- **Розш. інф. повідомлення** (Extended Information Message): 1,400-4,999 characters
- **Коментар** (Comment): 5,000+ characters
- **Error**: Invalid or unparseable articles

## Report Format

Generated reports include:

1. **Header**: Author name and date range
2. **Table** with columns:
   - №з/п (Row number)
   - Дата (Publication date)
   - Заголовок (Article title with hyperlink)
   - Жанр (Article type)
   - Кільк. знаків (Character count)

**Special Formatting**:
- Comment articles are highlighted in yellow
- Exclusive articles have "Ексклюзив" prefix highlighted in yellow
- Article titles are hyperlinked to source URLs

## Multi-Article Files

Documents can contain multiple articles. Use the naming pattern:
```
filename_1+N_rest.docx
```

Where `N` is the number of additional articles (e.g., `1+2` means 3 total articles).

### Document Structure Requirements

The application uses intelligent header area validation to correctly identify article hyperlinks:

- **Header Area**: Article hyperlinks must **start** within the first **150 characters** of the document
- **Validation Rule**: If a hyperlink's first character is within the 150-character limit, it's accepted (even if the hyperlink extends beyond)
- **Document Reading Order**: The parser uses paragraph iteration order (document reading order), not hyperlink collection order
- **Structure-Agnostic**: Works regardless of preceding content (file names, metadata, etc.)
- **Boundary Handling**: Only examines paragraph text within the limit, properly handling paragraphs that span the boundary

**Example**: If your document has:
```
Paragraph 1: "US_ABC_News_Washington" (22 chars)
Paragraph 2: "Article Title [hyperlink]" (120 chars)
```

The parser will correctly identify the article hyperlink in paragraph 2, even if:
- Other hyperlinks exist later in the document
- The hyperlink text extends slightly beyond 150 characters
- The paragraph spans the 150-character boundary

As long as the hyperlink's **first character** is within the first 150 characters, it will be accepted.

## Console Output Features

### Clickable File Names

When articles are displayed in the console, file names are shown as clickable hyperlinks:
- **Display**: Shows just the filename (e.g., `article.docx`)
- **Clickable**: Clicking the filename opens the file in its default application
- **Terminal Support**: Works in Windows Terminal, VS Code terminal, and other modern terminals
- **Fallback**: In unsupported terminals, filenames display as plain text

**Example Output**:
```
Article file name: article.docx
```
(The filename `article.docx` is clickable and opens the file when clicked)

## Error Handling

The application handles various error scenarios:

- **Missing Links**: Creates error articles with default values
- **Network Failures**: Logs errors and continues processing
- **Invalid Documents**: Skips problematic files with error messages
- **File Locks**: Kills Word process if report file is locked
- **Invalid Paths**: Validates paths before processing

## Troubleshooting

### Word Executable Not Found

If you see: `Warning: Word executable not found at '...'`

**Solution**: Update `WordExecutablePath` in `appsettings.json` with the correct path to WINWORD.EXE

### Report File is Locked

If you see: `Generated reports cannot be open during the runtime`

**Solution**: Close any open Word documents and try again

### No Articles Found

If parsing returns no articles:

1. Check that Word documents contain Ukrinform links
2. Verify links are valid and accessible
3. Check internet connection
4. Review console output for specific error messages

### Incorrect Week Number

If week numbers seem incorrect:

1. Verify the week spans the correct date range
2. Check that the 4-day rule is being applied correctly
3. Review `ReportDateCalculator` logic for edge cases

## Development

### Building from Source

```bash
# Restore dependencies
dotnet restore

# Build in Release mode
dotnet build -c Release

# Run tests (if available)
dotnet test
```

### Code Style

- Uses **file-scoped namespaces** (C# 10+)
- **Nullable reference types** enabled
- **Async/await** for all I/O operations
- **XML documentation** for all public APIs

### Dependencies

- **DocX 5.0.0**: Word document generation
- **HtmlAgilityPack 1.12.4**: HTML parsing
- **Microsoft.Extensions.Configuration 9.0.10**: Configuration management
- **Microsoft.Extensions.Configuration.Json 9.0.10**: JSON configuration provider
- **Microsoft.Extensions.DependencyInjection 9.0.10**: Dependency injection
- **System.IO.Packaging 9.0.10**: Package support

## Author

**Bogdan Dovgopol**

---

**Version**: 2.0  
**Last Updated**: 2025 
**Target Framework**: .NET 8.0



