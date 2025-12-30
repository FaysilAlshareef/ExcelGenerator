using ClosedXML.Excel;
using System.Reflection;
using ExcelGenerator.Core.CellFormatters;
using ExcelGenerator.Core.Aggregation;
using ExcelGenerator.Core.ConditionalFormatting;
using ExcelGenerator.Core.PropertyReflection;
using ExcelGenerator.Core.Generators;
using ExcelGenerator.Core;

namespace ExcelGenerator;

/// <summary>
/// Generates Excel sheets from IEnumerable collections
/// </summary>
public static class ExcelSheetGenerator
{
    // Lazy-initialized engine for coordinating all Excel generation (Facade pattern)
    private static readonly Lazy<ExcelGeneratorEngine> _engine =
        new Lazy<ExcelGeneratorEngine>(CreateEngine);

    /// <summary>
    /// Creates and wires up the ExcelGeneratorEngine with all dependencies
    /// Manual dependency injection without external DI framework
    /// </summary>
    private static ExcelGeneratorEngine CreateEngine()
    {
        // Create all dependencies
        var propertyExtractor = new PropertyExtractor();
        var cellFormatterFactory = new CellFormatterFactory();
        var aggregationFactory = new AggregationStrategyFactory();
        var formattingFactory = new FormattingRuleApplierFactory();

        // Create specialized generators
        var headerGenerator = new HeaderGenerator(propertyExtractor);
        var dataRowGenerator = new DataRowGenerator(cellFormatterFactory);
        var aggregationGenerator = new AggregationRowGenerator(aggregationFactory);
        var layoutManager = new WorksheetLayoutManager();

        // Wire up the engine
        return new ExcelGeneratorEngine(
            propertyExtractor,
            headerGenerator,
            dataRowGenerator,
            aggregationGenerator,
            formattingFactory,
            layoutManager);
    }

    /// <summary>
    /// Creates a new Excel configuration for advanced features
    /// </summary>
    /// <typeparam name="T">The type of objects in the collection</typeparam>
    /// <returns>A new ExcelConfiguration instance for fluent configuration</returns>
    public static ExcelConfiguration<T> Configure<T>()
    {
        return new ExcelConfiguration<T>();
    }
    /// <summary>
    /// Generates an Excel workbook from a collection of objects
    /// </summary>
    /// <typeparam name="T">The type of objects in the collection</typeparam>
    /// <param name="data">The collection of data to export</param>
    /// <param name="sheetName">The name of the worksheet</param>
    /// <param name="excludeIds">If true, excludes columns that end with "Id"</param>
    /// <param name="headerColor">The background color for header cells</param>
    /// <returns>An XLWorkbook containing the generated Excel sheet</returns>
    public static XLWorkbook GenerateExcel<T>(
        IEnumerable<T> data,
        string sheetName,
        bool excludeIds = false,
        XLColor? headerColor = null)
    {
        // Create configuration for backward compatibility
        var config = new ExcelConfiguration<T>()
            .WithHeaderColor(headerColor ?? XLColor.LightBlue)
            .WithAggregations(AggregationType.Sum);  // Old behavior was to add summation row

        if (excludeIds)
            config.WithExcludeIds();

        // Delegate to engine (Facade pattern)
        return _engine.Value.Generate(data, sheetName, config);
    }

    /// <summary>
    /// Generates an Excel file and saves it to the specified path
    /// </summary>
    /// <typeparam name="T">The type of objects in the collection</typeparam>
    /// <param name="data">The collection of data to export</param>
    /// <param name="sheetName">The name of the worksheet</param>
    /// <param name="filePath">The path where the Excel file will be saved</param>
    /// <param name="excludeIds">If true, excludes columns that end with "Id"</param>
    /// <param name="headerColor">The background color for header cells</param>
    public static void GenerateExcelFile<T>(
        IEnumerable<T> data,
        string sheetName,
        string filePath,
        bool excludeIds = false,
        XLColor? headerColor = null)
    {
        using var workbook = GenerateExcel(data, sheetName, excludeIds, headerColor);
        workbook.SaveAs(filePath);
    }

    /// <summary>
    /// Generates an Excel file and returns it as a byte array
    /// </summary>
    /// <typeparam name="T">The type of objects in the collection</typeparam>
    /// <param name="data">The collection of data to export</param>
    /// <param name="sheetName">The name of the worksheet</param>
    /// <param name="excludeIds">If true, excludes columns that end with "Id"</param>
    /// <param name="headerColor">The background color for header cells</param>
    /// <returns>A byte array containing the Excel file</returns>
    public static byte[] GenerateExcelBytes<T>(
        IEnumerable<T> data,
        string sheetName,
        bool excludeIds = false,
        XLColor? headerColor = null)
    {
        using var workbook = GenerateExcel(data, sheetName, excludeIds, headerColor);
        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }

    /// <summary>
    /// Generates an Excel file and returns it as a Stream
    /// </summary>
    /// <typeparam name="T">The type of objects in the collection</typeparam>
    /// <param name="data">The collection of data to export</param>
    /// <param name="sheetName">The name of the worksheet</param>
    /// <param name="excludeIds">If true, excludes columns that end with "Id"</param>
    /// <param name="headerColor">The background color for header cells</param>
    /// <returns>A MemoryStream containing the Excel file</returns>
    public static MemoryStream GenerateExcelStream<T>(
        IEnumerable<T> data,
        string sheetName,
        bool excludeIds = false,
        XLColor? headerColor = null)
    {
        using var workbook = GenerateExcel(data, sheetName, excludeIds, headerColor);
        var stream = new MemoryStream();
        workbook.SaveAs(stream);
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// Generates an Excel workbook from a collection using advanced configuration
    /// </summary>
    /// <typeparam name="T">The type of objects in the collection</typeparam>
    /// <param name="data">The collection of data to export</param>
    /// <param name="sheetName">The name of the worksheet</param>
    /// <param name="configuration">The configuration for Excel generation</param>
    /// <returns>An XLWorkbook containing the generated Excel sheet</returns>
    internal static XLWorkbook GenerateExcel<T>(
        IEnumerable<T> data,
        string sheetName,
        ExcelConfiguration<T> configuration)
    {
        // Delegate to engine (Facade pattern)
        return _engine.Value.Generate(data, sheetName, configuration);
    }

}
