using ClosedXML.Excel;

namespace ExcelGenerator;

/// <summary>
/// Configuration builder for Excel generation with advanced features
/// </summary>
/// <typeparam name="T">The type of objects in the collection</typeparam>
public class ExcelConfiguration<T>
{
    internal AggregationType Aggregations { get; private set; } = AggregationType.Sum;
    internal bool ExcludeIds { get; private set; }
    internal XLColor HeaderColor { get; private set; } = XLColor.LightBlue;
    internal ConditionalFormattingConfiguration? ConditionalFormatting { get; private set; }
    internal int FreezeRowCount { get; private set; }
    internal int FreezeColumnCount { get; private set; }
    internal IEnumerable<T> Data { get; private set; } = Enumerable.Empty<T>();
    internal string SheetName { get; private set; } = "Sheet1";

    /// <summary>
    /// Sets the data to be exported
    /// </summary>
    /// <param name="data">The collection of data</param>
    /// <param name="sheetName">The name of the worksheet</param>
    /// <returns>The configuration for chaining</returns>
    public ExcelConfiguration<T> WithData(IEnumerable<T> data, string sheetName)
    {
        Data = data;
        SheetName = sheetName;
        return this;
    }

    /// <summary>
    /// Specifies which aggregations to calculate for numeric columns
    /// </summary>
    /// <param name="aggregations">The aggregation types (can be combined with | operator)</param>
    /// <returns>The configuration for chaining</returns>
    public ExcelConfiguration<T> WithAggregations(AggregationType aggregations)
    {
        Aggregations = aggregations;
        return this;
    }

    /// <summary>
    /// Excludes columns that end with "Id" or "ID"
    /// </summary>
    /// <returns>The configuration for chaining</returns>
    public ExcelConfiguration<T> WithExcludeIds()
    {
        ExcludeIds = true;
        return this;
    }

    /// <summary>
    /// Sets the header row background color
    /// </summary>
    /// <param name="color">The color for the header row</param>
    /// <returns>The configuration for chaining</returns>
    public ExcelConfiguration<T> WithHeaderColor(XLColor color)
    {
        HeaderColor = color;
        return this;
    }

    /// <summary>
    /// Configures conditional formatting rules
    /// </summary>
    /// <param name="configure">Action to configure conditional formatting</param>
    /// <returns>The configuration for chaining</returns>
    public ExcelConfiguration<T> WithConditionalFormatting(Action<ConditionalFormattingConfiguration> configure)
    {
        var config = new ConditionalFormattingConfiguration();
        configure(config);
        ConditionalFormatting = config;
        return this;
    }

    /// <summary>
    /// Freezes the header row (first row)
    /// </summary>
    /// <returns>The configuration for chaining</returns>
    public ExcelConfiguration<T> FreezeHeaderRow()
    {
        FreezeRowCount = 1;
        return this;
    }

    /// <summary>
    /// Freezes specific rows and columns
    /// </summary>
    /// <param name="rowsToFreeze">Number of rows to freeze from the top</param>
    /// <param name="columnsToFreeze">Number of columns to freeze from the left</param>
    /// <returns>The configuration for chaining</returns>
    public ExcelConfiguration<T> FreezePanes(int rowsToFreeze, int columnsToFreeze = 0)
    {
        FreezeRowCount = rowsToFreeze;
        FreezeColumnCount = columnsToFreeze;
        return this;
    }

    /// <summary>
    /// Generates the Excel workbook with the configured settings
    /// </summary>
    /// <returns>The generated workbook</returns>
    public XLWorkbook GenerateExcel()
    {
        return ExcelSheetGenerator.GenerateExcel(Data, SheetName, this);
    }

    /// <summary>
    /// Generates an Excel file and saves it to the specified path
    /// </summary>
    /// <param name="filePath">The path where the Excel file will be saved</param>
    public void GenerateExcelFile(string filePath)
    {
        using var workbook = GenerateExcel();
        workbook.SaveAs(filePath);
    }

    /// <summary>
    /// Generates an Excel file and returns it as a byte array
    /// </summary>
    /// <returns>A byte array containing the Excel file</returns>
    public byte[] GenerateExcelBytes()
    {
        using var workbook = GenerateExcel();
        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }

    /// <summary>
    /// Generates an Excel file and returns it as a Stream
    /// </summary>
    /// <returns>A MemoryStream containing the Excel file</returns>
    public MemoryStream GenerateExcelStream()
    {
        using var workbook = GenerateExcel();
        var stream = new MemoryStream();
        workbook.SaveAs(stream);
        stream.Position = 0;
        return stream;
    }
}
