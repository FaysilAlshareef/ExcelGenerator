using ClosedXML.Excel;

namespace ExcelGenerator;

/// <summary>
/// Builder for creating Excel workbooks with multiple sheets
/// </summary>
public class ExcelWorkbookBuilder
{
    private readonly XLWorkbook _workbook = new();
    private readonly List<SheetConfiguration> _sheets = new();

    /// <summary>
    /// Adds a sheet to the workbook
    /// </summary>
    /// <typeparam name="T">The type of objects in the collection</typeparam>
    /// <param name="sheetName">The name of the worksheet</param>
    /// <param name="data">The collection of data to export</param>
    /// <param name="configure">Optional action to configure the sheet</param>
    /// <returns>The builder for chaining</returns>
    public ExcelWorkbookBuilder AddSheet<T>(
        string sheetName,
        IEnumerable<T> data,
        Action<ExcelConfiguration<T>>? configure = null)
    {
        var config = new ExcelConfiguration<T>().WithData(data, sheetName);
        configure?.Invoke(config);

        _sheets.Add(new SheetConfiguration
        {
            SheetName = sheetName,
            Generator = () => ExcelSheetGenerator.GenerateExcel(data, sheetName, config)
        });

        return this;
    }

    /// <summary>
    /// Builds the complete workbook with all configured sheets
    /// </summary>
    /// <returns>The generated workbook</returns>
    public XLWorkbook Build()
    {
        // If no sheets were added, return empty workbook
        if (_sheets.Count == 0)
            return _workbook;

        // Generate all sheets and copy them to the workbook
        foreach (var sheet in _sheets)
        {
            using var tempWorkbook = sheet.Generator();
            var sourceWorksheet = tempWorkbook.Worksheets.First();

            // Copy worksheet to our workbook
            sourceWorksheet.CopyTo(_workbook, sheet.SheetName);
        }

        return _workbook;
    }

    /// <summary>
    /// Builds the workbook and saves it to a file
    /// </summary>
    /// <param name="filePath">The path where the Excel file will be saved</param>
    public void SaveAs(string filePath)
    {
        using var workbook = Build();
        workbook.SaveAs(filePath);
    }

    /// <summary>
    /// Builds the workbook and returns it as a byte array
    /// </summary>
    /// <returns>A byte array containing the Excel file</returns>
    public byte[] ToBytes()
    {
        using var workbook = Build();
        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }

    /// <summary>
    /// Builds the workbook and returns it as a Stream
    /// </summary>
    /// <returns>A MemoryStream containing the Excel file</returns>
    public MemoryStream ToStream()
    {
        using var workbook = Build();
        var stream = new MemoryStream();
        workbook.SaveAs(stream);
        stream.Position = 0;
        return stream;
    }
}

internal class SheetConfiguration
{
    public required string SheetName { get; set; }
    public required Func<XLWorkbook> Generator { get; set; }
}
