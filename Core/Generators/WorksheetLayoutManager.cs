using ClosedXML.Excel;

namespace ExcelGenerator.Core.Generators;

/// <summary>
/// Manages worksheet layout settings (freeze panes, auto-fit)
/// Single responsibility: Worksheet layout configuration
/// </summary>
internal class WorksheetLayoutManager
{
    /// <summary>
    /// Applies layout settings to the worksheet
    /// </summary>
    public void ApplyLayout(IXLWorksheet worksheet, int freezeRowCount, int freezeColumnCount)
    {
        // Validate inputs
        if (worksheet == null)
            throw new ArgumentNullException(nameof(worksheet), "Worksheet cannot be null.");
        if (freezeRowCount < 0)
            throw new ArgumentOutOfRangeException(nameof(freezeRowCount), "Freeze row count cannot be negative.");
        if (freezeColumnCount < 0)
            throw new ArgumentOutOfRangeException(nameof(freezeColumnCount), "Freeze column count cannot be negative.");

        // Apply freeze panes
        if (freezeRowCount > 0 || freezeColumnCount > 0)
        {
            worksheet.SheetView.FreezeRows(freezeRowCount);
            worksheet.SheetView.FreezeColumns(freezeColumnCount);
        }

        // Auto-fit columns
        worksheet.Columns().AdjustToContents();
    }
}
