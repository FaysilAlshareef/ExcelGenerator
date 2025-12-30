using ClosedXML.Excel;

namespace ExcelGenerator.Core.CellFormatters;

/// <summary>
/// Handles null values by setting cell to empty string
/// </summary>
internal class NullValueFormatter : ICellValueFormatter
{
    public bool CanFormat(Type type)
    {
        // This formatter is used explicitly for null values, not based on type
        return false;
    }

    public void Format(IXLCell cell, object? value, Type type)
    {
        cell.Value = string.Empty;
    }

    public int Priority => 100; // Highest priority to check nulls first
}
