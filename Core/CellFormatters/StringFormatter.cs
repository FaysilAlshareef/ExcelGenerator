using ClosedXML.Excel;

namespace ExcelGenerator.Core.CellFormatters;

/// <summary>
/// Fallback formatter for any type - uses ToString()
/// </summary>
internal class StringFormatter : ICellValueFormatter
{
    public bool CanFormat(Type type)
    {
        // This is the fallback formatter, so it can handle any type
        return true;
    }

    public void Format(IXLCell cell, object? value, Type type)
    {
        if (value == null)
        {
            cell.Value = string.Empty;
            return;
        }

        cell.Value = value.ToString() ?? string.Empty;
    }

    public int Priority => 0; // Lowest priority - fallback formatter
}
