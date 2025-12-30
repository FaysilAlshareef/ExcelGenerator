using ClosedXML.Excel;

namespace ExcelGenerator.Core.CellFormatters;

/// <summary>
/// Formats boolean values as "Yes" or "No"
/// </summary>
internal class BooleanFormatter : ICellValueFormatter
{
    public bool CanFormat(Type type)
    {
        return type == typeof(bool);
    }

    public void Format(IXLCell cell, object? value, Type type)
    {
        if (value == null)
        {
            cell.Value = string.Empty;
            return;
        }

        cell.Value = (bool)value ? "Yes" : "No";
    }

    public int Priority => 10;
}
