using ClosedXML.Excel;

namespace ExcelGenerator.Core.CellFormatters;

/// <summary>
/// Formats DateOnly values with date-only format
/// </summary>
internal class DateOnlyFormatter : ICellValueFormatter
{
    public bool CanFormat(Type type)
    {
        return type == typeof(DateOnly);
    }

    public void Format(IXLCell cell, object? value, Type type)
    {
        if (value == null)
        {
            cell.Value = string.Empty;
            return;
        }

        var dateOnly = (DateOnly)value;
        cell.Value = dateOnly.ToDateTime(TimeOnly.MinValue);
        cell.Style.DateFormat.Format = "yyyy-MM-dd";
    }

    public int Priority => 10;
}
