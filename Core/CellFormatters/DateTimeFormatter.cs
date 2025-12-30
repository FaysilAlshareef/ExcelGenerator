using ClosedXML.Excel;

namespace ExcelGenerator.Core.CellFormatters;

/// <summary>
/// Formats DateTime values with standard date-time format
/// </summary>
internal class DateTimeFormatter : ICellValueFormatter
{
    public bool CanFormat(Type type)
    {
        return type == typeof(DateTime);
    }

    public void Format(IXLCell cell, object? value, Type type)
    {
        if (value == null)
        {
            cell.Value = string.Empty;
            return;
        }

        cell.Value = (DateTime)value;
        cell.Style.DateFormat.Format = "yyyy-MM-dd HH:mm:ss";
    }

    public int Priority => 10;
}
