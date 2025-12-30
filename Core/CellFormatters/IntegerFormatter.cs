using ClosedXML.Excel;

namespace ExcelGenerator.Core.CellFormatters;

/// <summary>
/// Formats integer types (int, long, short, byte) with thousand separators
/// </summary>
internal class IntegerFormatter : ICellValueFormatter
{
    public bool CanFormat(Type type)
    {
        return type == typeof(int) || type == typeof(long) ||
               type == typeof(short) || type == typeof(byte);
    }

    public void Format(IXLCell cell, object? value, Type type)
    {
        if (value == null)
        {
            cell.Value = string.Empty;
            return;
        }

        cell.Value = Convert.ToDouble(value);
        cell.Style.NumberFormat.Format = "#,##0";
    }

    public int Priority => 10;
}
