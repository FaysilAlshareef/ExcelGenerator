using ClosedXML.Excel;

namespace ExcelGenerator.Core.CellFormatters;

/// <summary>
/// Formats decimal, double, and float values with two decimal places
/// </summary>
internal class DecimalFormatter : ICellValueFormatter
{
    public bool CanFormat(Type type)
    {
        return type == typeof(decimal) || type == typeof(double) || type == typeof(float);
    }

    public void Format(IXLCell cell, object? value, Type type)
    {
        if (value == null)
        {
            cell.Value = string.Empty;
            return;
        }

        cell.Value = Convert.ToDouble(value);
        cell.Style.NumberFormat.Format = "#,##0.00";
    }

    public int Priority => 10;
}
