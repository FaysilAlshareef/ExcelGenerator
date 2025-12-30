using ClosedXML.Excel;
using System.Reflection;
using ExcelGenerator.Core.CellFormatters;

namespace ExcelGenerator.Core.Generators;

/// <summary>
/// Generates data rows in Excel worksheets
/// Single responsibility: Data row creation
/// </summary>
internal class DataRowGenerator
{
    private readonly CellFormatterFactory _cellFormatterFactory;

    public DataRowGenerator(CellFormatterFactory cellFormatterFactory)
    {
        _cellFormatterFactory = cellFormatterFactory;
    }

    /// <summary>
    /// Generates all data rows and returns the count of rows written
    /// </summary>
    public int Generate<T>(IXLWorksheet worksheet, List<T> dataList, PropertyInfo[] properties)
    {
        // Validate inputs
        if (worksheet == null)
            throw new ArgumentNullException(nameof(worksheet), "Worksheet cannot be null.");
        if (dataList == null)
            throw new ArgumentNullException(nameof(dataList), "Data list cannot be null.");
        if (properties == null)
            throw new ArgumentNullException(nameof(properties), "Properties array cannot be null.");

        for (int rowIndex = 0; rowIndex < dataList.Count; rowIndex++)
        {
            var item = dataList[rowIndex];
            if (item == null) continue;

            for (int colIndex = 0; colIndex < properties.Length; colIndex++)
            {
                var cell = worksheet.Cell(rowIndex + 2, colIndex + 1);
                var value = properties[colIndex].GetValue(item);

                _cellFormatterFactory.FormatCell(cell, value, properties[colIndex].PropertyType);
                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            }
        }

        return dataList.Count;
    }
}
