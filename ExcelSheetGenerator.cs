using ClosedXML.Excel;
using System.Reflection;

namespace ExcelGenerator;

/// <summary>
/// Generates Excel sheets from IEnumerable collections
/// </summary>
public static class ExcelSheetGenerator
{
    /// <summary>
    /// Generates an Excel workbook from a collection of objects
    /// </summary>
    /// <typeparam name="T">The type of objects in the collection</typeparam>
    /// <param name="data">The collection of data to export</param>
    /// <param name="sheetName">The name of the worksheet</param>
    /// <param name="excludeIds">If true, excludes columns that end with "Id"</param>
    /// <param name="headerColor">The background color for header cells</param>
    /// <returns>An XLWorkbook containing the generated Excel sheet</returns>
    public static XLWorkbook GenerateExcel<T>(
        IEnumerable<T> data,
        string sheetName,
        bool excludeIds = false,
        XLColor? headerColor = null)
    {
        var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add(sheetName);

        var properties = GetProperties<T>(excludeIds);

        if (properties.Length == 0)
            return workbook;

        // Add headers
        AddHeaders(worksheet, properties, headerColor ?? XLColor.LightBlue);

        // Add data rows
        var dataList = data.ToList();
        AddDataRows(worksheet, dataList, properties);

        // Add summation row for decimal columns
        AddSummationRow(worksheet, dataList, properties);

        // Auto-fit columns
        worksheet.Columns().AdjustToContents();

        return workbook;
    }

    /// <summary>
    /// Generates an Excel file and saves it to the specified path
    /// </summary>
    /// <typeparam name="T">The type of objects in the collection</typeparam>
    /// <param name="data">The collection of data to export</param>
    /// <param name="sheetName">The name of the worksheet</param>
    /// <param name="filePath">The path where the Excel file will be saved</param>
    /// <param name="excludeIds">If true, excludes columns that end with "Id"</param>
    /// <param name="headerColor">The background color for header cells</param>
    public static void GenerateExcelFile<T>(
        IEnumerable<T> data,
        string sheetName,
        string filePath,
        bool excludeIds = false,
        XLColor? headerColor = null)
    {
        using var workbook = GenerateExcel(data, sheetName, excludeIds, headerColor);
        workbook.SaveAs(filePath);
    }

    /// <summary>
    /// Generates an Excel file and returns it as a byte array
    /// </summary>
    /// <typeparam name="T">The type of objects in the collection</typeparam>
    /// <param name="data">The collection of data to export</param>
    /// <param name="sheetName">The name of the worksheet</param>
    /// <param name="excludeIds">If true, excludes columns that end with "Id"</param>
    /// <param name="headerColor">The background color for header cells</param>
    /// <returns>A byte array containing the Excel file</returns>
    public static byte[] GenerateExcelBytes<T>(
        IEnumerable<T> data,
        string sheetName,
        bool excludeIds = false,
        XLColor? headerColor = null)
    {
        using var workbook = GenerateExcel(data, sheetName, excludeIds, headerColor);
        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }

    /// <summary>
    /// Generates an Excel file and returns it as a Stream
    /// </summary>
    /// <typeparam name="T">The type of objects in the collection</typeparam>
    /// <param name="data">The collection of data to export</param>
    /// <param name="sheetName">The name of the worksheet</param>
    /// <param name="excludeIds">If true, excludes columns that end with "Id"</param>
    /// <param name="headerColor">The background color for header cells</param>
    /// <returns>A MemoryStream containing the Excel file</returns>
    public static MemoryStream GenerateExcelStream<T>(
        IEnumerable<T> data,
        string sheetName,
        bool excludeIds = false,
        XLColor? headerColor = null)
    {
        using var workbook = GenerateExcel(data, sheetName, excludeIds, headerColor);
        var stream = new MemoryStream();
        workbook.SaveAs(stream);
        stream.Position = 0;
        return stream;
    }

    private static PropertyInfo[] GetProperties<T>(bool excludeIds)
    {
        var properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(p => p.CanRead);

        if (excludeIds)
        {
            properties = properties.Where(p => 
                !p.Name.EndsWith("Id", StringComparison.OrdinalIgnoreCase) &&
                !p.Name.EndsWith("ID", StringComparison.Ordinal));
        }

        return properties.ToArray();
    }

    private static void AddHeaders(IXLWorksheet worksheet, PropertyInfo[] properties, XLColor headerColor)
    {
        for (int i = 0; i < properties.Length; i++)
        {
            var cell = worksheet.Cell(1, i + 1);
            cell.Value = FormatPropertyName(properties[i].Name);
            cell.Style.Fill.BackgroundColor = headerColor;
            cell.Style.Font.Bold = true;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        }
    }

    private static void AddDataRows<T>(IXLWorksheet worksheet, List<T> dataList, PropertyInfo[] properties)
    {
        for (int rowIndex = 0; rowIndex < dataList.Count; rowIndex++)
        {
            var item = dataList[rowIndex];
            if (item == null) continue;

            for (int colIndex = 0; colIndex < properties.Length; colIndex++)
            {
                var cell = worksheet.Cell(rowIndex + 2, colIndex + 1);
                var value = properties[colIndex].GetValue(item);
                
                SetCellValue(cell, value, properties[colIndex].PropertyType);
                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            }
        }
    }

    private static void SetCellValue(IXLCell cell, object? value, Type propertyType)
    {
        if (value == null)
        {
            cell.Value = string.Empty;
            return;
        }

        var underlyingType = Nullable.GetUnderlyingType(propertyType) ?? propertyType;

        if (underlyingType == typeof(decimal) || underlyingType == typeof(double) || underlyingType == typeof(float))
        {
            cell.Value = Convert.ToDouble(value);
            cell.Style.NumberFormat.Format = "#,##0.00";
        }
        else if (underlyingType == typeof(int) || underlyingType == typeof(long) || 
                 underlyingType == typeof(short) || underlyingType == typeof(byte))
        {
            cell.Value = Convert.ToDouble(value);
            cell.Style.NumberFormat.Format = "#,##0";
        }
        else if (underlyingType == typeof(DateTime))
        {
            cell.Value = (DateTime)value;
            cell.Style.DateFormat.Format = "yyyy-MM-dd HH:mm:ss";
        }
        else if (underlyingType == typeof(DateOnly))
        {
            cell.Value = ((DateOnly)value).ToDateTime(TimeOnly.MinValue);
            cell.Style.DateFormat.Format = "yyyy-MM-dd";
        }
        else if (underlyingType == typeof(bool))
        {
            cell.Value = (bool)value ? "Yes" : "No";
        }
        else
        {
            cell.Value = value.ToString();
        }
    }

    private static void AddSummationRow<T>(IXLWorksheet worksheet, List<T> dataList, PropertyInfo[] properties)
    {
        if (dataList.Count == 0) return;

        var summationRow = dataList.Count + 2;
        bool hasSummation = false;

        for (int colIndex = 0; colIndex < properties.Length; colIndex++)
        {
            var property = properties[colIndex];
            var underlyingType = Nullable.GetUnderlyingType(property.PropertyType) ?? property.PropertyType;

            // Check if this is a numeric type
            if (IsNumericType(underlyingType))
            {
                hasSummation = true;
                double sum = CalculateSum(dataList, property, underlyingType);

                var cell = worksheet.Cell(summationRow, colIndex + 1);
                cell.Value = sum;

                // Apply appropriate number format based on type
                if (IsFloatingPointType(underlyingType))
                {
                    cell.Style.NumberFormat.Format = "#,##0.00";
                }
                else
                {
                    cell.Style.NumberFormat.Format = "#,##0";
                }

                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = XLColor.LightGray;
                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            }
        }

        // Add "Total" label in the first column if there are summations
        if (hasSummation)
        {
            var firstCell = worksheet.Cell(summationRow, 1);
            if (string.IsNullOrEmpty(firstCell.GetString()) || !firstCell.Style.Font.Bold)
            {
                // Check if the first column is not a numeric column
                var firstProperty = properties[0];
                var firstUnderlyingType = Nullable.GetUnderlyingType(firstProperty.PropertyType) ?? firstProperty.PropertyType;

                if (!IsNumericType(firstUnderlyingType))
                {
                    firstCell.Value = "Total";
                    firstCell.Style.Font.Bold = true;
                    firstCell.Style.Fill.BackgroundColor = XLColor.LightGray;
                    firstCell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                }
            }
        }
    }

    private static bool IsNumericType(Type type)
    {
        return type == typeof(decimal) || type == typeof(double) || type == typeof(float) ||
               type == typeof(int) || type == typeof(long) || type == typeof(short) || type == typeof(byte);
    }

    private static bool IsFloatingPointType(Type type)
    {
        return type == typeof(decimal) || type == typeof(double) || type == typeof(float);
    }

    private static double CalculateSum<T>(List<T> dataList, PropertyInfo property, Type underlyingType)
    {
        if (underlyingType == typeof(decimal))
        {
            var sum = dataList
                .Select(item => item == null ? 0m : (decimal)(property.GetValue(item) ?? 0m))
                .Sum();
            return (double)sum.RefineValue();
        }
        else if (underlyingType == typeof(double))
        {
            var sum = dataList
                .Select(item => item == null ? 0.0 : (double)(property.GetValue(item) ?? 0.0))
                .Sum();
            return (double)((decimal)sum).RefineValue();
        }
        else if (underlyingType == typeof(float))
        {
            var sum = dataList
                .Select(item => item == null ? 0f : (float)(property.GetValue(item) ?? 0f))
                .Sum();
            return (double)((decimal)sum).RefineValue();
        }
        else if (underlyingType == typeof(int))
        {
            return dataList
                .Select(item => item == null ? 0 : (int)(property.GetValue(item) ?? 0))
                .Sum();
        }
        else if (underlyingType == typeof(long))
        {
            return dataList
                .Select(item => item == null ? 0L : (long)(property.GetValue(item) ?? 0L))
                .Sum();
        }
        else if (underlyingType == typeof(short))
        {
            return dataList
                .Select(item => item == null ? 0 : (int)(short)(property.GetValue(item) ?? (short)0))
                .Sum();
        }
        else if (underlyingType == typeof(byte))
        {
            return dataList
                .Select(item => item == null ? 0 : (int)(byte)(property.GetValue(item) ?? (byte)0))
                .Sum();
        }

        return 0;
    }

    private static string FormatPropertyName(string propertyName)
    {
        // Insert spaces before capital letters (for PascalCase properties)
        var formatted = System.Text.RegularExpressions.Regex.Replace(
            propertyName,
            "([a-z])([A-Z])",
            "$1 $2");

        return formatted;
    }
}
