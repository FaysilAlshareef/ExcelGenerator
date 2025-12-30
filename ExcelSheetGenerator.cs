using ClosedXML.Excel;
using System.Reflection;

namespace ExcelGenerator;

/// <summary>
/// Generates Excel sheets from IEnumerable collections
/// </summary>
public static class ExcelSheetGenerator
{
    /// <summary>
    /// Creates a new Excel configuration for advanced features
    /// </summary>
    /// <typeparam name="T">The type of objects in the collection</typeparam>
    /// <returns>A new ExcelConfiguration instance for fluent configuration</returns>
    public static ExcelConfiguration<T> Configure<T>()
    {
        return new ExcelConfiguration<T>();
    }
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

    /// <summary>
    /// Generates an Excel workbook from a collection using advanced configuration
    /// </summary>
    /// <typeparam name="T">The type of objects in the collection</typeparam>
    /// <param name="data">The collection of data to export</param>
    /// <param name="sheetName">The name of the worksheet</param>
    /// <param name="configuration">The configuration for Excel generation</param>
    /// <returns>An XLWorkbook containing the generated Excel sheet</returns>
    internal static XLWorkbook GenerateExcel<T>(
        IEnumerable<T> data,
        string sheetName,
        ExcelConfiguration<T> configuration)
    {
        var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add(sheetName);

        var properties = GetProperties<T>(configuration.ExcludeIds);

        if (properties.Length == 0)
            return workbook;

        // Add headers
        AddHeaders(worksheet, properties, configuration.HeaderColor);

        // Add data rows
        var dataList = data.ToList();
        AddDataRows(worksheet, dataList, properties);

        // Add aggregation rows based on configuration
        AddAggregationRows(worksheet, dataList, properties, configuration.Aggregations);

        // Apply conditional formatting
        if (configuration.ConditionalFormatting != null)
        {
            ApplyConditionalFormatting(worksheet, properties, dataList.Count, configuration.ConditionalFormatting);
        }

        // Apply freeze panes
        if (configuration.FreezeRowCount > 0 || configuration.FreezeColumnCount > 0)
        {
            worksheet.SheetView.FreezeRows(configuration.FreezeRowCount);
            worksheet.SheetView.FreezeColumns(configuration.FreezeColumnCount);
        }

        // Auto-fit columns
        worksheet.Columns().AdjustToContents();

        return workbook;
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

    private static void AddAggregationRows<T>(IXLWorksheet worksheet, List<T> dataList, PropertyInfo[] properties, AggregationType aggregations)
    {
        if (dataList.Count == 0 || aggregations == AggregationType.None) return;

        var startRow = dataList.Count + 2;
        var currentRow = startRow;

        // Add Sum aggregation
        if (aggregations.HasFlag(AggregationType.Sum))
        {
            AddAggregationRow(worksheet, dataList, properties, currentRow, "Sum", AggregationType.Sum, XLColor.LightGray);
            currentRow++;
        }

        // Add Average aggregation
        if (aggregations.HasFlag(AggregationType.Average))
        {
            AddAggregationRow(worksheet, dataList, properties, currentRow, "Average", AggregationType.Average, XLColor.AliceBlue);
            currentRow++;
        }

        // Add Min aggregation
        if (aggregations.HasFlag(AggregationType.Min))
        {
            AddAggregationRow(worksheet, dataList, properties, currentRow, "Min", AggregationType.Min, XLColor.LightYellow);
            currentRow++;
        }

        // Add Max aggregation
        if (aggregations.HasFlag(AggregationType.Max))
        {
            AddAggregationRow(worksheet, dataList, properties, currentRow, "Max", AggregationType.Max, XLColor.LightGreen);
            currentRow++;
        }

        // Add Count aggregation
        if (aggregations.HasFlag(AggregationType.Count))
        {
            AddAggregationRow(worksheet, dataList, properties, currentRow, "Count", AggregationType.Count, XLColor.Lavender);
        }
    }

    private static void AddAggregationRow<T>(IXLWorksheet worksheet, List<T> dataList, PropertyInfo[] properties,
        int row, string label, AggregationType aggregationType, XLColor backgroundColor)
    {
        bool hasAggregation = false;

        for (int colIndex = 0; colIndex < properties.Length; colIndex++)
        {
            var property = properties[colIndex];
            var underlyingType = Nullable.GetUnderlyingType(property.PropertyType) ?? property.PropertyType;

            if (IsNumericType(underlyingType))
            {
                hasAggregation = true;
                double value = CalculateAggregation(dataList, property, underlyingType, aggregationType);

                var cell = worksheet.Cell(row, colIndex + 1);
                cell.Value = value;

                // Apply appropriate number format based on type and aggregation
                if (aggregationType == AggregationType.Count)
                {
                    cell.Style.NumberFormat.Format = "#,##0";
                }
                else if (IsFloatingPointType(underlyingType))
                {
                    cell.Style.NumberFormat.Format = "#,##0.00";
                }
                else
                {
                    cell.Style.NumberFormat.Format = "#,##0";
                }

                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = backgroundColor;
                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            }
        }

        // Add label in the first column if there are aggregations
        if (hasAggregation)
        {
            var firstCell = worksheet.Cell(row, 1);
            if (string.IsNullOrEmpty(firstCell.GetString()) || !firstCell.Style.Font.Bold)
            {
                var firstProperty = properties[0];
                var firstUnderlyingType = Nullable.GetUnderlyingType(firstProperty.PropertyType) ?? firstProperty.PropertyType;

                if (!IsNumericType(firstUnderlyingType))
                {
                    firstCell.Value = label;
                    firstCell.Style.Font.Bold = true;
                    firstCell.Style.Fill.BackgroundColor = backgroundColor;
                    firstCell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                }
            }
        }
    }

    private static double CalculateAggregation<T>(List<T> dataList, PropertyInfo property, Type underlyingType, AggregationType aggregationType)
    {
        return aggregationType switch
        {
            AggregationType.Sum => CalculateSum(dataList, property, underlyingType),
            AggregationType.Average => CalculateAverage(dataList, property, underlyingType),
            AggregationType.Min => CalculateMin(dataList, property, underlyingType),
            AggregationType.Max => CalculateMax(dataList, property, underlyingType),
            AggregationType.Count => dataList.Count,
            _ => 0
        };
    }

    private static double CalculateAverage<T>(List<T> dataList, PropertyInfo property, Type underlyingType)
    {
        if (dataList.Count == 0) return 0;

        var sum = CalculateSum(dataList, property, underlyingType);
        var average = sum / dataList.Count;

        // Apply refinement for floating-point types
        if (underlyingType == typeof(decimal) || underlyingType == typeof(double) || underlyingType == typeof(float))
        {
            return (double)((decimal)average).RefineValue();
        }

        return average;
    }

    private static double CalculateMin<T>(List<T> dataList, PropertyInfo property, Type underlyingType)
    {
        if (underlyingType == typeof(decimal))
        {
            var min = dataList
                .Select(item => item == null ? decimal.MaxValue : (decimal)(property.GetValue(item) ?? decimal.MaxValue))
                .Min();
            return (double)min.RefineValue();
        }
        else if (underlyingType == typeof(double))
        {
            var min = dataList
                .Select(item => item == null ? double.MaxValue : (double)(property.GetValue(item) ?? double.MaxValue))
                .Min();
            return (double)((decimal)min).RefineValue();
        }
        else if (underlyingType == typeof(float))
        {
            var min = dataList
                .Select(item => item == null ? float.MaxValue : (float)(property.GetValue(item) ?? float.MaxValue))
                .Min();
            return (double)((decimal)min).RefineValue();
        }
        else if (underlyingType == typeof(int))
        {
            return dataList
                .Select(item => item == null ? int.MaxValue : (int)(property.GetValue(item) ?? int.MaxValue))
                .Min();
        }
        else if (underlyingType == typeof(long))
        {
            return dataList
                .Select(item => item == null ? long.MaxValue : (long)(property.GetValue(item) ?? long.MaxValue))
                .Min();
        }
        else if (underlyingType == typeof(short))
        {
            return dataList
                .Select(item => item == null ? short.MaxValue : (int)(short)(property.GetValue(item) ?? short.MaxValue))
                .Min();
        }
        else if (underlyingType == typeof(byte))
        {
            return dataList
                .Select(item => item == null ? byte.MaxValue : (int)(byte)(property.GetValue(item) ?? byte.MaxValue))
                .Min();
        }

        return 0;
    }

    private static double CalculateMax<T>(List<T> dataList, PropertyInfo property, Type underlyingType)
    {
        if (underlyingType == typeof(decimal))
        {
            var max = dataList
                .Select(item => item == null ? decimal.MinValue : (decimal)(property.GetValue(item) ?? decimal.MinValue))
                .Max();
            return (double)max.RefineValue();
        }
        else if (underlyingType == typeof(double))
        {
            var max = dataList
                .Select(item => item == null ? double.MinValue : (double)(property.GetValue(item) ?? double.MinValue))
                .Max();
            return (double)((decimal)max).RefineValue();
        }
        else if (underlyingType == typeof(float))
        {
            var max = dataList
                .Select(item => item == null ? float.MinValue : (float)(property.GetValue(item) ?? float.MinValue))
                .Max();
            return (double)((decimal)max).RefineValue();
        }
        else if (underlyingType == typeof(int))
        {
            return dataList
                .Select(item => item == null ? int.MinValue : (int)(property.GetValue(item) ?? int.MinValue))
                .Max();
        }
        else if (underlyingType == typeof(long))
        {
            return dataList
                .Select(item => item == null ? long.MinValue : (long)(property.GetValue(item) ?? long.MinValue))
                .Max();
        }
        else if (underlyingType == typeof(short))
        {
            return dataList
                .Select(item => item == null ? short.MinValue : (int)(short)(property.GetValue(item) ?? short.MinValue))
                .Max();
        }
        else if (underlyingType == typeof(byte))
        {
            return dataList
                .Select(item => item == null ? byte.MinValue : (int)(byte)(property.GetValue(item) ?? byte.MinValue))
                .Max();
        }

        return 0;
    }

    private static void ApplyConditionalFormatting(IXLWorksheet worksheet, PropertyInfo[] properties,
        int dataCount, ConditionalFormattingConfiguration config)
    {
        foreach (var rule in config.Rules)
        {
            // Find the column index for this property
            var colIndex = Array.FindIndex(properties, p => p.Name == rule.ColumnName);
            if (colIndex < 0) continue;

            var columnLetter = GetColumnLetter(colIndex + 1);
            var dataRange = worksheet.Range($"{columnLetter}2:{columnLetter}{dataCount + 1}");

            switch (rule.Type)
            {
                case ConditionalFormattingRuleType.HighlightNegatives:
                    var negativeRule = dataRange.AddConditionalFormat();
                    negativeRule.WhenLessThan(0)
                        .Fill.SetBackgroundColor(XLColor.LightPink);
                    break;

                case ConditionalFormattingRuleType.HighlightPositives:
                    var positiveRule = dataRange.AddConditionalFormat();
                    positiveRule.WhenGreaterThan(0)
                        .Fill.SetBackgroundColor(XLColor.LightGreen);
                    break;

                case ConditionalFormattingRuleType.ColorScale:
                    var colorScaleRule = dataRange.AddConditionalFormat();
                    colorScaleRule.ColorScale()
                        .LowestValue(rule.MinColor ?? XLColor.Red)
                        .HighestValue(rule.MaxColor ?? XLColor.Green);
                    break;

                case ConditionalFormattingRuleType.DataBars:
                    var dataBarRule = dataRange.AddConditionalFormat();
                    dataBarRule.DataBar(rule.BarColor ?? XLColor.Blue);
                    break;

                case ConditionalFormattingRuleType.HighlightDuplicates:
                    var duplicateRule = dataRange.AddConditionalFormat();
                    duplicateRule.WhenIsDuplicate()
                        .Fill.SetBackgroundColor(XLColor.Yellow);
                    break;

                case ConditionalFormattingRuleType.HighlightTopN:
                    var topNRule = dataRange.AddConditionalFormat();
                    topNRule.WhenIsTop(rule.TopN)
                        .Fill.SetBackgroundColor(XLColor.LightGreen);
                    break;
            }
        }
    }

    private static string GetColumnLetter(int columnNumber)
    {
        string columnName = "";
        while (columnNumber > 0)
        {
            int modulo = (columnNumber - 1) % 26;
            columnName = Convert.ToChar('A' + modulo) + columnName;
            columnNumber = (columnNumber - modulo) / 26;
        }
        return columnName;
    }
}
