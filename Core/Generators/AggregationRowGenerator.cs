using ClosedXML.Excel;
using System.Reflection;
using ExcelGenerator.Core.Aggregation;

namespace ExcelGenerator.Core.Generators;

/// <summary>
/// Generates aggregation rows (Sum, Average, Min, Max, Count) in Excel worksheets
/// Single responsibility: Aggregation row creation
/// </summary>
internal class AggregationRowGenerator
{
    private readonly AggregationStrategyFactory _aggregationFactory;

    public AggregationRowGenerator(AggregationStrategyFactory aggregationFactory)
    {
        _aggregationFactory = aggregationFactory;
    }

    /// <summary>
    /// Generates aggregation rows based on the specified aggregation types
    /// </summary>
    public void Generate<T>(IXLWorksheet worksheet, List<T> dataList, PropertyInfo[] properties,
        int dataRowCount, AggregationType aggregations)
    {
        // Validate inputs
        if (worksheet == null)
            throw new ArgumentNullException(nameof(worksheet), "Worksheet cannot be null.");
        if (dataList == null)
            throw new ArgumentNullException(nameof(dataList), "Data list cannot be null.");
        if (properties == null)
            throw new ArgumentNullException(nameof(properties), "Properties array cannot be null.");
        if (dataRowCount < 0)
            throw new ArgumentOutOfRangeException(nameof(dataRowCount), "Data row count cannot be negative.");

        if (dataList.Count == 0 || aggregations == AggregationType.None) return;

        var startRow = dataRowCount + 2;
        var currentRow = startRow;

        // Add Sum aggregation
        if (aggregations.HasFlag(AggregationType.Sum))
        {
            AddAggregationRow(worksheet, dataList, properties, currentRow, "Sum",
                AggregationType.Sum, XLColor.LightGray);
            currentRow++;
        }

        // Add Average aggregation
        if (aggregations.HasFlag(AggregationType.Average))
        {
            AddAggregationRow(worksheet, dataList, properties, currentRow, "Average",
                AggregationType.Average, XLColor.AliceBlue);
            currentRow++;
        }

        // Add Min aggregation
        if (aggregations.HasFlag(AggregationType.Min))
        {
            AddAggregationRow(worksheet, dataList, properties, currentRow, "Min",
                AggregationType.Min, XLColor.LightYellow);
            currentRow++;
        }

        // Add Max aggregation
        if (aggregations.HasFlag(AggregationType.Max))
        {
            AddAggregationRow(worksheet, dataList, properties, currentRow, "Max",
                AggregationType.Max, XLColor.LightGreen);
            currentRow++;
        }

        // Add Count aggregation
        if (aggregations.HasFlag(AggregationType.Count))
        {
            AddAggregationRow(worksheet, dataList, properties, currentRow, "Count",
                AggregationType.Count, XLColor.Lavender);
        }
    }

    private void AddAggregationRow<T>(IXLWorksheet worksheet, List<T> dataList, PropertyInfo[] properties,
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

                var strategy = _aggregationFactory.GetStrategy(aggregationType);
                double value = strategy.Calculate(dataList, property, underlyingType);

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

    private static bool IsNumericType(Type type)
    {
        return type == typeof(decimal) || type == typeof(double) || type == typeof(float) ||
               type == typeof(int) || type == typeof(long) || type == typeof(short) || type == typeof(byte);
    }

    private static bool IsFloatingPointType(Type type)
    {
        return type == typeof(decimal) || type == typeof(double) || type == typeof(float);
    }
}
