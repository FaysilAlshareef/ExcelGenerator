using ClosedXML.Excel;
using System.Reflection;
using ExcelGenerator.Core.PropertyReflection;

namespace ExcelGenerator.Core.Generators;

/// <summary>
/// Generates and formats header rows in Excel worksheets
/// Single responsibility: Header creation
/// </summary>
internal class HeaderGenerator
{
    private readonly PropertyExtractor _propertyExtractor;

    public HeaderGenerator(PropertyExtractor propertyExtractor)
    {
        _propertyExtractor = propertyExtractor;
    }

    /// <summary>
    /// Generates header row with formatting
    /// </summary>
    public void Generate(IXLWorksheet worksheet, PropertyInfo[] properties, XLColor headerColor)
    {
        // Validate inputs
        if (worksheet == null)
            throw new ArgumentNullException(nameof(worksheet), "Worksheet cannot be null.");
        if (properties == null)
            throw new ArgumentNullException(nameof(properties), "Properties array cannot be null.");

        for (int i = 0; i < properties.Length; i++)
        {
            var cell = worksheet.Cell(1, i + 1);
            cell.Value = _propertyExtractor.FormatPropertyName(properties[i].Name);
            cell.Style.Fill.BackgroundColor = headerColor;
            cell.Style.Font.Bold = true;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        }
    }
}
