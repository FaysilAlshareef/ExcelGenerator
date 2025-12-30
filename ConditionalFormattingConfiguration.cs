using ClosedXML.Excel;

namespace ExcelGenerator;

/// <summary>
/// Configuration for conditional formatting rules
/// </summary>
public class ConditionalFormattingConfiguration
{
    internal List<ConditionalFormattingRule> Rules { get; } = new();

    /// <summary>
    /// Highlights cells with negative values in red
    /// </summary>
    /// <param name="columnName">The property name of the column to apply formatting to</param>
    /// <returns>The configuration for chaining</returns>
    public ConditionalFormattingConfiguration HighlightNegatives(string columnName)
    {
        Rules.Add(new ConditionalFormattingRule
        {
            ColumnName = columnName,
            Type = ConditionalFormattingRuleType.HighlightNegatives
        });
        return this;
    }

    /// <summary>
    /// Highlights cells with positive values in green
    /// </summary>
    /// <param name="columnName">The property name of the column to apply formatting to</param>
    /// <returns>The configuration for chaining</returns>
    public ConditionalFormattingConfiguration HighlightPositives(string columnName)
    {
        Rules.Add(new ConditionalFormattingRule
        {
            ColumnName = columnName,
            Type = ConditionalFormattingRuleType.HighlightPositives
        });
        return this;
    }

    /// <summary>
    /// Applies a color scale from minimum to maximum value
    /// </summary>
    /// <param name="columnName">The property name of the column to apply formatting to</param>
    /// <param name="minColor">Color for minimum values (default: Red)</param>
    /// <param name="maxColor">Color for maximum values (default: Green)</param>
    /// <returns>The configuration for chaining</returns>
    public ConditionalFormattingConfiguration ColorScale(string columnName, XLColor? minColor = null, XLColor? maxColor = null)
    {
        Rules.Add(new ConditionalFormattingRule
        {
            ColumnName = columnName,
            Type = ConditionalFormattingRuleType.ColorScale,
            MinColor = minColor ?? XLColor.Red,
            MaxColor = maxColor ?? XLColor.Green
        });
        return this;
    }

    /// <summary>
    /// Adds data bars to show value magnitude
    /// </summary>
    /// <param name="columnName">The property name of the column to apply formatting to</param>
    /// <param name="barColor">Color of the data bars (default: Blue)</param>
    /// <returns>The configuration for chaining</returns>
    public ConditionalFormattingConfiguration DataBars(string columnName, XLColor? barColor = null)
    {
        Rules.Add(new ConditionalFormattingRule
        {
            ColumnName = columnName,
            Type = ConditionalFormattingRuleType.DataBars,
            BarColor = barColor ?? XLColor.Blue
        });
        return this;
    }

    /// <summary>
    /// Highlights duplicate values in yellow
    /// </summary>
    /// <param name="columnName">The property name of the column to apply formatting to</param>
    /// <returns>The configuration for chaining</returns>
    public ConditionalFormattingConfiguration HighlightDuplicates(string columnName)
    {
        Rules.Add(new ConditionalFormattingRule
        {
            ColumnName = columnName,
            Type = ConditionalFormattingRuleType.HighlightDuplicates
        });
        return this;
    }

    /// <summary>
    /// Highlights the top N values in green
    /// </summary>
    /// <param name="columnName">The property name of the column to apply formatting to</param>
    /// <param name="topN">Number of top values to highlight</param>
    /// <returns>The configuration for chaining</returns>
    public ConditionalFormattingConfiguration HighlightTopN(string columnName, int topN = 10)
    {
        Rules.Add(new ConditionalFormattingRule
        {
            ColumnName = columnName,
            Type = ConditionalFormattingRuleType.HighlightTopN,
            TopN = topN
        });
        return this;
    }
}

internal class ConditionalFormattingRule
{
    public required string ColumnName { get; set; }
    public required ConditionalFormattingRuleType Type { get; set; }
    public XLColor? MinColor { get; set; }
    public XLColor? MaxColor { get; set; }
    public XLColor? BarColor { get; set; }
    public int TopN { get; set; }
}

internal enum ConditionalFormattingRuleType
{
    HighlightNegatives,
    HighlightPositives,
    ColorScale,
    DataBars,
    HighlightDuplicates,
    HighlightTopN
}
