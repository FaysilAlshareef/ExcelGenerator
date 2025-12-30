using ClosedXML.Excel;

namespace ExcelGenerator.Core.ConditionalFormatting;

/// <summary>
/// Applies color scale formatting
/// </summary>
internal class ColorScaleApplier : IFormattingRuleApplier
{
    public void Apply(IXLRange range, ConditionalFormattingRule rule)
    {
        var colorScaleRule = range.AddConditionalFormat();
        colorScaleRule.ColorScale()
            .LowestValue(rule.MinColor ?? XLColor.Red)
            .HighestValue(rule.MaxColor ?? XLColor.Green);
    }

    public ConditionalFormattingRuleType RuleType => ConditionalFormattingRuleType.ColorScale;
}
