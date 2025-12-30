using ClosedXML.Excel;

namespace ExcelGenerator.Core.ConditionalFormatting;

/// <summary>
/// Applies highlighting to negative values
/// </summary>
internal class NegativeHighlightApplier : IFormattingRuleApplier
{
    public void Apply(IXLRange range, ConditionalFormattingRule rule)
    {
        var negativeRule = range.AddConditionalFormat();
        negativeRule.WhenLessThan(0)
            .Fill.SetBackgroundColor(XLColor.LightPink);
    }

    public ConditionalFormattingRuleType RuleType => ConditionalFormattingRuleType.HighlightNegatives;
}
