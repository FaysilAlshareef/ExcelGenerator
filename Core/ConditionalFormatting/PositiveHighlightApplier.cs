using ClosedXML.Excel;

namespace ExcelGenerator.Core.ConditionalFormatting;

/// <summary>
/// Applies highlighting to positive values
/// </summary>
internal class PositiveHighlightApplier : IFormattingRuleApplier
{
    public void Apply(IXLRange range, ConditionalFormattingRule rule)
    {
        var positiveRule = range.AddConditionalFormat();
        positiveRule.WhenGreaterThan(0)
            .Fill.SetBackgroundColor(XLColor.LightGreen);
    }

    public ConditionalFormattingRuleType RuleType => ConditionalFormattingRuleType.HighlightPositives;
}
