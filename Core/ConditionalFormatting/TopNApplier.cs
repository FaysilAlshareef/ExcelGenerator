using ClosedXML.Excel;

namespace ExcelGenerator.Core.ConditionalFormatting;

/// <summary>
/// Applies highlighting to top N values
/// </summary>
internal class TopNApplier : IFormattingRuleApplier
{
    public void Apply(IXLRange range, ConditionalFormattingRule rule)
    {
        var topNRule = range.AddConditionalFormat();
        topNRule.WhenIsTop(rule.TopN)
            .Fill.SetBackgroundColor(XLColor.LightGreen);
    }

    public ConditionalFormattingRuleType RuleType => ConditionalFormattingRuleType.HighlightTopN;
}
