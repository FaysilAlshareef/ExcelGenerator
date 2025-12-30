using ClosedXML.Excel;

namespace ExcelGenerator.Core.ConditionalFormatting;

/// <summary>
/// Applies data bars formatting
/// </summary>
internal class DataBarsApplier : IFormattingRuleApplier
{
    public void Apply(IXLRange range, ConditionalFormattingRule rule)
    {
        var dataBarRule = range.AddConditionalFormat();
        dataBarRule.DataBar(rule.BarColor ?? XLColor.Blue);
    }

    public ConditionalFormattingRuleType RuleType => ConditionalFormattingRuleType.DataBars;
}
