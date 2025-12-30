using ClosedXML.Excel;

namespace ExcelGenerator.Core.ConditionalFormatting;

/// <summary>
/// Applies highlighting to duplicate values
/// </summary>
internal class DuplicatesApplier : IFormattingRuleApplier
{
    public void Apply(IXLRange range, ConditionalFormattingRule rule)
    {
        var duplicateRule = range.AddConditionalFormat();
        duplicateRule.WhenIsDuplicate()
            .Fill.SetBackgroundColor(XLColor.Yellow);
    }

    public ConditionalFormattingRuleType RuleType => ConditionalFormattingRuleType.HighlightDuplicates;
}
