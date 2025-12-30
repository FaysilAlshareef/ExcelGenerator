using ClosedXML.Excel;
using System.Reflection;

namespace ExcelGenerator.Core.ConditionalFormatting;

/// <summary>
/// Defines a strategy for applying conditional formatting rules to worksheet ranges
/// </summary>
internal interface IFormattingRuleApplier
{
    /// <summary>
    /// Applies the formatting rule to the specified range
    /// </summary>
    /// <param name="range">The range to apply formatting to</param>
    /// <param name="rule">The conditional formatting rule configuration</param>
    void Apply(IXLRange range, ConditionalFormattingRule rule);

    /// <summary>
    /// Gets the rule type that this applier handles
    /// </summary>
    ConditionalFormattingRuleType RuleType { get; }
}
