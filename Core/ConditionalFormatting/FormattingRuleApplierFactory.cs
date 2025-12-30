namespace ExcelGenerator.Core.ConditionalFormatting;

/// <summary>
/// Factory for creating formatting rule appliers based on rule type
/// </summary>
internal class FormattingRuleApplierFactory
{
    private readonly Dictionary<ConditionalFormattingRuleType, IFormattingRuleApplier> _appliers;

    public FormattingRuleApplierFactory()
    {
        _appliers = new Dictionary<ConditionalFormattingRuleType, IFormattingRuleApplier>
        {
            { ConditionalFormattingRuleType.HighlightNegatives, new NegativeHighlightApplier() },
            { ConditionalFormattingRuleType.HighlightPositives, new PositiveHighlightApplier() },
            { ConditionalFormattingRuleType.ColorScale, new ColorScaleApplier() },
            { ConditionalFormattingRuleType.DataBars, new DataBarsApplier() },
            { ConditionalFormattingRuleType.HighlightDuplicates, new DuplicatesApplier() },
            { ConditionalFormattingRuleType.HighlightTopN, new TopNApplier() }
        };
    }

    /// <summary>
    /// Gets the appropriate applier for the specified rule type
    /// </summary>
    public IFormattingRuleApplier GetApplier(ConditionalFormattingRuleType ruleType)
    {
        if (_appliers.TryGetValue(ruleType, out var applier))
        {
            return applier;
        }

        throw new ArgumentException($"Unknown formatting rule type: {ruleType}", nameof(ruleType));
    }
}
