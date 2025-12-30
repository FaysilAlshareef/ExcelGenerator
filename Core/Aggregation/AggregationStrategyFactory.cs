namespace ExcelGenerator.Core.Aggregation;

/// <summary>
/// Factory for creating aggregation strategies based on AggregationType
/// </summary>
internal class AggregationStrategyFactory
{
    private readonly Dictionary<AggregationType, IAggregationStrategy> _strategies;

    public AggregationStrategyFactory()
    {
        _strategies = new Dictionary<AggregationType, IAggregationStrategy>
        {
            { AggregationType.Sum, new SumAggregationStrategy() },
            { AggregationType.Average, new AverageAggregationStrategy() },
            { AggregationType.Min, new MinAggregationStrategy() },
            { AggregationType.Max, new MaxAggregationStrategy() },
            { AggregationType.Count, new CountAggregationStrategy() }
        };
    }

    /// <summary>
    /// Gets the appropriate aggregation strategy for the specified type
    /// </summary>
    public IAggregationStrategy GetStrategy(AggregationType aggregationType)
    {
        if (_strategies.TryGetValue(aggregationType, out var strategy))
        {
            return strategy;
        }

        throw new ArgumentException($"Unknown aggregation type: {aggregationType}", nameof(aggregationType));
    }
}
