using System.Reflection;

namespace ExcelGenerator.Core.Aggregation;

/// <summary>
/// Strategy for calculating average aggregations
/// </summary>
internal class AverageAggregationStrategy : IAggregationStrategy
{
    public double Calculate<T>(List<T> dataList, PropertyInfo property, Type underlyingType)
    {
        return NumericAggregator.CalculateAverage(dataList, property, underlyingType);
    }

    public string Name => "Average";
}
