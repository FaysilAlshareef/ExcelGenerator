using System.Reflection;

namespace ExcelGenerator.Core.Aggregation;

/// <summary>
/// Strategy for calculating maximum aggregations
/// </summary>
internal class MaxAggregationStrategy : IAggregationStrategy
{
    public double Calculate<T>(List<T> dataList, PropertyInfo property, Type underlyingType)
    {
        return NumericAggregator.CalculateMax(dataList, property, underlyingType);
    }

    public string Name => "Max";
}
