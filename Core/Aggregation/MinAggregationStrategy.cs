using System.Reflection;

namespace ExcelGenerator.Core.Aggregation;

/// <summary>
/// Strategy for calculating minimum aggregations
/// </summary>
internal class MinAggregationStrategy : IAggregationStrategy
{
    public double Calculate<T>(List<T> dataList, PropertyInfo property, Type underlyingType)
    {
        return NumericAggregator.CalculateMin(dataList, property, underlyingType);
    }

    public string Name => "Min";
}
