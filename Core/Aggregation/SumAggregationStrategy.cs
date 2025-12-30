using System.Reflection;

namespace ExcelGenerator.Core.Aggregation;

/// <summary>
/// Strategy for calculating sum aggregations
/// </summary>
internal class SumAggregationStrategy : IAggregationStrategy
{
    public double Calculate<T>(List<T> dataList, PropertyInfo property, Type underlyingType)
    {
        return NumericAggregator.CalculateSum(dataList, property, underlyingType);
    }

    public string Name => "Sum";
}
