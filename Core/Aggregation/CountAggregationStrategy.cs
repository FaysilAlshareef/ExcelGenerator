using System.Reflection;

namespace ExcelGenerator.Core.Aggregation;

/// <summary>
/// Strategy for counting records
/// </summary>
internal class CountAggregationStrategy : IAggregationStrategy
{
    public double Calculate<T>(List<T> dataList, PropertyInfo property, Type underlyingType)
    {
        return dataList.Count;
    }

    public string Name => "Count";
}
