using System.Reflection;

namespace ExcelGenerator.Core.Aggregation;

/// <summary>
/// Defines a strategy for calculating aggregations on numeric properties
/// </summary>
internal interface IAggregationStrategy
{
    /// <summary>
    /// Calculates the aggregation for the specified property across all items in the dataset
    /// </summary>
    /// <typeparam name="T">The type of items in the dataset</typeparam>
    /// <param name="dataList">The list of items to aggregate</param>
    /// <param name="property">The property to aggregate</param>
    /// <param name="underlyingType">The underlying type of the property (unwrapped from Nullable if applicable)</param>
    /// <returns>The calculated aggregation value</returns>
    double Calculate<T>(List<T> dataList, PropertyInfo property, Type underlyingType);

    /// <summary>
    /// Gets the name of this aggregation strategy (e.g., "Sum", "Average")
    /// </summary>
    string Name { get; }
}
