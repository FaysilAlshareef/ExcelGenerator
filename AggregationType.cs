namespace ExcelGenerator;

/// <summary>
/// Defines the types of aggregations that can be applied to numeric columns
/// </summary>
[Flags]
public enum AggregationType
{
    /// <summary>
    /// No aggregation
    /// </summary>
    None = 0,

    /// <summary>
    /// Sum of all values
    /// </summary>
    Sum = 1,

    /// <summary>
    /// Average of all values
    /// </summary>
    Average = 2,

    /// <summary>
    /// Minimum value
    /// </summary>
    Min = 4,

    /// <summary>
    /// Maximum value
    /// </summary>
    Max = 8,

    /// <summary>
    /// Count of all values
    /// </summary>
    Count = 16
}
