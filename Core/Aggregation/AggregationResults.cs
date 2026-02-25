namespace ExcelGenerator.Core.Aggregation;

/// <summary>
/// Result of single-pass aggregation calculation
/// Contains all aggregation values computed in one iteration
/// </summary>
internal class AggregationResults
{
    public double Sum { get; set; }
    public double Average { get; set; }
    public double Min { get; set; }
    public double Max { get; set; }
    public int Count { get; set; }
}
