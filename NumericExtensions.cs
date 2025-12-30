namespace ExcelGenerator;

/// <summary>
/// Extension methods for numeric types
/// </summary>
public static class NumericExtensions
{
    /// <summary>
    /// Refines a decimal value by truncating to 3 decimal places
    /// </summary>
    /// <param name="value">The decimal value to refine</param>
    /// <returns>The refined decimal value truncated to 3 decimal places</returns>
    public static decimal RefineValue(this decimal value) => Math.Truncate(1000 * value) / 1000;
}
