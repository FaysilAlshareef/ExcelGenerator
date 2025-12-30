using ClosedXML.Excel;

namespace ExcelGenerator.Core.CellFormatters;

/// <summary>
/// Defines a strategy for formatting cell values of specific types
/// </summary>
internal interface ICellValueFormatter
{
    /// <summary>
    /// Determines if this formatter can handle the specified type
    /// </summary>
    /// <param name="type">The type to check</param>
    /// <returns>True if this formatter can handle the type, false otherwise</returns>
    bool CanFormat(Type type);

    /// <summary>
    /// Formats the cell value and applies appropriate styling
    /// </summary>
    /// <param name="cell">The cell to format</param>
    /// <param name="value">The value to set in the cell</param>
    /// <param name="type">The type of the value</param>
    void Format(IXLCell cell, object? value, Type type);

    /// <summary>
    /// Gets the priority of this formatter. Higher priority formatters are checked first.
    /// </summary>
    int Priority { get; }
}
