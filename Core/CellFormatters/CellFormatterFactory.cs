using ClosedXML.Excel;

namespace ExcelGenerator.Core.CellFormatters;

/// <summary>
/// Factory for creating and selecting appropriate cell value formatters
/// </summary>
internal class CellFormatterFactory
{
    private readonly List<ICellValueFormatter> _formatters;
    private readonly NullValueFormatter _nullFormatter;
    private readonly StringFormatter _fallbackFormatter;

    public CellFormatterFactory()
    {
        _nullFormatter = new NullValueFormatter();
        _fallbackFormatter = new StringFormatter();

        // Register all formatters ordered by priority (high to low)
        _formatters = new List<ICellValueFormatter>
        {
            new DecimalFormatter(),
            new IntegerFormatter(),
            new DateTimeFormatter(),
            new DateOnlyFormatter(),
            new BooleanFormatter(),
            _fallbackFormatter
        };
    }

    /// <summary>
    /// Formats a cell value using the appropriate formatter
    /// </summary>
    /// <param name="cell">The cell to format</param>
    /// <param name="value">The value to set</param>
    /// <param name="type">The type of the property</param>
    public void FormatCell(IXLCell cell, object? value, Type type)
    {
        // Handle null values first
        if (value == null)
        {
            _nullFormatter.Format(cell, value, type);
            return;
        }

        // Get underlying type if nullable
        var underlyingType = Nullable.GetUnderlyingType(type) ?? type;

        // Find the first formatter that can handle this type
        var formatter = GetFormatter(underlyingType);
        formatter.Format(cell, value, underlyingType);
    }

    /// <summary>
    /// Gets the appropriate formatter for the specified type
    /// </summary>
    private ICellValueFormatter GetFormatter(Type type)
    {
        return _formatters
            .Where(f => f.CanFormat(type))
            .OrderByDescending(f => f.Priority)
            .FirstOrDefault() ?? _fallbackFormatter;
    }
}
