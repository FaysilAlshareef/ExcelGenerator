using ClosedXML.Excel;
using ExcelGenerator.Core.PropertyReflection;
using ExcelGenerator.Core.Generators;
using ExcelGenerator.Core.ConditionalFormatting;
using System.Text;

namespace ExcelGenerator.Core;

/// <summary>
/// Main orchestrator for Excel generation using dependency injection
/// Coordinates all specialized components following Single Responsibility Principle
/// </summary>
internal class ExcelGeneratorEngine
{
    private readonly PropertyExtractor _propertyExtractor;
    private readonly HeaderGenerator _headerGenerator;
    private readonly DataRowGenerator _dataRowGenerator;
    private readonly AggregationRowGenerator _aggregationGenerator;
    private readonly FormattingRuleApplierFactory _formattingFactory;
    private readonly WorksheetLayoutManager _layoutManager;

    public ExcelGeneratorEngine(
        PropertyExtractor propertyExtractor,
        HeaderGenerator headerGenerator,
        DataRowGenerator dataRowGenerator,
        AggregationRowGenerator aggregationGenerator,
        FormattingRuleApplierFactory formattingFactory,
        WorksheetLayoutManager layoutManager)
    {
        _propertyExtractor = propertyExtractor;
        _headerGenerator = headerGenerator;
        _dataRowGenerator = dataRowGenerator;
        _aggregationGenerator = aggregationGenerator;
        _formattingFactory = formattingFactory;
        _layoutManager = layoutManager;
    }

    /// <summary>
    /// Generates Excel workbook with full configuration support
    /// </summary>
    public XLWorkbook Generate<T>(
        IEnumerable<T> data,
        string sheetName,
        ExcelConfiguration<T> configuration)
    {
        // Validate inputs
        ValidateInputs(data, sheetName, configuration);

        var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add(sheetName);

        var properties = _propertyExtractor.Extract<T>(configuration.ExcludeIds);

        if (properties.Length == 0)
        {
            throw new InvalidOperationException(
                $"Type '{typeof(T).Name}' has no readable properties. Cannot generate Excel sheet.");
        }

        var dataList = data.ToList();

        // Generate headers
        _headerGenerator.Generate(worksheet, properties, configuration.HeaderColor);

        // Generate data rows
        var rowCount = _dataRowGenerator.Generate(worksheet, dataList, properties);

        // Generate aggregation rows if configured
        if (configuration.Aggregations != AggregationType.None)
        {
            _aggregationGenerator.Generate(worksheet, dataList, properties, rowCount, configuration.Aggregations);
        }

        // Apply conditional formatting if configured
        if (configuration.ConditionalFormatting != null)
        {
            ApplyConditionalFormatting(worksheet, properties, rowCount, configuration.ConditionalFormatting);
        }

        // Apply layout settings
        _layoutManager.ApplyLayout(worksheet, configuration.FreezeRowCount, configuration.FreezeColumnCount);

        return workbook;
    }

    /// <summary>
    /// Simplified generation for basic scenarios (backward compatibility)
    /// </summary>
    public XLWorkbook Generate<T>(
        IEnumerable<T> data,
        string sheetName,
        bool excludeIds = false,
        XLColor? headerColor = null)
    {
        // Create configuration (validation happens in main Generate method)
        var config = new ExcelConfiguration<T>();
        if (excludeIds) config.WithExcludeIds();
        if (headerColor != null) config.WithHeaderColor(headerColor);

        return Generate(data, sheetName, config);
    }

    private void ApplyConditionalFormatting(IXLWorksheet worksheet, System.Reflection.PropertyInfo[] properties,
        int dataCount, ConditionalFormattingConfiguration config)
    {
        // Build a dictionary for O(1) property name lookups
        var propertyIndexMap = new Dictionary<string, int>(properties.Length);
        for (int i = 0; i < properties.Length; i++)
        {
            propertyIndexMap[properties[i].Name] = i;
        }

        foreach (var rule in config.Rules)
        {
            // Find the column index for this property using dictionary lookup
            if (!propertyIndexMap.TryGetValue(rule.ColumnName, out var colIndex))
                continue;

            var columnLetter = GetColumnLetter(colIndex + 1);
            var dataRange = worksheet.Range($"{columnLetter}2:{columnLetter}{dataCount + 1}");

            // Use Strategy pattern via FormattingRuleApplierFactory
            var applier = _formattingFactory.GetApplier(rule.Type);
            applier.Apply(dataRange, rule);
        }
    }

    private static string GetColumnLetter(int columnNumber)
    {
        var sb = new StringBuilder();
        while (columnNumber > 0)
        {
            int modulo = (columnNumber - 1) % 26;
            sb.Insert(0, (char)('A' + modulo));
            columnNumber = (columnNumber - modulo) / 26;
        }
        return sb.ToString();
    }

    /// <summary>
    /// Validates all input parameters for Excel generation
    /// </summary>
    private static void ValidateInputs<T>(IEnumerable<T> data, string sheetName, ExcelConfiguration<T> configuration)
    {
        // Validate data parameter
        if (data == null)
        {
            throw new ArgumentNullException(nameof(data),
                "Data collection cannot be null. Provide an empty collection if no data is available.");
        }

        // Validate sheet name
        ValidateSheetName(sheetName);

        // Validate configuration
        if (configuration == null)
        {
            throw new ArgumentNullException(nameof(configuration),
                "Configuration cannot be null. Use ExcelConfiguration<T> constructor to create a valid configuration.");
        }

        // Validate conditional formatting column names if configured
        if (configuration.ConditionalFormatting != null)
        {
            foreach (var rule in configuration.ConditionalFormatting.Rules)
            {
                if (string.IsNullOrWhiteSpace(rule.ColumnName))
                {
                    throw new ArgumentException(
                        "Conditional formatting rule has null or empty column name.",
                        nameof(configuration));
                }
            }
        }
    }

    /// <summary>
    /// Validates sheet name according to Excel requirements
    /// </summary>
    private static void ValidateSheetName(string sheetName)
    {
        if (string.IsNullOrWhiteSpace(sheetName))
        {
            throw new ArgumentException(
                "Sheet name cannot be null or empty.",
                nameof(sheetName));
        }

        if (sheetName.Length > 31)
        {
            throw new ArgumentException(
                $"Sheet name '{sheetName}' exceeds maximum length of 31 characters. Current length: {sheetName.Length}.",
                nameof(sheetName));
        }

        // Excel sheet name invalid characters: : \ / ? * [ ]
        char[] invalidChars = { ':', '\\', '/', '?', '*', '[', ']' };
        foreach (var invalidChar in invalidChars)
        {
            if (sheetName.Contains(invalidChar))
            {
                throw new ArgumentException(
                    $"Sheet name '{sheetName}' contains invalid character '{invalidChar}'. " +
                    $"Excel sheet names cannot contain: : \\ / ? * [ ]",
                    nameof(sheetName));
            }
        }
    }
}
