# ExcelGenerator

A lightweight .NET library to generate Excel files from `IEnumerable<T>` collections using ClosedXML.

## Version History

### V3 (Current)
- **Framework**: .NET 10.0
- **C# Version**: 14
- **Package Version**: 3.0.0 (latest)
- **Features**: Advanced fluent API with aggregations, conditional formatting, multi-sheet support, and freeze panes
- **Breaking Changes**: None - fully backward compatible with V2.x and V1

### V2
- **Framework**: .NET 10.0
- **C# Version**: 14
- **Package Version**: 2.0.1, 2.0.0
- **Features**: All V1 features with modern .NET 10 runtime performance improvements, plus automatic totals for all numeric types
- **Breaking Changes**: None - fully backward compatible API

### V1
- **Framework**: .NET 9.0
- **Package Version**: 1.0.0
- **Status**: Legacy (still available on NuGet)

## Supported Frameworks

- **.NET 10.0** (with C# 14 support) - **Current V2**
- **.NET 9.0** - Legacy V1

## Installation

### For .NET 10 Projects (V3 - Recommended)

```bash
dotnet add package Faysil.ExcelGenerator --version 3.0.0
```

Or via NuGet Package Manager:

```powershell
Install-Package Faysil.ExcelGenerator -Version 3.0.0
```

### For .NET 9 Projects (V1 - Legacy)

```bash
dotnet add package Faysil.ExcelGenerator --version 1.0.0
```

## Features

### Core Features
- ✅ Generate Excel files from any `IEnumerable<T>` or `List<T>`
- ✅ Fluent configuration API for advanced scenarios
- ✅ Simple API for basic use cases (backward compatible)
- ✅ Auto-formatted column headers (PascalCase to spaced text)
- ✅ Auto-fit column widths
- ✅ Multiple output formats: File, Byte Array, Stream, or XLWorkbook

### Advanced Features (V3)
- ✅ **Multiple Aggregations**: Sum, Average, Min, Max, Count for all numeric columns
- ✅ **Conditional Formatting**: Color scales, data bars, highlight rules
- ✅ **Multi-Sheet Workbooks**: Create workbooks with multiple sheets in one call
- ✅ **Freeze Panes**: Freeze header rows and columns for easier navigation
- ✅ **Customizable Colors**: Set header and aggregation row colors
- ✅ **Column Filtering**: Option to exclude columns ending with "Id"

## Quick Start

### Basic Usage

```csharp
using ExcelGenerator;
using ClosedXML.Excel;

// Your data
var products = new List<Product>
{
    new Product { ProductId = 1, Name = "Laptop", Price = 999.99m, Quantity = 10 },
    new Product { ProductId = 2, Name = "Mouse", Price = 29.99m, Quantity = 50 },
    new Product { ProductId = 3, Name = "Keyboard", Price = 79.99m, Quantity = 30 }
};

// Generate and save to file
ExcelSheetGenerator.GenerateExcelFile(
    data: products,
    sheetName: "Products",
    filePath: "products.xlsx",
    excludeIds: true,           // Removes ProductId column
    headerColor: XLColor.Green  // Custom header color
);
```

### Get as Byte Array (for web downloads)

```csharp
byte[] excelBytes = ExcelSheetGenerator.GenerateExcelBytes(
    data: products,
    sheetName: "Products",
    excludeIds: true
);

// In ASP.NET Core
return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "products.xlsx");
```

### Get as Stream

```csharp
using var stream = ExcelSheetGenerator.GenerateExcelStream(
    data: products,
    sheetName: "Products"
);
```

### Get XLWorkbook (for advanced customization)

```csharp
using var workbook = ExcelSheetGenerator.GenerateExcel(
    data: products,
    sheetName: "Products",
    excludeIds: false,
    headerColor: XLColor.LightBlue
);

// Add more sheets, customize, etc.
workbook.SaveAs("output.xlsx");
```

## Advanced Usage (V3)

### Using Fluent Configuration API

```csharp
using ExcelGenerator;
using ClosedXML.Excel;

// Configure with multiple aggregations
var workbook = ExcelSheetGenerator
    .Configure<Product>()
    .WithData(products, "Products")
    .WithAggregations(AggregationType.Sum | AggregationType.Average | AggregationType.Count)
    .WithExcludeIds()
    .WithHeaderColor(XLColor.LightBlue)
    .FreezeHeaderRow()
    .GenerateExcel();

workbook.SaveAs("products-advanced.xlsx");
```

### Conditional Formatting

```csharp
var workbook = ExcelSheetGenerator
    .Configure<SalesData>()
    .WithData(salesData, "Sales")
    .WithConditionalFormatting(fmt => fmt
        .HighlightNegatives("Profit")        // Red background for negative profits
        .ColorScale("Revenue", XLColor.Red, XLColor.Green)  // Color gradient
        .DataBars("Quantity")                // Data bars for quantity
        .HighlightTopN("Sales", 10))         // Highlight top 10 sales
    .FreezeHeaderRow()
    .GenerateExcel();

workbook.SaveAs("sales-formatted.xlsx");
```

### Multiple Sheets in One Workbook

```csharp
var workbook = new ExcelWorkbookBuilder()
    .AddSheet("Products", products, cfg => cfg
        .WithAggregations(AggregationType.Sum | AggregationType.Average)
        .WithHeaderColor(XLColor.LightBlue)
        .FreezeHeaderRow())
    .AddSheet("Orders", orders, cfg => cfg
        .WithAggregations(AggregationType.Sum | AggregationType.Count)
        .WithConditionalFormatting(fmt => fmt.HighlightNegatives("Total"))
        .WithHeaderColor(XLColor.LightGreen)
        .FreezeHeaderRow())
    .AddSheet("Customers", customers, cfg => cfg
        .WithExcludeIds()
        .WithHeaderColor(XLColor.LightYellow))
    .Build();

workbook.SaveAs("multi-sheet-report.xlsx");
```

### All Aggregation Types

```csharp
// Generate report with all aggregations
var workbook = ExcelSheetGenerator
    .Configure<FinancialData>()
    .WithData(financialData, "Financial Report")
    .WithAggregations(
        AggregationType.Sum |       // Total
        AggregationType.Average |   // Average
        AggregationType.Min |       // Minimum
        AggregationType.Max |       // Maximum
        AggregationType.Count)      // Count
    .WithHeaderColor(XLColor.DarkBlue)
    .FreezePanes(rowsToFreeze: 1, columnsToFreeze: 2)
    .GenerateExcelFile("financial-report.xlsx");
```

### Conditional Formatting Options

```csharp
var config = ExcelSheetGenerator
    .Configure<Product>()
    .WithData(products, "Products")
    .WithConditionalFormatting(fmt => fmt
        .HighlightNegatives("Stock")           // Highlight negative values
        .HighlightPositives("Profit")          // Highlight positive values
        .ColorScale("Price")                   // Color gradient (red to green)
        .DataBars("Quantity")                  // Data bars
        .HighlightDuplicates("SKU")            // Highlight duplicates
        .HighlightTopN("Revenue", topN: 5));   // Highlight top 5

config.GenerateExcelFile("products-formatted.xlsx");
```

## Parameters

| Parameter | Type | Description |
|-----------|------|-------------|
| `data` | `IEnumerable<T>` | The collection of objects to export |
| `sheetName` | `string` | Name of the Excel worksheet |
| `excludeIds` | `bool` | If `true`, excludes columns ending with "Id" or "ID" |
| `headerColor` | `XLColor?` | Background color for header row (default: LightBlue) |

## Output Features

- **Headers**: Automatically formatted from PascalCase (e.g., `ProductName` → `Product Name`)
- **Numbers**: Formatted with thousand separators
- **Decimals**: Displayed with 2 decimal places
- **Dates**: Formatted as `yyyy-MM-dd HH:mm:ss`
- **Booleans**: Displayed as "Yes" or "No"
- **Summation Row**: Automatically added for all numeric columns (decimal, double, float, int, long, short, byte) at the bottom

## Dependencies

- [ClosedXML](https://github.com/ClosedXML/ClosedXML) v0.105.0 (latest stable version)
  - Compatible with .NET Standard 2.0+
  - Works seamlessly with .NET 10

## What's New in V3.0.0?

### Major Features
- **Fluent Configuration API**: New `ExcelConfiguration<T>` builder pattern for advanced scenarios
  ```csharp
  ExcelSheetGenerator.Configure<T>()
      .WithData(data, "SheetName")
      .WithAggregations(AggregationType.Sum | AggregationType.Average)
      .WithConditionalFormatting(...)
      .GenerateExcel();
  ```

- **Multiple Aggregation Types**: Beyond Sum, now supports:
  - `Sum` - Total of all values (light gray background)
  - `Average` - Mean of all values (alice blue background)
  - `Min` - Minimum value (light yellow background)
  - `Max` - Maximum value (light green background)
  - `Count` - Number of rows (lavender background)
  - Combine multiple: `AggregationType.Sum | AggregationType.Average`

- **Conditional Formatting**: Six predefined formatting rules
  - `HighlightNegatives(columnName)` - Red background for values < 0
  - `HighlightPositives(columnName)` - Green background for values > 0
  - `ColorScale(columnName, minColor, maxColor)` - Gradient from red to green
  - `DataBars(columnName, barColor)` - Excel data bars
  - `HighlightDuplicates(columnName)` - Yellow background for duplicates
  - `HighlightTopN(columnName, topN)` - Green background for top N values

- **Multi-Sheet Workbooks**: New `ExcelWorkbookBuilder` class
  ```csharp
  new ExcelWorkbookBuilder()
      .AddSheet("Sheet1", data1, config1)
      .AddSheet("Sheet2", data2, config2)
      .Build();
  ```

- **Freeze Panes**: Lock rows and columns for easier navigation
  - `FreezeHeaderRow()` - Freeze first row only
  - `FreezePanes(rows, columns)` - Freeze specific rows and columns

### Backward Compatibility
- ✅ All V2.x and V1 code continues to work without changes
- ✅ Simple API methods remain unchanged
- ✅ New features are opt-in through fluent configuration

### API Additions
- `ExcelSheetGenerator.Configure<T>()` - Entry point for fluent API
- `ExcelConfiguration<T>` - Builder class for configuration
- `ExcelWorkbookBuilder` - Multi-sheet workbook builder
- `ConditionalFormattingConfiguration` - Formatting rules configuration
- `AggregationType` - Enum for aggregation types (flags)

## What's New in V2.0.1?

### New Features
- **All Numeric Types Totals**: Automatic summation row now supports ALL numeric types, not just decimal
  - Supported types: `decimal`, `double`, `float`, `int`, `long`, `short`, `byte`
  - Floating-point numbers display with 2 decimal places
  - Integer types display without decimals
- **RefineValue Extension**: New public extension method for precise decimal calculations
  - Truncates to 3 decimal places instead of rounding
  - Available for use in your own code via `NumericExtensions.RefineValue()`
  - Applied automatically to decimal and double totals

### Improvements
- Enhanced number formatting consistency across all numeric types
- More accurate summation using truncation for floating-point values
- Better handling of mixed numeric column types

## What's New in V2?

### Performance Improvements
- **Native .NET 10 Runtime**: Benefits from improved JIT compilation, faster stack allocations, and enhanced code generation
- **AVX10.2 & ARM64 SVE Support**: Automatic use of advanced CPU instructions for better performance
- **Smaller Footprint**: Leverages .NET 10's optimized runtime

### Developer Experience
- **C# 14 Features**: Access to the latest language features:
  - Extension members & blocks
  - Implicit span conversions for better memory efficiency
  - Null-conditional assignment operators
  - Enhanced partial types support

### Compatibility
- **Long-Term Support**: .NET 10 is an LTS release supported until November 2028
- **Backward Compatible**: Same API as V1 - no code changes needed for migration
- **Modern Tooling**: Full support in Visual Studio 2026 and latest .NET CLI

## Migration from V1 to V2

Upgrading from V1 to V2 is straightforward:

1. **Update your project** to target .NET 10:
   ```xml
   <TargetFramework>net10.0</TargetFramework>
   ```

2. **Update the package reference**:
   ```bash
   dotnet add package Faysil.ExcelGenerator --version 2.0.0
   ```

3. **No code changes required** - V2 maintains 100% API compatibility with V1

### Why Upgrade to V2?

- ✅ **Better Performance**: Native .NET 10 runtime optimizations
- ✅ **Long-Term Support**: LTS release with support until 2028
- ✅ **Modern Features**: Access to C# 14 language improvements
- ✅ **Future-Proof**: Stay current with the latest .NET ecosystem

## License

MIT License

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.
