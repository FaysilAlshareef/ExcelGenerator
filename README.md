# ExcelGenerator

A lightweight .NET library to generate Excel files from `IEnumerable<T>` collections using ClosedXML.

## Version History

### V2 (Current)
- **Framework**: .NET 10.0
- **C# Version**: 14
- **Package Version**: 2.0.1 (latest), 2.0.0
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

### For .NET 10 Projects (V2 - Recommended)

```bash
dotnet add package Faysil.ExcelGenerator --version 2.0.1
```

Or via NuGet Package Manager:

```powershell
Install-Package Faysil.ExcelGenerator -Version 2.0.1
```

### For .NET 9 Projects (V1 - Legacy)

```bash
dotnet add package Faysil.ExcelGenerator --version 1.0.0
```

## Features

- ✅ Generate Excel files from any `IEnumerable<T>` or `List<T>`
- ✅ Customizable header colors
- ✅ Option to exclude columns ending with "Id"
- ✅ Automatic summation row for all numeric columns (decimal, double, float, int, long, short, byte)
- ✅ Auto-formatted column headers (PascalCase to spaced text)
- ✅ Auto-fit column widths
- ✅ Multiple output formats: File, Byte Array, Stream, or XLWorkbook

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
