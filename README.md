# ExcelGenerator

A lightweight .NET library to generate Excel files from `IEnumerable<T>` collections using ClosedXML.

## Supported Frameworks

- **.NET 9.0**

## Installation

```bash
dotnet add package ExcelGenerator
```

Or via NuGet Package Manager:

```powershell
Install-Package ExcelGenerator
```

## Features

- ✅ Generate Excel files from any `IEnumerable<T>` or `List<T>`
- ✅ Customizable header colors
- ✅ Option to exclude columns ending with "Id"
- ✅ Automatic summation row for decimal columns
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
- **Summation Row**: Automatically added for decimal columns at the bottom

## Dependencies

- [ClosedXML](https://github.com/ClosedXML/ClosedXML) (v0.105.0+)

## License

MIT License

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.
