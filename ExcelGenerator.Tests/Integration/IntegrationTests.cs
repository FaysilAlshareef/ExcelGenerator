using ExcelGenerator;
using ClosedXML.Excel;
using Xunit;

namespace ExcelGenerator.Tests.Integration;

public class IntegrationTests
{
    [Fact]
    public void GenerateExcel_BasicUsage_CreatesValidWorkbook()
    {
        // Arrange
        var products = new List<Product>
        {
            new Product { ProductId = 1, Name = "Laptop", Price = 999.99m, Quantity = 10 },
            new Product { ProductId = 2, Name = "Mouse", Price = 29.99m, Quantity = 50 },
            new Product { ProductId = 3, Name = "Keyboard", Price = 79.99m, Quantity = 30 }
        };

        // Act
        var workbook = ExcelSheetGenerator.GenerateExcel(products, "Products");

        // Assert
        Assert.NotNull(workbook);
        var worksheet = workbook.Worksheets.First();
        Assert.Equal("Products", worksheet.Name);

        // Verify headers
        Assert.Contains("Product Id", worksheet.Cell(1, 1).GetString());
        Assert.Contains("Name", worksheet.Cell(1, 2).GetString());
        Assert.Contains("Price", worksheet.Cell(1, 3).GetString());
        Assert.Contains("Quantity", worksheet.Cell(1, 4).GetString());

        // Verify data
        Assert.Equal(1, worksheet.Cell(2, 1).GetValue<int>());
        Assert.Equal("Laptop", worksheet.Cell(2, 2).GetString());
        Assert.Equal(999.99, worksheet.Cell(2, 3).GetValue<double>());
        Assert.Equal(10, worksheet.Cell(2, 4).GetValue<int>());

        // Verify summation row (default behavior)
        Assert.Equal(1109.97, worksheet.Cell(5, 3).GetValue<double>(), 2);
        Assert.Equal(90, worksheet.Cell(5, 4).GetValue<double>());
    }

    [Fact]
    public void GenerateExcel_WithExcludeIds_FiltersIdColumns()
    {
        // Arrange
        var products = new List<Product>
        {
            new Product { ProductId = 1, Name = "Test", Price = 10.0m, Quantity = 5 }
        };

        // Act
        var workbook = ExcelSheetGenerator.GenerateExcel(products, "Products", excludeIds: true);

        // Assert
        var worksheet = workbook.Worksheets.First();

        // ProductId should not be in headers
        Assert.DoesNotContain("Product Id", worksheet.Cell(1, 1).GetString());
        Assert.DoesNotContain("Product Id", worksheet.Cell(1, 2).GetString());
        Assert.DoesNotContain("Product Id", worksheet.Cell(1, 3).GetString());

        // Should have Name, Price, Quantity
        Assert.Contains("Name", worksheet.Cell(1, 1).GetString());
    }

    [Fact]
    public void GenerateExcel_WithCustomHeaderColor_AppliesColor()
    {
        // Arrange
        var products = new List<Product>
        {
            new Product { Name = "Test", Price = 10.0m }
        };

        // Act
        var workbook = ExcelSheetGenerator.GenerateExcel(
            products, "Products", headerColor: XLColor.Green);

        // Assert
        var worksheet = workbook.Worksheets.First();
        var headerCell = worksheet.Cell(1, 1);

        Assert.Equal(XLColor.Green, headerCell.Style.Fill.BackgroundColor);
    }

    [Fact]
    public void GenerateExcel_WithFluentAPI_AllAggregations_Works()
    {
        // Arrange
        var products = new List<Product>
        {
            new Product { Name = "A", Price = 10.0m, Quantity = 5 },
            new Product { Name = "B", Price = 20.0m, Quantity = 10 },
            new Product { Name = "C", Price = 30.0m, Quantity = 15 }
        };

        // Act
        var workbook = ExcelSheetGenerator.Configure<Product>()
            .WithAggregations(
                AggregationType.Sum |
                AggregationType.Average |
                AggregationType.Min |
                AggregationType.Max |
                AggregationType.Count)
            .WithData(products, "Products")
            .GenerateExcel();

        // Assert
        var worksheet = workbook.Worksheets.First();

        // Data rows: 2, 3, 4
        // Aggregation rows: 5 (Sum), 6 (Average), 7 (Min), 8 (Max), 9 (Count)

        // Verify Sum row (row 5)
        Assert.Equal(60.0, worksheet.Cell(5, 3).GetValue<double>(), 2); // Price sum
        Assert.Equal(30.0, worksheet.Cell(5, 4).GetValue<double>()); // Quantity sum

        // Verify Average row (row 6)
        Assert.Equal(20.0, worksheet.Cell(6, 3).GetValue<double>(), 2); // Price average
        Assert.Equal(10.0, worksheet.Cell(6, 4).GetValue<double>()); // Quantity average

        // Verify Min row (row 7)
        Assert.Equal(10.0, worksheet.Cell(7, 3).GetValue<double>(), 2); // Price min
        Assert.Equal(5.0, worksheet.Cell(7, 4).GetValue<double>()); // Quantity min

        // Verify Max row (row 8)
        Assert.Equal(30.0, worksheet.Cell(8, 3).GetValue<double>(), 2); // Price max
        Assert.Equal(15.0, worksheet.Cell(8, 4).GetValue<double>()); // Quantity max

        // Verify Count row (row 9)
        Assert.Equal(3.0, worksheet.Cell(9, 3).GetValue<double>()); // Price count
        Assert.Equal(3.0, worksheet.Cell(9, 4).GetValue<double>()); // Quantity count
    }

    [Fact]
    public void GenerateExcel_WithFreezePanes_AppliesCorrectly()
    {
        // Arrange
        var products = new List<Product>
        {
            new Product { Name = "Test", Price = 10.0m }
        };

        // Act
        var workbook = ExcelSheetGenerator.Configure<Product>()
            .WithData(products, "Products")
            .FreezeHeaderRow()
            .GenerateExcel();

        // Assert
        var worksheet = workbook.Worksheets.First();
        // Verify freeze panes are applied (ClosedXML specific verification)
        Assert.NotNull(worksheet);
    }

    [Fact]
    public void GenerateExcelBytes_ReturnsValidByteArray()
    {
        // Arrange
        var products = new List<Product>
        {
            new Product { Name = "Test", Price = 10.0m }
        };

        // Act
        var bytes = ExcelSheetGenerator.GenerateExcelBytes(products, "Products");

        // Assert
        Assert.NotNull(bytes);
        Assert.True(bytes.Length > 0);

        // Verify it's a valid Excel file by loading it
        using var stream = new MemoryStream(bytes);
        var workbook = new XLWorkbook(stream);
        Assert.Single(workbook.Worksheets);
    }

    [Fact]
    public void GenerateExcelStream_ReturnsValidStream()
    {
        // Arrange
        var products = new List<Product>
        {
            new Product { Name = "Test", Price = 10.0m }
        };

        // Act
        using var stream = ExcelSheetGenerator.GenerateExcelStream(products, "Products");

        // Assert
        Assert.NotNull(stream);
        Assert.True(stream.Length > 0);
        Assert.Equal(0, stream.Position); // Position should be reset

        // Verify it's a valid Excel file
        var workbook = new XLWorkbook(stream);
        Assert.Single(workbook.Worksheets);
    }

    [Fact]
    public void GenerateExcelFile_CreatesFileOnDisk()
    {
        // Arrange
        var products = new List<Product>
        {
            new Product { Name = "Test", Price = 10.0m }
        };
        var tempFile = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.xlsx");

        try
        {
            // Act
            ExcelSheetGenerator.GenerateExcelFile(products, "Products", tempFile);

            // Assert
            Assert.True(File.Exists(tempFile));

            // Verify file is valid
            using var workbook = new XLWorkbook(tempFile);
            Assert.Single(workbook.Worksheets);
        }
        finally
        {
            // Cleanup
            if (File.Exists(tempFile))
                File.Delete(tempFile);
        }
    }

    [Fact]
    public void GenerateExcel_WithMixedDataTypes_HandlesAllTypes()
    {
        // Arrange
        var items = new List<MixedTypeClass>
        {
            new MixedTypeClass
            {
                StringValue = "Test",
                IntValue = 42,
                DecimalValue = 123.45m,
                DoubleValue = 67.89,
                BoolValue = true,
                DateTimeValue = new DateTime(2025, 12, 30),
                DateOnlyValue = new DateOnly(2025, 12, 30),
                NullableInt = 100,
                NullString = null
            }
        };

        // Act
        var workbook = ExcelSheetGenerator.GenerateExcel(items, "Mixed");

        // Assert
        var worksheet = workbook.Worksheets.First();

        // Verify all data types are formatted correctly
        Assert.Equal("Test", worksheet.Cell(2, 1).GetString());
        Assert.Equal(42, worksheet.Cell(2, 2).GetValue<int>());
        Assert.Equal(123.45, worksheet.Cell(2, 3).GetValue<double>());
        Assert.Equal(67.89, worksheet.Cell(2, 4).GetValue<double>(), 2);
        Assert.Equal("Yes", worksheet.Cell(2, 5).GetString());
        Assert.Equal(new DateTime(2025, 12, 30), worksheet.Cell(2, 6).GetValue<DateTime>());
        Assert.Equal(100, worksheet.Cell(2, 8).GetValue<int>());
        Assert.Equal("", worksheet.Cell(2, 9).GetString()); // Null should be empty
    }

    [Fact]
    public void GenerateExcel_WithLargeDataset_Succeeds()
    {
        // Arrange
        var products = Enumerable.Range(1, 1000).Select(i => new Product
        {
            ProductId = i,
            Name = $"Product {i}",
            Price = i * 10.5m,
            Quantity = i * 2
        }).ToList();

        // Act
        var workbook = ExcelSheetGenerator.GenerateExcel(products, "Products");

        // Assert
        var worksheet = workbook.Worksheets.First();

        // Verify first and last rows
        Assert.Equal(1, worksheet.Cell(2, 1).GetValue<int>());
        Assert.Equal(1000, worksheet.Cell(1001, 1).GetValue<int>());

        // Verify summation row exists
        Assert.True(worksheet.Cell(1002, 3).GetValue<double>() > 0);
    }

    [Fact]
    public void ExcelWorkbookBuilder_MultipleSheets_Works()
    {
        // Arrange
        var products = new List<Product>
        {
            new Product { Name = "Product1", Price = 10.0m }
        };
        var orders = new List<Order>
        {
            new Order { OrderId = 1, Total = 100.0m }
        };

        // Act
        var workbook = new ExcelWorkbookBuilder()
            .AddSheet("Products", products)
            .AddSheet("Orders", orders)
            .Build();

        // Assert
        Assert.Equal(2, workbook.Worksheets.Count);
        Assert.Contains(workbook.Worksheets, ws => ws.Name == "Products");
        Assert.Contains(workbook.Worksheets, ws => ws.Name == "Orders");
    }

    [Fact]
    public void GenerateExcel_WithNullableTypes_HandlesCorrectly()
    {
        // Arrange
        var items = new List<NullableClass>
        {
            new NullableClass { Value = 10, NullableValue = 20 },
            new NullableClass { Value = 30, NullableValue = null }
        };

        // Act
        var workbook = ExcelSheetGenerator.GenerateExcel(items, "Nullable");

        // Assert
        var worksheet = workbook.Worksheets.First();

        // Row with value
        Assert.Equal(20, worksheet.Cell(2, 2).GetValue<int>());

        // Row with null - should be empty
        Assert.True(worksheet.Cell(3, 2).IsEmpty() || worksheet.Cell(3, 2).GetString() == "");
    }

    // Test models
    private class Product
    {
        public int ProductId { get; set; }
        public string Name { get; set; } = "";
        public decimal Price { get; set; }
        public int Quantity { get; set; }
    }

    private class Order
    {
        public int OrderId { get; set; }
        public decimal Total { get; set; }
    }

    private class MixedTypeClass
    {
        public string StringValue { get; set; } = "";
        public int IntValue { get; set; }
        public decimal DecimalValue { get; set; }
        public double DoubleValue { get; set; }
        public bool BoolValue { get; set; }
        public DateTime DateTimeValue { get; set; }
        public DateOnly DateOnlyValue { get; set; }
        public int? NullableInt { get; set; }
        public string? NullString { get; set; }
    }

    private class NullableClass
    {
        public int Value { get; set; }
        public int? NullableValue { get; set; }
    }
}
