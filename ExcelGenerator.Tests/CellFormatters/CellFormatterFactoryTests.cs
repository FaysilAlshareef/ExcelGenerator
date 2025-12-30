using ClosedXML.Excel;
using ExcelGenerator.Core.CellFormatters;

namespace ExcelGenerator.Tests.CellFormatters;

public class CellFormatterFactoryTests
{
    private readonly CellFormatterFactory _factory;
    private readonly XLWorkbook _workbook;
    private readonly IXLWorksheet _worksheet;

    public CellFormatterFactoryTests()
    {
        _factory = new CellFormatterFactory();
        _workbook = new XLWorkbook();
        _worksheet = _workbook.Worksheets.Add("Test");
    }

    [Fact]
    public void FormatCell_WithDecimal_AppliesCorrectFormat()
    {
        // Arrange
        var cell = _worksheet.Cell(1, 1);
        decimal value = 1234.56m;

        // Act
        _factory.FormatCell(cell, value, typeof(decimal));

        // Assert
        Assert.Equal(1234.56, cell.GetValue<double>());
        Assert.Equal("#,##0.00", cell.Style.NumberFormat.Format);
    }

    [Fact]
    public void FormatCell_WithNullableDecimal_AppliesCorrectFormat()
    {
        // Arrange
        var cell = _worksheet.Cell(1, 1);
        decimal? value = 999.99m;

        // Act
        _factory.FormatCell(cell, value, typeof(decimal?));

        // Assert
        Assert.Equal(999.99, cell.GetValue<double>());
        Assert.Equal("#,##0.00", cell.Style.NumberFormat.Format);
    }

    [Fact]
    public void FormatCell_WithDouble_AppliesCorrectFormat()
    {
        // Arrange
        var cell = _worksheet.Cell(1, 1);
        double value = 123.456;

        // Act
        _factory.FormatCell(cell, value, typeof(double));

        // Assert
        Assert.Equal(123.456, cell.GetValue<double>());
        Assert.Equal("#,##0.00", cell.Style.NumberFormat.Format);
    }

    [Fact]
    public void FormatCell_WithFloat_AppliesCorrectFormat()
    {
        // Arrange
        var cell = _worksheet.Cell(1, 1);
        float value = 78.9f;

        // Act
        _factory.FormatCell(cell, value, typeof(float));

        // Assert
        Assert.Equal(78.9, cell.GetValue<double>(), 1); // 1 decimal tolerance
        Assert.Equal("#,##0.00", cell.Style.NumberFormat.Format);
    }

    [Fact]
    public void FormatCell_WithInteger_AppliesCorrectFormat()
    {
        // Arrange
        var cell = _worksheet.Cell(1, 1);
        int value = 42;

        // Act
        _factory.FormatCell(cell, value, typeof(int));

        // Assert
        Assert.Equal(42, cell.GetValue<int>());
        Assert.Equal("#,##0", cell.Style.NumberFormat.Format);
    }

    [Fact]
    public void FormatCell_WithLong_AppliesCorrectFormat()
    {
        // Arrange
        var cell = _worksheet.Cell(1, 1);
        long value = 1000000L;

        // Act
        _factory.FormatCell(cell, value, typeof(long));

        // Assert
        Assert.Equal(1000000, cell.GetValue<long>());
        Assert.Equal("#,##0", cell.Style.NumberFormat.Format);
    }

    [Fact]
    public void FormatCell_WithShort_AppliesCorrectFormat()
    {
        // Arrange
        var cell = _worksheet.Cell(1, 1);
        short value = 100;

        // Act
        _factory.FormatCell(cell, value, typeof(short));

        // Assert
        Assert.Equal(100, cell.GetValue<short>());
        Assert.Equal("#,##0", cell.Style.NumberFormat.Format);
    }

    [Fact]
    public void FormatCell_WithByte_AppliesCorrectFormat()
    {
        // Arrange
        var cell = _worksheet.Cell(1, 1);
        byte value = 255;

        // Act
        _factory.FormatCell(cell, value, typeof(byte));

        // Assert
        Assert.Equal(255, cell.GetValue<byte>());
        Assert.Equal("#,##0", cell.Style.NumberFormat.Format);
    }

    [Fact]
    public void FormatCell_WithDateTime_AppliesCorrectFormat()
    {
        // Arrange
        var cell = _worksheet.Cell(1, 1);
        var value = new DateTime(2025, 12, 30, 14, 30, 0);

        // Act
        _factory.FormatCell(cell, value, typeof(DateTime));

        // Assert
        Assert.Equal(value, cell.GetValue<DateTime>());
        Assert.Equal("yyyy-MM-dd HH:mm:ss", cell.Style.NumberFormat.Format);
    }

    [Fact]
    public void FormatCell_WithDateOnly_AppliesCorrectFormat()
    {
        // Arrange
        var cell = _worksheet.Cell(1, 1);
        var value = new DateOnly(2025, 12, 30);

        // Act
        _factory.FormatCell(cell, value, typeof(DateOnly));

        // Assert
        var cellValue = cell.GetString();
        Assert.Contains("2025", cellValue);
        Assert.Contains("12", cellValue);
        Assert.Contains("30", cellValue);
    }

    [Fact]
    public void FormatCell_WithBoolean_True_FormatsAsYes()
    {
        // Arrange
        var cell = _worksheet.Cell(1, 1);
        bool value = true;

        // Act
        _factory.FormatCell(cell, value, typeof(bool));

        // Assert
        Assert.Equal("Yes", cell.GetString());
    }

    [Fact]
    public void FormatCell_WithBoolean_False_FormatsAsNo()
    {
        // Arrange
        var cell = _worksheet.Cell(1, 1);
        bool value = false;

        // Act
        _factory.FormatCell(cell, value, typeof(bool));

        // Assert
        Assert.Equal("No", cell.GetString());
    }

    [Fact]
    public void FormatCell_WithString_SetsValue()
    {
        // Arrange
        var cell = _worksheet.Cell(1, 1);
        string value = "Hello World";

        // Act
        _factory.FormatCell(cell, value, typeof(string));

        // Assert
        Assert.Equal("Hello World", cell.GetString());
    }

    [Fact]
    public void FormatCell_WithNull_SetsEmptyString()
    {
        // Arrange
        var cell = _worksheet.Cell(1, 1);

        // Act
        _factory.FormatCell(cell, null, typeof(string));

        // Assert
        Assert.Equal("", cell.GetString());
    }

    [Fact]
    public void FormatCell_WithNullableInt_Null_SetsEmptyString()
    {
        // Arrange
        var cell = _worksheet.Cell(1, 1);
        int? value = null;

        // Act
        _factory.FormatCell(cell, value, typeof(int?));

        // Assert
        Assert.Equal("", cell.GetString());
    }

    [Fact]
    public void FormatCell_WithCustomObject_UsesToString()
    {
        // Arrange
        var cell = _worksheet.Cell(1, 1);
        var value = new TestClass { Name = "Test" };

        // Act
        _factory.FormatCell(cell, value, typeof(TestClass));

        // Assert
        Assert.Equal("TestClass: Test", cell.GetString());
    }

    private class TestClass
    {
        public string Name { get; set; } = "";
        public override string ToString() => $"TestClass: {Name}";
    }

}
