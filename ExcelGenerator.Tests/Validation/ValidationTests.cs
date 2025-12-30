using ExcelGenerator.Core;
using ExcelGenerator.Core.Generators;
using ExcelGenerator.Core.CellFormatters;
using ExcelGenerator.Core.Aggregation;
using ExcelGenerator.Core.ConditionalFormatting;
using ExcelGenerator.Core.PropertyReflection;
using Xunit;

namespace ExcelGenerator.Tests.Validation;

public class ValidationTests
{
    private readonly ExcelGeneratorEngine _engine;

    public ValidationTests()
    {
        // Create engine with all dependencies
        var propertyExtractor = new PropertyExtractor();
        var cellFormatterFactory = new CellFormatterFactory();
        var aggregationFactory = new AggregationStrategyFactory();
        var formattingFactory = new FormattingRuleApplierFactory();
        var headerGenerator = new HeaderGenerator(propertyExtractor);
        var dataRowGenerator = new DataRowGenerator(cellFormatterFactory);
        var aggregationGenerator = new AggregationRowGenerator(aggregationFactory);
        var layoutManager = new WorksheetLayoutManager();

        _engine = new ExcelGeneratorEngine(
            propertyExtractor,
            headerGenerator,
            dataRowGenerator,
            aggregationGenerator,
            formattingFactory,
            layoutManager);
    }

    [Fact]
    public void Generate_WithNullData_ThrowsArgumentNullException()
    {
        // Arrange
        var config = new ExcelConfiguration<Product>();

        // Act & Assert
        var exception = Assert.Throws<ArgumentNullException>(() =>
            _engine.Generate<Product>(null!, "Sheet1", config));

        Assert.Contains("Data collection cannot be null", exception.Message);
    }

    [Fact]
    public void Generate_WithNullSheetName_ThrowsArgumentException()
    {
        // Arrange
        var data = new List<Product>();
        var config = new ExcelConfiguration<Product>();

        // Act & Assert
        var exception = Assert.Throws<ArgumentException>(() =>
            _engine.Generate(data, null!, config));

        Assert.Contains("Sheet name cannot be null or empty", exception.Message);
    }

    [Fact]
    public void Generate_WithEmptySheetName_ThrowsArgumentException()
    {
        // Arrange
        var data = new List<Product>();
        var config = new ExcelConfiguration<Product>();

        // Act & Assert
        var exception = Assert.Throws<ArgumentException>(() =>
            _engine.Generate(data, "", config));

        Assert.Contains("Sheet name cannot be null or empty", exception.Message);
    }

    [Fact]
    public void Generate_WithWhitespaceSheetName_ThrowsArgumentException()
    {
        // Arrange
        var data = new List<Product>();
        var config = new ExcelConfiguration<Product>();

        // Act & Assert
        var exception = Assert.Throws<ArgumentException>(() =>
            _engine.Generate(data, "   ", config));

        Assert.Contains("Sheet name cannot be null or empty", exception.Message);
    }

    [Fact]
    public void Generate_WithSheetNameTooLong_ThrowsArgumentException()
    {
        // Arrange
        var data = new List<Product>();
        var config = new ExcelConfiguration<Product>();
        var longName = new string('A', 32); // 32 characters (max is 31)

        // Act & Assert
        var exception = Assert.Throws<ArgumentException>(() =>
            _engine.Generate(data, longName, config));

        Assert.Contains("exceeds maximum length of 31 characters", exception.Message);
        Assert.Contains("Current length: 32", exception.Message);
    }

    [Theory]
    [InlineData("Sheet:Name", ':')]
    [InlineData("Sheet\\Name", '\\')]
    [InlineData("Sheet/Name", '/')]
    [InlineData("Sheet?Name", '?')]
    [InlineData("Sheet*Name", '*')]
    [InlineData("Sheet[Name]", '[')]
    [InlineData("Sheet]Name", ']')]
    public void Generate_WithInvalidCharacterInSheetName_ThrowsArgumentException(string sheetName, char invalidChar)
    {
        // Arrange
        var data = new List<Product>();
        var config = new ExcelConfiguration<Product>();

        // Act & Assert
        var exception = Assert.Throws<ArgumentException>(() =>
            _engine.Generate(data, sheetName, config));

        Assert.Contains($"contains invalid character '{invalidChar}'", exception.Message);
        Assert.Contains("Excel sheet names cannot contain", exception.Message);
    }

    [Fact]
    public void Generate_WithNullConfiguration_ThrowsArgumentNullException()
    {
        // Arrange
        var data = new List<Product>();

        // Act & Assert
        var exception = Assert.Throws<ArgumentNullException>(() =>
            _engine.Generate(data, "Sheet1", null!));

        Assert.Contains("Configuration cannot be null", exception.Message);
    }

    [Fact]
    public void Generate_WithValidMaxLengthSheetName_Succeeds()
    {
        // Arrange
        var data = new List<Product> { new Product { Name = "Test" } };
        var config = new ExcelConfiguration<Product>();
        var maxLengthName = new string('A', 31); // Exactly 31 characters

        // Act
        var workbook = _engine.Generate(data, maxLengthName, config);

        // Assert
        Assert.NotNull(workbook);
        Assert.Single(workbook.Worksheets);
    }

    [Fact]
    public void Generate_WithEmptyData_ReturnsWorkbookWithHeadersOnly()
    {
        // Arrange
        var data = new List<Product>();
        var config = new ExcelConfiguration<Product>();

        // Act
        var workbook = _engine.Generate(data, "Sheet1", config);

        // Assert
        Assert.NotNull(workbook);
        var worksheet = workbook.Worksheets.First();
        Assert.NotNull(worksheet);
        // Should have headers but no data rows
        Assert.False(worksheet.Cell(2, 1).IsEmpty() == false); // Row 2 should be empty
    }

    [Fact]
    public void Generate_WithClassWithNoProperties_ThrowsInvalidOperationException()
    {
        // Arrange
        var data = new List<EmptyClass> { new EmptyClass() };
        var config = new ExcelConfiguration<EmptyClass>();

        // Act & Assert
        var exception = Assert.Throws<InvalidOperationException>(() =>
            _engine.Generate(data, "Sheet1", config));

        Assert.Contains("has no readable properties", exception.Message);
        Assert.Contains("Cannot generate Excel sheet", exception.Message);
    }

    [Fact]
    public void Generate_WithValidData_Succeeds()
    {
        // Arrange
        var data = new List<Product>
        {
            new Product { ProductId = 1, Name = "Test", Price = 10.5m }
        };
        var config = new ExcelConfiguration<Product>();

        // Act
        var workbook = _engine.Generate(data, "Sheet1", config);

        // Assert
        Assert.NotNull(workbook);
        var worksheet = workbook.Worksheets.First();
        Assert.NotNull(worksheet);

        // Verify headers exist
        Assert.False(string.IsNullOrEmpty(worksheet.Cell(1, 1).GetString()));

        // Verify data exists
        Assert.False(worksheet.Cell(2, 1).IsEmpty());
    }

    [Fact]
    public void Generate_WithSpecialCharactersInData_HandlesCorrectly()
    {
        // Arrange
        var data = new List<Product>
        {
            new Product { Name = "Test\"Quote", Price = 10.5m },
            new Product { Name = "Test'Apostrophe", Price = 20.5m },
            new Product { Name = "Test<>Brackets", Price = 30.5m }
        };
        var config = new ExcelConfiguration<Product>();

        // Act
        var workbook = _engine.Generate(data, "Sheet1", config);

        // Assert
        Assert.NotNull(workbook);
        var worksheet = workbook.Worksheets.First();

        // Verify data was written correctly
        Assert.Contains("Quote", worksheet.Cell(2, 2).GetString());
        Assert.Contains("Apostrophe", worksheet.Cell(3, 2).GetString());
        Assert.Contains("Brackets", worksheet.Cell(4, 2).GetString());
    }

    [Fact]
    public void Generate_WithNullValuesInData_HandlesCorrectly()
    {
        // Arrange
        var data = new List<Product>
        {
            new Product { Name = null, Price = 10.5m },
            new Product { Name = "Test", NullablePrice = null }
        };
        var config = new ExcelConfiguration<Product>();

        // Act
        var workbook = _engine.Generate(data, "Sheet1", config);

        // Assert
        Assert.NotNull(workbook);
        var worksheet = workbook.Worksheets.First();

        // Null string should be empty
        Assert.Equal("", worksheet.Cell(2, 2).GetString());

        // Null nullable decimal should be empty
        Assert.True(worksheet.Cell(3, 4).IsEmpty() || worksheet.Cell(3, 4).GetString() == "");
    }

    // Test models
    private class Product
    {
        public int ProductId { get; set; }
        public string? Name { get; set; }
        public decimal Price { get; set; }
        public decimal? NullablePrice { get; set; }
    }

    private class EmptyClass
    {
        // No public readable properties
    }
}
