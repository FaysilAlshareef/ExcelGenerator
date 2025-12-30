using ExcelGenerator.Core.PropertyReflection;

namespace ExcelGenerator.Tests.PropertyReflection;

public class PropertyExtractorTests
{
    private readonly PropertyExtractor _extractor;

    public PropertyExtractorTests()
    {
        _extractor = new PropertyExtractor();
    }

    [Fact]
    public void Extract_WithSimpleClass_ReturnsAllProperties()
    {
        // Act
        var properties = _extractor.Extract<SimpleClass>(excludeIds: false);

        // Assert
        Assert.Equal(3, properties.Length);
        Assert.Contains(properties, p => p.Name == "Id");
        Assert.Contains(properties, p => p.Name == "Name");
        Assert.Contains(properties, p => p.Name == "Value");
    }

    [Fact]
    public void Extract_WithExcludeIds_FiltersIdProperties()
    {
        // Act
        var properties = _extractor.Extract<SimpleClass>(excludeIds: true);

        // Assert
        Assert.Equal(2, properties.Length);
        Assert.DoesNotContain(properties, p => p.Name == "Id");
        Assert.Contains(properties, p => p.Name == "Name");
        Assert.Contains(properties, p => p.Name == "Value");
    }

    [Fact]
    public void Extract_WithMultipleIdColumns_FiltersAll()
    {
        // Act
        var properties = _extractor.Extract<ClassWithMultipleIds>(excludeIds: true);

        // Assert
        Assert.Single(properties);
        Assert.DoesNotContain(properties, p => p.Name == "ProductId");
        Assert.DoesNotContain(properties, p => p.Name == "CategoryID");
        Assert.Contains(properties, p => p.Name == "Name");
    }

    [Fact]
    public void Extract_WithWriteOnlyProperty_ExcludesIt()
    {
        // Act
        var properties = _extractor.Extract<ClassWithWriteOnly>(excludeIds: false);

        // Assert
        Assert.DoesNotContain(properties, p => p.Name == "WriteOnly");
        Assert.Contains(properties, p => p.Name == "ReadWrite");
    }

    [Fact]
    public void Extract_WithNoReadableProperties_ReturnsEmpty()
    {
        // Act
        var properties = _extractor.Extract<ClassWithNoReadable>(excludeIds: false);

        // Assert
        Assert.Empty(properties);
    }

    [Fact]
    public void FormatPropertyName_WithPascalCase_AddsSpaces()
    {
        // Act
        var result = _extractor.FormatPropertyName("ProductName");

        // Assert
        Assert.Equal("Product Name", result);
    }

    [Fact]
    public void FormatPropertyName_WithSingleWord_RemainsUnchanged()
    {
        // Act
        var result = _extractor.FormatPropertyName("Name");

        // Assert
        Assert.Equal("Name", result);
    }

    [Fact]
    public void FormatPropertyName_WithAcronym_HandlesCorrectly()
    {
        // Act
        var result = _extractor.FormatPropertyName("ProductID");

        // Assert
        Assert.Equal("Product ID", result);
    }

    [Fact]
    public void FormatPropertyName_WithMultipleWords_AddsAllSpaces()
    {
        // Act
        var result = _extractor.FormatPropertyName("CustomerFirstName");

        // Assert
        Assert.Equal("Customer First Name", result);
    }

    [Fact]
    public void FormatPropertyName_WithNumberInName_HandlesCorrectly()
    {
        // Act
        var result = _extractor.FormatPropertyName("Product2Price");

        // Assert
        // Numbers don't trigger spacing - only uppercase letters do
        Assert.Equal("Product2Price", result);
    }

    [Fact]
    public void FormatPropertyName_WithEmptyString_ReturnsEmpty()
    {
        // Act
        var result = _extractor.FormatPropertyName("");

        // Assert
        Assert.Equal("", result);
    }

    [Fact]
    public void FormatPropertyName_WithAllUppercase_RemainsUnchanged()
    {
        // Act
        var result = _extractor.FormatPropertyName("PRODUCTNAME");

        // Assert
        // All uppercase doesn't get formatted - stays as-is
        Assert.Equal("PRODUCTNAME", result);
    }

    [Fact]
    public void Extract_WithInheritedProperties_IncludesAll()
    {
        // Act
        var properties = _extractor.Extract<DerivedClass>(excludeIds: false);

        // Assert
        Assert.Contains(properties, p => p.Name == "BaseProperty");
        Assert.Contains(properties, p => p.Name == "DerivedProperty");
    }

    [Fact]
    public void Extract_WithAllNumericTypes_IncludesAll()
    {
        // Act
        var properties = _extractor.Extract<NumericTypesClass>(excludeIds: false);

        // Assert
        Assert.Contains(properties, p => p.Name == "DecimalValue");
        Assert.Contains(properties, p => p.Name == "DoubleValue");
        Assert.Contains(properties, p => p.Name == "FloatValue");
        Assert.Contains(properties, p => p.Name == "IntValue");
        Assert.Contains(properties, p => p.Name == "LongValue");
        Assert.Contains(properties, p => p.Name == "ShortValue");
        Assert.Contains(properties, p => p.Name == "ByteValue");
    }

    [Fact]
    public void Extract_WithNullableTypes_IncludesAll()
    {
        // Act
        var properties = _extractor.Extract<NullableTypesClass>(excludeIds: false);

        // Assert
        Assert.Contains(properties, p => p.Name == "NullableInt");
        Assert.Contains(properties, p => p.Name == "NullableDecimal");
        Assert.Contains(properties, p => p.Name == "NullableDateTime");
    }

    // Test models
    private class SimpleClass
    {
        public int Id { get; set; }
        public string Name { get; set; } = "";
        public decimal Value { get; set; }
    }

    private class ClassWithMultipleIds
    {
        public int ProductId { get; set; }
        public int CategoryID { get; set; }
        public string Name { get; set; } = "";
    }

    private class ClassWithWriteOnly
    {
        public string ReadWrite { get; set; } = "";
        public string WriteOnly { set { } }
    }

    private class ClassWithNoReadable
    {
        public string WriteOnly { set { } }
    }

    private class BaseClass
    {
        public string BaseProperty { get; set; } = "";
    }

    private class DerivedClass : BaseClass
    {
        public string DerivedProperty { get; set; } = "";
    }

    private class NumericTypesClass
    {
        public decimal DecimalValue { get; set; }
        public double DoubleValue { get; set; }
        public float FloatValue { get; set; }
        public int IntValue { get; set; }
        public long LongValue { get; set; }
        public short ShortValue { get; set; }
        public byte ByteValue { get; set; }
    }

    private class NullableTypesClass
    {
        public int? NullableInt { get; set; }
        public decimal? NullableDecimal { get; set; }
        public DateTime? NullableDateTime { get; set; }
    }
}
