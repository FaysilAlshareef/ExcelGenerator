using ExcelGenerator.Core.Aggregation;
using System.Reflection;
using Xunit;

namespace ExcelGenerator.Tests.Aggregation;

public class AggregationStrategyTests
{
    private readonly AggregationStrategyFactory _factory;

    public AggregationStrategyTests()
    {
        _factory = new AggregationStrategyFactory();
    }

    #region Sum Tests

    [Fact]
    public void SumStrategy_WithDecimalValues_CalculatesCorrectSum()
    {
        // Arrange
        var data = new List<Product>
        {
            new Product { Price = 10.5m },
            new Product { Price = 20.5m },
            new Product { Price = 30.0m }
        };
        var property = typeof(Product).GetProperty(nameof(Product.Price))!;
        var strategy = _factory.GetStrategy(AggregationType.Sum);

        // Act
        var result = strategy.Calculate(data, property, typeof(decimal));

        // Assert
        Assert.Equal(61.0, result);
    }

    [Fact]
    public void SumStrategy_WithIntegerValues_CalculatesCorrectSum()
    {
        // Arrange
        var data = new List<Product>
        {
            new Product { Quantity = 10 },
            new Product { Quantity = 20 },
            new Product { Quantity = 30 }
        };
        var property = typeof(Product).GetProperty(nameof(Product.Quantity))!;
        var strategy = _factory.GetStrategy(AggregationType.Sum);

        // Act
        var result = strategy.Calculate(data, property, typeof(int));

        // Assert
        Assert.Equal(60.0, result);
    }

    [Fact]
    public void SumStrategy_WithDoubleValues_CalculatesCorrectSum()
    {
        // Arrange
        var data = new List<Product>
        {
            new Product { Weight = 1.5 },
            new Product { Weight = 2.5 },
            new Product { Weight = 3.0 }
        };
        var property = typeof(Product).GetProperty(nameof(Product.Weight))!;
        var strategy = _factory.GetStrategy(AggregationType.Sum);

        // Act
        var result = strategy.Calculate(data, property, typeof(double));

        // Assert
        Assert.Equal(7.0, result, 2);
    }

    [Fact]
    public void SumStrategy_WithNullableValues_SkipsNulls()
    {
        // Arrange
        var data = new List<Product>
        {
            new Product { NullablePrice = 10.0m },
            new Product { NullablePrice = null },
            new Product { NullablePrice = 20.0m }
        };
        var property = typeof(Product).GetProperty(nameof(Product.NullablePrice))!;
        var strategy = _factory.GetStrategy(AggregationType.Sum);

        // Act
        var result = strategy.Calculate(data, property, typeof(decimal));

        // Assert
        Assert.Equal(30.0, result);
    }

    [Fact]
    public void SumStrategy_WithEmptyList_ReturnsZero()
    {
        // Arrange
        var data = new List<Product>();
        var property = typeof(Product).GetProperty(nameof(Product.Price))!;
        var strategy = _factory.GetStrategy(AggregationType.Sum);

        // Act
        var result = strategy.Calculate(data, property, typeof(decimal));

        // Assert
        Assert.Equal(0.0, result);
    }

    #endregion

    #region Average Tests

    [Fact]
    public void AverageStrategy_WithDecimalValues_CalculatesCorrectAverage()
    {
        // Arrange
        var data = new List<Product>
        {
            new Product { Price = 10.0m },
            new Product { Price = 20.0m },
            new Product { Price = 30.0m }
        };
        var property = typeof(Product).GetProperty(nameof(Product.Price))!;
        var strategy = _factory.GetStrategy(AggregationType.Average);

        // Act
        var result = strategy.Calculate(data, property, typeof(decimal));

        // Assert
        Assert.Equal(20.0, result);
    }

    [Fact]
    public void AverageStrategy_WithIntegerValues_CalculatesCorrectAverage()
    {
        // Arrange
        var data = new List<Product>
        {
            new Product { Quantity = 10 },
            new Product { Quantity = 20 },
            new Product { Quantity = 30 }
        };
        var property = typeof(Product).GetProperty(nameof(Product.Quantity))!;
        var strategy = _factory.GetStrategy(AggregationType.Average);

        // Act
        var result = strategy.Calculate(data, property, typeof(int));

        // Assert
        Assert.Equal(20.0, result);
    }

    [Fact]
    public void AverageStrategy_WithNullableValues_DividesByTotalCount()
    {
        // Arrange
        var data = new List<Product>
        {
            new Product { NullablePrice = 10.0m },
            new Product { NullablePrice = null },
            new Product { NullablePrice = 30.0m }
        };
        var property = typeof(Product).GetProperty(nameof(Product.NullablePrice))!;
        var strategy = _factory.GetStrategy(AggregationType.Average);

        // Act
        var result = strategy.Calculate(data, property, typeof(decimal));

        // Assert
        // Average divides by total count, treating null as 0
        // (10 + 0 + 30) / 3 = 40 / 3 = 13.333...
        Assert.Equal(13.333, result, 3);
    }

    [Fact]
    public void AverageStrategy_WithEmptyList_ReturnsZero()
    {
        // Arrange
        var data = new List<Product>();
        var property = typeof(Product).GetProperty(nameof(Product.Price))!;
        var strategy = _factory.GetStrategy(AggregationType.Average);

        // Act
        var result = strategy.Calculate(data, property, typeof(decimal));

        // Assert
        Assert.Equal(0.0, result);
    }

    #endregion

    #region Min Tests

    [Fact]
    public void MinStrategy_WithDecimalValues_FindsMinimum()
    {
        // Arrange
        var data = new List<Product>
        {
            new Product { Price = 30.0m },
            new Product { Price = 10.0m },
            new Product { Price = 20.0m }
        };
        var property = typeof(Product).GetProperty(nameof(Product.Price))!;
        var strategy = _factory.GetStrategy(AggregationType.Min);

        // Act
        var result = strategy.Calculate(data, property, typeof(decimal));

        // Assert
        Assert.Equal(10.0, result);
    }

    [Fact]
    public void MinStrategy_WithIntegerValues_FindsMinimum()
    {
        // Arrange
        var data = new List<Product>
        {
            new Product { Quantity = 50 },
            new Product { Quantity = 5 },
            new Product { Quantity = 25 }
        };
        var property = typeof(Product).GetProperty(nameof(Product.Quantity))!;
        var strategy = _factory.GetStrategy(AggregationType.Min);

        // Act
        var result = strategy.Calculate(data, property, typeof(int));

        // Assert
        Assert.Equal(5.0, result);
    }

    [Fact]
    public void MinStrategy_WithNegativeValues_FindsMinimum()
    {
        // Arrange
        var data = new List<Product>
        {
            new Product { Quantity = -10 },
            new Product { Quantity = 5 },
            new Product { Quantity = -20 }
        };
        var property = typeof(Product).GetProperty(nameof(Product.Quantity))!;
        var strategy = _factory.GetStrategy(AggregationType.Min);

        // Act
        var result = strategy.Calculate(data, property, typeof(int));

        // Assert
        Assert.Equal(-20.0, result);
    }

    #endregion

    #region Max Tests

    [Fact]
    public void MaxStrategy_WithDecimalValues_FindsMaximum()
    {
        // Arrange
        var data = new List<Product>
        {
            new Product { Price = 10.0m },
            new Product { Price = 50.0m },
            new Product { Price = 30.0m }
        };
        var property = typeof(Product).GetProperty(nameof(Product.Price))!;
        var strategy = _factory.GetStrategy(AggregationType.Max);

        // Act
        var result = strategy.Calculate(data, property, typeof(decimal));

        // Assert
        Assert.Equal(50.0, result);
    }

    [Fact]
    public void MaxStrategy_WithIntegerValues_FindsMaximum()
    {
        // Arrange
        var data = new List<Product>
        {
            new Product { Quantity = 25 },
            new Product { Quantity = 100 },
            new Product { Quantity = 50 }
        };
        var property = typeof(Product).GetProperty(nameof(Product.Quantity))!;
        var strategy = _factory.GetStrategy(AggregationType.Max);

        // Act
        var result = strategy.Calculate(data, property, typeof(int));

        // Assert
        Assert.Equal(100.0, result);
    }

    #endregion

    #region Count Tests

    [Fact]
    public void CountStrategy_WithValues_CountsAll()
    {
        // Arrange
        var data = new List<Product>
        {
            new Product { Price = 10.0m },
            new Product { Price = 20.0m },
            new Product { Price = 30.0m }
        };
        var property = typeof(Product).GetProperty(nameof(Product.Price))!;
        var strategy = _factory.GetStrategy(AggregationType.Count);

        // Act
        var result = strategy.Calculate(data, property, typeof(decimal));

        // Assert
        Assert.Equal(3.0, result);
    }

    [Fact]
    public void CountStrategy_WithNullableValues_CountsAllItems()
    {
        // Arrange
        var data = new List<Product>
        {
            new Product { NullablePrice = 10.0m },
            new Product { NullablePrice = null },
            new Product { NullablePrice = 20.0m }
        };
        var property = typeof(Product).GetProperty(nameof(Product.NullablePrice))!;
        var strategy = _factory.GetStrategy(AggregationType.Count);

        // Act
        var result = strategy.Calculate(data, property, typeof(decimal));

        // Assert
        // Count counts all items, not just non-null values
        Assert.Equal(3.0, result);
    }

    [Fact]
    public void CountStrategy_WithEmptyList_ReturnsZero()
    {
        // Arrange
        var data = new List<Product>();
        var property = typeof(Product).GetProperty(nameof(Product.Price))!;
        var strategy = _factory.GetStrategy(AggregationType.Count);

        // Act
        var result = strategy.Calculate(data, property, typeof(decimal));

        // Assert
        Assert.Equal(0.0, result);
    }

    #endregion

    #region All Numeric Types Tests

    [Theory]
    [InlineData(typeof(decimal))]
    [InlineData(typeof(double))]
    [InlineData(typeof(float))]
    [InlineData(typeof(int))]
    [InlineData(typeof(long))]
    [InlineData(typeof(short))]
    [InlineData(typeof(byte))]
    public void SumStrategy_WithAllNumericTypes_Works(Type numericType)
    {
        // Arrange
        var data = CreateSampleDataForType(numericType);
        var property = data[0].GetType().GetProperty("Value")!;
        var strategy = _factory.GetStrategy(AggregationType.Sum);

        // Act
        var result = strategy.Calculate(data, property, numericType);

        // Assert
        Assert.True(result > 0); // Should have some positive sum
    }

    #endregion

    private List<object> CreateSampleDataForType(Type type)
    {
        if (type == typeof(decimal))
            return new List<object> { new { Value = 10.5m }, new { Value = 20.5m } };
        if (type == typeof(double))
            return new List<object> { new { Value = 10.5 }, new { Value = 20.5 } };
        if (type == typeof(float))
            return new List<object> { new { Value = 10.5f }, new { Value = 20.5f } };
        if (type == typeof(int))
            return new List<object> { new { Value = 10 }, new { Value = 20 } };
        if (type == typeof(long))
            return new List<object> { new { Value = 10L }, new { Value = 20L } };
        if (type == typeof(short))
            return new List<object> { new { Value = (short)10 }, new { Value = (short)20 } };
        if (type == typeof(byte))
            return new List<object> { new { Value = (byte)10 }, new { Value = (byte)20 } };

        throw new ArgumentException($"Unsupported type: {type}");
    }

    // Test model
    private class Product
    {
        public decimal Price { get; set; }
        public int Quantity { get; set; }
        public double Weight { get; set; }
        public decimal? NullablePrice { get; set; }
    }
}
