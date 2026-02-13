using System.Reflection;
using System.Collections.Concurrent;
using System.Linq.Expressions;

namespace ExcelGenerator.Core.Aggregation;

/// <summary>
/// Generic aggregator that handles numeric calculations for all numeric types
/// Eliminates code duplication by using generics and delegates
/// </summary>
internal class NumericAggregator
{
    private static readonly ConcurrentDictionary<PropertyInfo, Func<object, object>> _propertyAccessorCache = new();

    /// <summary>
    /// Gets or creates a compiled property accessor for better performance
    /// </summary>
    private static Func<object, object> GetPropertyAccessor(PropertyInfo property)
    {
        return _propertyAccessorCache.GetOrAdd(property, prop =>
        {
            var parameter = Expression.Parameter(typeof(object), "obj");
            var cast = Expression.Convert(parameter, prop.DeclaringType!);
            var propertyAccess = Expression.Property(cast, prop);
            var convertToObject = Expression.Convert(propertyAccess, typeof(object));
            return Expression.Lambda<Func<object, object>>(convertToObject, parameter).Compile();
        });
    }

    /// <summary>
    /// Calculates sum for the specified numeric type
    /// </summary>
    public static double CalculateSum<T>(List<T> dataList, PropertyInfo property, Type underlyingType)
    {
        var accessor = GetPropertyAccessor(property);
        if (underlyingType == typeof(decimal))
        {
            var sum = dataList
                .Select(item => item == null ? 0m : (decimal)(accessor(item) ?? 0m))
                .Sum();
            return (double)sum.RefineValue();
        }
        else if (underlyingType == typeof(double))
        {
            var sum = dataList
                .Select(item => item == null ? 0.0 : (double)(accessor(item) ?? 0.0))
                .Sum();
            return (double)((decimal)sum).RefineValue();
        }
        else if (underlyingType == typeof(float))
        {
            var sum = dataList
                .Select(item => item == null ? 0f : (float)(accessor(item) ?? 0f))
                .Sum();
            return (double)((decimal)sum).RefineValue();
        }
        else if (underlyingType == typeof(int))
        {
            return dataList
                .Select(item => item == null ? 0 : (int)(accessor(item) ?? 0))
                .Sum();
        }
        else if (underlyingType == typeof(long))
        {
            return dataList
                .Select(item => item == null ? 0L : (long)(accessor(item) ?? 0L))
                .Sum();
        }
        else if (underlyingType == typeof(short))
        {
            return dataList
                .Select(item => item == null ? 0 : (int)(short)(accessor(item) ?? (short)0))
                .Sum();
        }
        else if (underlyingType == typeof(byte))
        {
            return dataList
                .Select(item => item == null ? 0 : (int)(byte)(accessor(item) ?? (byte)0))
                .Sum();
        }

        return 0;
    }

    /// <summary>
    /// Calculates minimum for the specified numeric type
    /// </summary>
    public static double CalculateMin<T>(List<T> dataList, PropertyInfo property, Type underlyingType)
    {
        var accessor = GetPropertyAccessor(property);

        if (underlyingType == typeof(decimal))
        {
            var min = dataList
                .Select(item => item == null ? decimal.MaxValue : (decimal)(accessor(item) ?? decimal.MaxValue))
                .Min();
            return (double)min.RefineValue();
        }
        else if (underlyingType == typeof(double))
        {
            var min = dataList
                .Select(item => item == null ? double.MaxValue : (double)(accessor(item) ?? double.MaxValue))
                .Min();
            return (double)((decimal)min).RefineValue();
        }
        else if (underlyingType == typeof(float))
        {
            var min = dataList
                .Select(item => item == null ? float.MaxValue : (float)(accessor(item) ?? float.MaxValue))
                .Min();
            return (double)((decimal)min).RefineValue();
        }
        else if (underlyingType == typeof(int))
        {
            return dataList
                .Select(item => item == null ? int.MaxValue : (int)(accessor(item) ?? int.MaxValue))
                .Min();
        }
        else if (underlyingType == typeof(long))
        {
            return dataList
                .Select(item => item == null ? long.MaxValue : (long)(accessor(item) ?? long.MaxValue))
                .Min();
        }
        else if (underlyingType == typeof(short))
        {
            return dataList
                .Select(item => item == null ? short.MaxValue : (int)(short)(accessor(item) ?? short.MaxValue))
                .Min();
        }
        else if (underlyingType == typeof(byte))
        {
            return dataList
                .Select(item => item == null ? byte.MaxValue : (int)(byte)(accessor(item) ?? byte.MaxValue))
                .Min();
        }

        return 0;
    }

    /// <summary>
    /// Calculates maximum for the specified numeric type
    /// </summary>
    public static double CalculateMax<T>(List<T> dataList, PropertyInfo property, Type underlyingType)
    {
        var accessor = GetPropertyAccessor(property);

        if (underlyingType == typeof(decimal))
        {
            var max = dataList
                .Select(item => item == null ? decimal.MinValue : (decimal)(accessor(item) ?? decimal.MinValue))
                .Max();
            return (double)max.RefineValue();
        }
        else if (underlyingType == typeof(double))
        {
            var max = dataList
                .Select(item => item == null ? double.MinValue : (double)(accessor(item) ?? double.MinValue))
                .Max();
            return (double)((decimal)max).RefineValue();
        }
        else if (underlyingType == typeof(float))
        {
            var max = dataList
                .Select(item => item == null ? float.MinValue : (float)(accessor(item) ?? float.MinValue))
                .Max();
            return (double)((decimal)max).RefineValue();
        }
        else if (underlyingType == typeof(int))
        {
            return dataList
                .Select(item => item == null ? int.MinValue : (int)(accessor(item) ?? int.MinValue))
                .Max();
        }
        else if (underlyingType == typeof(long))
        {
            return dataList
                .Select(item => item == null ? long.MinValue : (long)(accessor(item) ?? long.MinValue))
                .Max();
        }
        else if (underlyingType == typeof(short))
        {
            return dataList
                .Select(item => item == null ? short.MinValue : (int)(short)(accessor(item) ?? short.MinValue))
                .Max();
        }
        else if (underlyingType == typeof(byte))
        {
            return dataList
                .Select(item => item == null ? byte.MinValue : (int)(byte)(accessor(item) ?? byte.MinValue))
                .Max();
        }

        return 0;
    }

    /// <summary>
    /// Calculates average for the specified numeric type
    /// </summary>
    public static double CalculateAverage<T>(List<T> dataList, PropertyInfo property, Type underlyingType)
    {
        if (dataList.Count == 0) return 0;

        var sum = CalculateSum(dataList, property, underlyingType);
        var average = sum / dataList.Count;

        // Apply refinement for floating-point types
        if (underlyingType == typeof(decimal) || underlyingType == typeof(double) || underlyingType == typeof(float))
        {
            return (double)((decimal)average).RefineValue();
        }

        return average;
    }
}
