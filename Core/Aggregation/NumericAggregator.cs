using System.Reflection;

namespace ExcelGenerator.Core.Aggregation;

/// <summary>
/// Generic aggregator that handles numeric calculations for all numeric types
/// Eliminates code duplication by using generics and delegates
/// </summary>
internal class NumericAggregator
{
    /// <summary>
    /// Calculates sum for the specified numeric type
    /// </summary>
    public static double CalculateSum<T>(List<T> dataList, PropertyInfo property, Type underlyingType)
    {
        if (underlyingType == typeof(decimal))
        {
            var sum = dataList
                .Select(item => item == null ? 0m : (decimal)(property.GetValue(item) ?? 0m))
                .Sum();
            return (double)sum.RefineValue();
        }
        else if (underlyingType == typeof(double))
        {
            var sum = dataList
                .Select(item => item == null ? 0.0 : (double)(property.GetValue(item) ?? 0.0))
                .Sum();
            return (double)((decimal)sum).RefineValue();
        }
        else if (underlyingType == typeof(float))
        {
            var sum = dataList
                .Select(item => item == null ? 0f : (float)(property.GetValue(item) ?? 0f))
                .Sum();
            return (double)((decimal)sum).RefineValue();
        }
        else if (underlyingType == typeof(int))
        {
            return dataList
                .Select(item => item == null ? 0 : (int)(property.GetValue(item) ?? 0))
                .Sum();
        }
        else if (underlyingType == typeof(long))
        {
            return dataList
                .Select(item => item == null ? 0L : (long)(property.GetValue(item) ?? 0L))
                .Sum();
        }
        else if (underlyingType == typeof(short))
        {
            return dataList
                .Select(item => item == null ? 0 : (int)(short)(property.GetValue(item) ?? (short)0))
                .Sum();
        }
        else if (underlyingType == typeof(byte))
        {
            return dataList
                .Select(item => item == null ? 0 : (int)(byte)(property.GetValue(item) ?? (byte)0))
                .Sum();
        }

        return 0;
    }

    /// <summary>
    /// Calculates minimum for the specified numeric type
    /// </summary>
    public static double CalculateMin<T>(List<T> dataList, PropertyInfo property, Type underlyingType)
    {
        if (underlyingType == typeof(decimal))
        {
            var min = dataList
                .Select(item => item == null ? decimal.MaxValue : (decimal)(property.GetValue(item) ?? decimal.MaxValue))
                .Min();
            return (double)min.RefineValue();
        }
        else if (underlyingType == typeof(double))
        {
            var min = dataList
                .Select(item => item == null ? double.MaxValue : (double)(property.GetValue(item) ?? double.MaxValue))
                .Min();
            return (double)((decimal)min).RefineValue();
        }
        else if (underlyingType == typeof(float))
        {
            var min = dataList
                .Select(item => item == null ? float.MaxValue : (float)(property.GetValue(item) ?? float.MaxValue))
                .Min();
            return (double)((decimal)min).RefineValue();
        }
        else if (underlyingType == typeof(int))
        {
            return dataList
                .Select(item => item == null ? int.MaxValue : (int)(property.GetValue(item) ?? int.MaxValue))
                .Min();
        }
        else if (underlyingType == typeof(long))
        {
            return dataList
                .Select(item => item == null ? long.MaxValue : (long)(property.GetValue(item) ?? long.MaxValue))
                .Min();
        }
        else if (underlyingType == typeof(short))
        {
            return dataList
                .Select(item => item == null ? short.MaxValue : (int)(short)(property.GetValue(item) ?? short.MaxValue))
                .Min();
        }
        else if (underlyingType == typeof(byte))
        {
            return dataList
                .Select(item => item == null ? byte.MaxValue : (int)(byte)(property.GetValue(item) ?? byte.MaxValue))
                .Min();
        }

        return 0;
    }

    /// <summary>
    /// Calculates maximum for the specified numeric type
    /// </summary>
    public static double CalculateMax<T>(List<T> dataList, PropertyInfo property, Type underlyingType)
    {
        if (underlyingType == typeof(decimal))
        {
            var max = dataList
                .Select(item => item == null ? decimal.MinValue : (decimal)(property.GetValue(item) ?? decimal.MinValue))
                .Max();
            return (double)max.RefineValue();
        }
        else if (underlyingType == typeof(double))
        {
            var max = dataList
                .Select(item => item == null ? double.MinValue : (double)(property.GetValue(item) ?? double.MinValue))
                .Max();
            return (double)((decimal)max).RefineValue();
        }
        else if (underlyingType == typeof(float))
        {
            var max = dataList
                .Select(item => item == null ? float.MinValue : (float)(property.GetValue(item) ?? float.MinValue))
                .Max();
            return (double)((decimal)max).RefineValue();
        }
        else if (underlyingType == typeof(int))
        {
            return dataList
                .Select(item => item == null ? int.MinValue : (int)(property.GetValue(item) ?? int.MinValue))
                .Max();
        }
        else if (underlyingType == typeof(long))
        {
            return dataList
                .Select(item => item == null ? long.MinValue : (long)(property.GetValue(item) ?? long.MinValue))
                .Max();
        }
        else if (underlyingType == typeof(short))
        {
            return dataList
                .Select(item => item == null ? short.MinValue : (int)(short)(property.GetValue(item) ?? short.MinValue))
                .Max();
        }
        else if (underlyingType == typeof(byte))
        {
            return dataList
                .Select(item => item == null ? byte.MinValue : (int)(byte)(property.GetValue(item) ?? byte.MinValue))
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
