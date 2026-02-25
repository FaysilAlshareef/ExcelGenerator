using System.Collections.Concurrent;
using System.Linq.Expressions;
using System.Reflection;

namespace ExcelGenerator.Core.PropertyReflection;

/// <summary>
/// Cache for compiled property accessors using Expression Trees
/// Provides 10-100x faster property access compared to reflection
/// </summary>
internal static class PropertyAccessorCache<T>
{
    private static readonly ConcurrentDictionary<PropertyInfo, Func<T, object?>> _getters = new();

    /// <summary>
    /// Gets or creates a compiled accessor for the specified property
    /// </summary>
    public static Func<T, object?> GetAccessor(PropertyInfo property)
    {
        return _getters.GetOrAdd(property, CompileAccessor);
    }

    private static Func<T, object?> CompileAccessor(PropertyInfo property)
    {
        // Create parameter: (T instance)
        var instance = Expression.Parameter(typeof(T), "instance");

        // Create property access: instance.PropertyName
        var propertyAccess = Expression.Property(instance, property);

        // Convert to object: (object)instance.PropertyName
        var castToObject = Expression.Convert(propertyAccess, typeof(object));

        // Compile to delegate: (T instance) => (object)instance.PropertyName
        return Expression.Lambda<Func<T, object?>>(castToObject, instance).Compile();
    }

    /// <summary>
    /// Clears the cache (useful for testing or memory management)
    /// </summary>
    public static void Clear()
    {
        _getters.Clear();
    }
}
