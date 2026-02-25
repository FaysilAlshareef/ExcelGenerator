using System.Reflection;

namespace ExcelGenerator.Core.PropertyReflection;

/// <summary>
/// Cached metadata about a property to avoid repeated reflection and type checking
/// </summary>
internal class PropertyMetadata
{
    public PropertyInfo Property { get; }
    public string Name { get; }
    public Type PropertyType { get; }
    public Type UnderlyingType { get; }
    public bool IsNumeric { get; }
    public bool IsFloatingPoint { get; }

    public PropertyMetadata(PropertyInfo property)
    {
        Property = property;
        Name = property.Name;
        PropertyType = property.PropertyType;
        UnderlyingType = Nullable.GetUnderlyingType(PropertyType) ?? PropertyType;
        IsNumeric = CheckIsNumeric(UnderlyingType);
        IsFloatingPoint = CheckIsFloatingPoint(UnderlyingType);
    }

    private static bool CheckIsNumeric(Type type)
    {
        return type == typeof(decimal) || type == typeof(double) || type == typeof(float) ||
               type == typeof(int) || type == typeof(long) || type == typeof(short) || type == typeof(byte);
    }

    private static bool CheckIsFloatingPoint(Type type)
    {
        return type == typeof(decimal) || type == typeof(double) || type == typeof(float);
    }
}
