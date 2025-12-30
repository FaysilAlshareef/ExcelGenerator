using System.Reflection;
using System.Text.RegularExpressions;

namespace ExcelGenerator.Core.PropertyReflection;

/// <summary>
/// Service for extracting and filtering properties from types
/// </summary>
internal class PropertyExtractor : IPropertyExtractor
{
    public PropertyInfo[] Extract<T>(bool excludeIds = false)
    {
        var properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(p => p.CanRead);

        if (excludeIds)
        {
            properties = properties.Where(p =>
                !p.Name.EndsWith("Id", StringComparison.OrdinalIgnoreCase) &&
                !p.Name.EndsWith("ID", StringComparison.Ordinal));
        }

        return properties.ToArray();
    }

    public string FormatPropertyName(string propertyName)
    {
        // Insert spaces before capital letters (for PascalCase properties)
        var formatted = Regex.Replace(
            propertyName,
            "([a-z])([A-Z])",
            "$1 $2");

        return formatted;
    }
}
