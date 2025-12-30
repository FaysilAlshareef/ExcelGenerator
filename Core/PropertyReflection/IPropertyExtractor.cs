using System.Reflection;

namespace ExcelGenerator.Core.PropertyReflection;

/// <summary>
/// Defines a service for extracting and filtering properties from types
/// </summary>
internal interface IPropertyExtractor
{
    /// <summary>
    /// Extracts readable properties from the specified type
    /// </summary>
    /// <typeparam name="T">The type to extract properties from</typeparam>
    /// <param name="excludeIds">If true, excludes properties that end with "Id"</param>
    /// <returns>An array of properties that meet the criteria</returns>
    PropertyInfo[] Extract<T>(bool excludeIds = false);

    /// <summary>
    /// Formats a property name for display (e.g., converts PascalCase to Pascal Case)
    /// </summary>
    /// <param name="propertyName">The property name to format</param>
    /// <returns>The formatted property name</returns>
    string FormatPropertyName(string propertyName);
}
