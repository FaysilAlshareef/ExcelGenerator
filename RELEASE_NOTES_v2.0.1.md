# v2.0.1 - Multi-Type Numeric Totals

## What's New in v2.0.1

### Features
- ✨ **All Numeric Types Totals**: Automatic summation row now supports ALL numeric types, not just decimal
  - Supported types: `decimal`, `double`, `float`, `int`, `long`, `short`, `byte`
- 🔧 **RefineValue Extension**: New public extension method for precise decimal calculations
  - Truncates to 3 decimal places instead of rounding
  - Available for use in your own code via `NumericExtensions.RefineValue()`

### Improvements
- Enhanced number formatting consistency across all numeric types
- Floating-point numbers display with 2 decimal places
- Integer types display without decimals

### Technical Details
- Summation logic now handles type-specific conversions
- Decimal values refined using truncation for precision
- Maintains backward compatibility - no breaking changes

## Installation

```bash
dotnet add package Faysil.ExcelGenerator --version 2.0.1
```

## Full Changelog
- Add `NumericExtensions` class with `RefineValue` extension method
- Extend `AddSummationRow` to calculate totals for all numeric types
- Update package version to 2.0.1
- Update documentation to reflect new capabilities
