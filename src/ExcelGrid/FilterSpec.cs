using System;
using System.Collections.Generic;
using System.Globalization;

namespace ExcelGrid.Ssms;

internal enum TextFilterOperator
{
    None,
    Contains,
    DoesNotContain,
    Equals,
    DoesNotEqual,
    StartsWith,
    EndsWith,
    IsBlank,
    IsNotBlank
}
internal sealed class FilterSpec
{
    public HashSet<string>? AllowedValues { get; set; }
    public TextFilterOperator Operator { get; set; }
    public string Operand { get; set; } = string.Empty;

    public bool IsActive => AllowedValues != null || Operator != TextFilterOperator.None;

    public bool Matches(string? value)
    {
        value ??= string.Empty;
        if (AllowedValues != null && !AllowedValues.Contains(value))
            return false;

        var comparison = CultureInfo.CurrentCulture.CompareInfo;
        var options = CompareOptions.IgnoreCase | CompareOptions.IgnoreWidth;
        return Operator switch
        {
            TextFilterOperator.None => true,
            TextFilterOperator.Contains => comparison.IndexOf(value, Operand, options) >= 0,
            TextFilterOperator.DoesNotContain => comparison.IndexOf(value, Operand, options) < 0,
            TextFilterOperator.Equals => comparison.Compare(value, Operand, options) == 0,
            TextFilterOperator.DoesNotEqual => comparison.Compare(value, Operand, options) != 0,
            TextFilterOperator.StartsWith => comparison.IsPrefix(value, Operand, options),
            TextFilterOperator.EndsWith => comparison.IsSuffix(value, Operand, options),
            TextFilterOperator.IsBlank => string.IsNullOrWhiteSpace(value) || value == "NULL",
            TextFilterOperator.IsNotBlank => !string.IsNullOrWhiteSpace(value) && value != "NULL",
            _ => true
        };
    }
}
