using System.Globalization;
using System.Text;

namespace VbaNg.Runtime;

/// <summary>
/// Number text the way VBA writes and reads it (MS-VBAL 5.5.1.2.4 Let-coercion to and from
/// String). Formatting keeps 15 significant digits for Double and 7 for Single, and uses fixed
/// notation whenever the integral part fits in that many digits and the fraction needs no more,
/// otherwise scientific notation with a signed two-digit exponent; both recorded from Excel
/// (Conversion goldens). Parsing accepts what Excel accepts: surrounding whitespace, a leading
/// sign, a trailing minus, parentheses, the currency symbol, digit grouping anywhere before the
/// decimal separator, an E or D exponent, and the &amp;H and &amp;O prefixes.
/// </summary>
public static class NumberText
{
    public const int DoubleDigits = 15;
    public const int SingleDigits = 7;

    /// <summary>Formats with "." as the decimal separator, as <c>Str</c> does.</summary>
    public static string FormatDouble(double value, int significantDigits = DoubleDigits) => Format(value, significantDigits, ".");

    public static string FormatSingle(float value) => Format(value, SingleDigits, ".");

    /// <summary>Formats with the given decimal separator, as <c>CStr</c> does under the current regional settings.</summary>
    public static string Format(double value, int significantDigits, string decimalSeparator)
    {
        ArgumentNullException.ThrowIfNull(decimalSeparator);
        if (double.IsNaN(value))
        {
            return "-1.#IND";
        }

        if (double.IsPositiveInfinity(value))
        {
            return "1.#INF";
        }

        if (double.IsNegativeInfinity(value))
        {
            return "-1.#INF";
        }

        if (value == 0)
        {
            return double.IsNegative(value) ? "-0" : "0";
        }

        // d.ddddE+xxx with the requested significant digits, correctly rounded by the runtime.
        var scientific = Math.Abs(value).ToString("E" + (significantDigits - 1).ToString(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
        var e = scientific.IndexOf('E', StringComparison.Ordinal);
        var digits = scientific[..e].Replace(".", string.Empty, StringComparison.Ordinal).TrimEnd('0');
        var exponent = int.Parse(scientific.AsSpan(e + 1), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture);
        if (digits.Length == 0)
        {
            digits = "0";
        }

        var text = new StringBuilder();
        if (value < 0)
        {
            text.Append('-');
        }

        var fractionDigitsNeeded = digits.Length - 1 - exponent;
        var integralDigits = exponent + 1;
        if (integralDigits > significantDigits || fractionDigitsNeeded > significantDigits)
        {
            text.Append(digits[0]);
            if (digits.Length > 1)
            {
                text.Append(decimalSeparator).Append(digits, 1, digits.Length - 1);
            }

            text.Append('E').Append(exponent < 0 ? '-' : '+').Append(Math.Abs(exponent).ToString("00", CultureInfo.InvariantCulture));
            return text.ToString();
        }

        if (exponent < 0)
        {
            text.Append('0').Append(decimalSeparator).Append('0', -exponent - 1).Append(digits);
        }
        else if (digits.Length <= integralDigits)
        {
            text.Append(digits).Append('0', integralDigits - digits.Length);
        }
        else
        {
            text.Append(digits, 0, integralDigits).Append(decimalSeparator).Append(digits, integralDigits, digits.Length - integralDigits);
        }

        return text.ToString();
    }

    /// <summary>
    /// Parses a numeric-coercion-string under the given culture's regional settings. Returns false
    /// when the text is not a number (the caller raises error 13). Hex and octal prefixes yield
    /// unsigned integers, unlike literals: "&amp;HFFFF" is 65535 here (Conversion goldens).
    /// </summary>
    public static bool TryParse(string text, CultureInfo culture, out decimal integral, out double floating, out bool isIntegral)
    {
        ArgumentNullException.ThrowIfNull(text);
        return TryParse(text.AsSpan(), culture, out integral, out floating, out isIntegral);
    }

    /// <summary>The same over a span, so a String's BSTR is read in place (ARCHITECTURE.md D20).</summary>
    public static bool TryParse(ReadOnlySpan<char> text, CultureInfo culture, out decimal integral, out double floating, out bool isIntegral)
    {
        ArgumentNullException.ThrowIfNull(culture);
        integral = 0;
        floating = 0;
        isIntegral = false;

        var span = text.Trim(Whitespace);
        if (span.Length == 0)
        {
            return false;
        }

        if (span.Length > 2 && span[0] == '&' && (span[1] is 'H' or 'h' or 'O' or 'o'))
        {
            if (!TryParseRadixLiteral(span, out var value))
            {
                return false;
            }

            integral = value;
            floating = value;
            isIntegral = true;
            return true;
        }

        var negative = false;
        var format = culture.NumberFormat;

        // Parentheses mean negative, as in accounting formats.
        if (span.Length >= 2 && span[0] == '(' && span[^1] == ')')
        {
            negative = true;
            span = span[1..^1].Trim(Whitespace);
        }

        if (span.Length > 0 && span[0] is '+' or '-')
        {
            negative ^= span[0] == '-';
            span = span[1..].TrimStart(Whitespace);
        }
        else if (span.Length > 0 && span[^1] == '-')
        {
            negative = true;
            span = span[..^1].TrimEnd(Whitespace);
        }

        var currency = format.CurrencySymbol;
        if (currency.Length > 0 && span.StartsWith(currency, StringComparison.Ordinal))
        {
            span = span[currency.Length..].TrimStart(Whitespace);
        }

        if (span.Length > 0 && span[0] is '+' or '-')
        {
            negative ^= span[0] == '-';
            span = span[1..].TrimStart(Whitespace);
        }

        if (span.Length == 0)
        {
            return false;
        }

        var digits = new StringBuilder();
        var sawDigit = false;
        var sawDecimal = false;
        var scale = 0;
        var i = 0;
        var decimalSeparator = format.NumberDecimalSeparator;
        var groupSeparator = format.NumberGroupSeparator;
        while (i < span.Length)
        {
            var c = span[i];
            if (c is >= '0' and <= '9')
            {
                digits.Append(c);
                sawDigit = true;
                if (sawDecimal)
                {
                    scale++;
                }

                i++;
            }
            else if (!sawDecimal && span[i..].StartsWith(decimalSeparator, StringComparison.Ordinal))
            {
                sawDecimal = true;
                i += decimalSeparator.Length;
            }
            else if (!sawDecimal && groupSeparator.Length > 0 && span[i..].StartsWith(groupSeparator, StringComparison.Ordinal))
            {
                i += groupSeparator.Length;
            }
            else
            {
                break;
            }
        }

        if (!sawDigit)
        {
            return false;
        }

        var exponent = 0;
        if (i < span.Length && span[i] is 'e' or 'E' or 'd' or 'D')
        {
            i++;
            var rest = span[i..].TrimStart(Whitespace);
            var exponentNegative = false;
            if (rest.Length > 0 && rest[0] is '+' or '-')
            {
                exponentNegative = rest[0] == '-';
                rest = rest[1..].TrimStart(Whitespace);
            }

            var start = 0;
            while (start < rest.Length && rest[start] is >= '0' and <= '9')
            {
                start++;
            }

            if (start == 0 || start > 9)
            {
                return false;
            }

            exponent = int.Parse(rest[..start], NumberStyles.None, CultureInfo.InvariantCulture);
            if (exponentNegative)
            {
                exponent = -exponent;
            }

            rest = rest[start..].TrimStart(Whitespace);
            if (rest.Length != 0)
            {
                return false;
            }
        }
        else if (span[i..].TrimStart(Whitespace).Length != 0)
        {
            return false;
        }

        // digits × 10^(exponent - scale), sign applied: build an invariant literal the runtime parses exactly.
        var literal = new StringBuilder(digits.Length + 16);
        if (negative)
        {
            literal.Append('-');
        }

        literal.Append(digits).Append('E').Append((exponent - scale).ToString(CultureInfo.InvariantCulture));
        var invariant = literal.ToString();
        floating = double.Parse(invariant, NumberStyles.Float, CultureInfo.InvariantCulture);
        if (TryBuildDecimal(digits, exponent - scale, negative, out integral))
        {
            // The Decimal keeps the text's scale and sign, as VBA's does: "-0.0" is a negative zero with one place (Conversion golden).
            isIntegral = true;
        }
        else if (decimal.TryParse(invariant, NumberStyles.Float, CultureInfo.InvariantCulture, out integral))
        {
            isIntegral = true;
        }
        else if (double.IsInfinity(floating))
        {
            // Beyond Double: the caller reports overflow.
            isIntegral = false;
        }

        return true;
    }

    private static ReadOnlySpan<char> Whitespace => [' ', '\t', '\n', '\r'];

    /// <summary>
    /// digits × 10^shift as a Decimal built from its parts when the magnitude fits 96 bits and
    /// the scale 28 places, so trailing zeros and a negative zero keep their scale and sign
    /// (what VarDecFromStr does); false leaves the rounding parse to the caller.
    /// </summary>
    private static bool TryBuildDecimal(StringBuilder digits, int shift, bool negative, out decimal value)
    {
        value = 0;
        if (shift < -28 || digits.Length > 40)
        {
            return false;
        }

        var magnitude = System.Numerics.BigInteger.Parse(digits.ToString(), CultureInfo.InvariantCulture);
        if (shift > 0)
        {
            magnitude *= System.Numerics.BigInteger.Pow(10, shift);
        }

        if (magnitude >= System.Numerics.BigInteger.One << 96)
        {
            return false;
        }

        var low = unchecked((int)(uint)(magnitude & uint.MaxValue));
        var mid = unchecked((int)(uint)((magnitude >> 32) & uint.MaxValue));
        var high = unchecked((int)(uint)((magnitude >> 64) & uint.MaxValue));
        value = new decimal(low, mid, high, negative, (byte)Math.Max(0, -shift));
        return true;
    }

    /// <summary>&amp;H and &amp;O in a string read as an unsigned magnitude ("&amp;HFFFF" is 65535), and a type suffix is not accepted (Conversion goldens).</summary>
    private static bool TryParseRadixLiteral(ReadOnlySpan<char> span, out long value)
    {
        value = 0;
        var radix = span[1] is 'H' or 'h' ? 16 : 8;
        var body = span[2..];
        if (body.Length == 0)
        {
            return false;
        }

        ulong magnitude = 0;
        foreach (var c in body)
        {
            var digit = c switch
            {
                >= '0' and <= '9' => c - '0',
                >= 'A' and <= 'F' => c - 'A' + 10,
                >= 'a' and <= 'f' => c - 'a' + 10,
                _ => -1,
            };
            if (digit < 0 || digit >= radix)
            {
                return false;
            }

            magnitude = checked(magnitude * (ulong)radix + (ulong)digit);
            if (magnitude > long.MaxValue)
            {
                return false;
            }
        }

        value = (long)magnitude;
        return true;
    }
}
