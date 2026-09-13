using System.Globalization;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;

namespace VbaNg.Runtime.Library;

/// <summary>
/// The Strings module of the VBA standard library (MS-VBAL 6.1.2, Strings module). Functions
/// that VBA declares as returning Variant return Variant here and pass Null through; the $
/// forms are the same functions with a String result. Argument checks and error numbers follow
/// the Strings goldens. The functions read their arguments as spans over the BSTR (D20), so a
/// byte string of odd length is exact and no .NET string is made on the way; a result is a new
/// BSTR the statement owns.
/// </summary>
public static partial class Strings
{
    /// <summary>The system ANSI code page, as VBA uses it for byte strings and for the Declare parameters.</summary>
    internal static readonly Encoding Ansi = CreateAnsi();

    /// <summary>Len: the length of a string, or of the string form of a Variant value; Null gives Null.</summary>
    public static Variant Len(in Variant value)
    {
        if (value.IsNull)
        {
            return Variant.Null;
        }

        return Variant.FromInt32(Coerce.ToText(value).Length);
    }

    /// <summary>Len of a declared non-String variable: its byte size, as an Integer (Strings goldens).</summary>
    public static Variant LenOfDeclared(VarType type) => Variant.FromInt16((short)LenOfType(type));

    /// <summary>The byte size of a declared non-String variable, which Len reports for typed arguments.</summary>
    public static int LenOfType(VarType type) => type switch
    {
        VarType.Byte => 1,
        VarType.Integer or VarType.Boolean => 2,
        VarType.Long or VarType.Single => 4,
        VarType.Double or VarType.Currency or VarType.Date or VarType.LongLong => 8,
        _ => 16,
    };

    /// <summary>LenB: every byte of the BSTR, an odd one included (Strings golden).</summary>
    public static Variant LenB(in Variant value) => value.IsNull ? Variant.Null : Variant.FromInt32(Coerce.ToText(value).ByteLength);

    public static Variant Left(in Variant text, in Variant length)
    {
        var count = Coerce.ToInt32(length);
        if (count < 0)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        if (text.IsNull)
        {
            return Variant.Null;
        }

        var value = Coerce.ToText(text).Chars;
        return Result(count >= value.Length ? value : value[..count]);
    }

    public static Variant Right(in Variant text, in Variant length)
    {
        var count = Coerce.ToInt32(length);
        if (count < 0)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        if (text.IsNull)
        {
            return Variant.Null;
        }

        var value = Coerce.ToText(text).Chars;
        return Result(count >= value.Length ? value : value[^count..]);
    }

    /// <summary>Mid(String, Start, [Length]): Start below 1 or a negative Length raises 5; a start past the end gives "".</summary>
    public static Variant Mid(in Variant text, in Variant start, in Variant length)
    {
        var first = Coerce.ToInt32(start);
        if (first < 1)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        var count = int.MaxValue;
        if (!length.IsMissing)
        {
            count = Coerce.ToInt32(length);
            if (count < 0)
            {
                throw VbaErrors.InvalidProcedureCall();
            }
        }

        if (text.IsNull)
        {
            return Variant.Null;
        }

        var value = Coerce.ToText(text).Chars;
        if (first > value.Length)
        {
            return Variant.EmptyString;
        }

        var available = value.Length - first + 1;
        return Result(value.Slice(first - 1, Math.Min(count, available)));
    }

    /// <summary>LeftB(String, Length): the first Length bytes, an odd count included (Strings golden: LeftB("hello", 3) is three bytes that read as "h").</summary>
    public static Variant LeftB(in Variant text, in Variant length)
    {
        var count = Coerce.ToInt32(length);
        if (count < 0)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        if (text.IsNull)
        {
            return Variant.Null;
        }

        var bytes = Coerce.ToText(text).Bytes;
        return ResultBytes(count >= bytes.Length ? bytes : bytes[..count]);
    }

    /// <summary>RightB(String, Length): the last Length bytes.</summary>
    public static Variant RightB(in Variant text, in Variant length)
    {
        var count = Coerce.ToInt32(length);
        if (count < 0)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        if (text.IsNull)
        {
            return Variant.Null;
        }

        var bytes = Coerce.ToText(text).Bytes;
        return ResultBytes(count >= bytes.Length ? bytes : bytes[^count..]);
    }

    /// <summary>MidB(String, Start, [Length]): bytes from the 1-based byte position Start.</summary>
    public static Variant MidB(in Variant text, in Variant start, in Variant length)
    {
        var first = Coerce.ToInt32(start);
        if (first < 1)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        var count = int.MaxValue;
        if (!length.IsMissing)
        {
            count = Coerce.ToInt32(length);
            if (count < 0)
            {
                throw VbaErrors.InvalidProcedureCall();
            }
        }

        if (text.IsNull)
        {
            return Variant.Null;
        }

        var bytes = Coerce.ToText(text).Bytes;
        if (first > bytes.Length)
        {
            return Variant.EmptyString;
        }

        var available = bytes.Length - first + 1;
        return ResultBytes(bytes.Slice(first - 1, Math.Min(count, available)));
    }

    /// <summary>InStr([Start,] String1, String2, [Compare]): 1-based position or 0; Null strings give Null; Start below 1 raises 5.</summary>
    public static Variant InStr(in Variant start, in Variant text, in Variant find, in Variant compare, CompareMode moduleMode)
    {
        var first = start.IsMissing ? 1 : Coerce.ToInt32(start);
        if (first < 1)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        var mode = CompareModeArgument(compare, moduleMode);
        if (text.IsNull || find.IsNull)
        {
            return Variant.Null;
        }

        var haystack = Coerce.ToText(text).Chars;
        var needle = Coerce.ToText(find).Chars;
        if (haystack.Length == 0)
        {
            return Variant.FromInt32(0);
        }

        // An empty search string is found at Start, even past the end (Strings golden).
        if (needle.Length == 0)
        {
            return Variant.FromInt32(first);
        }

        if (first > haystack.Length)
        {
            return Variant.FromInt32(0);
        }

        var index = IndexOf(haystack, needle, first - 1, mode);
        return Variant.FromInt32(index < 0 ? 0 : index + 1);
    }

    /// <summary>InStr with its arguments as written: with three or more, a leading number is Start; otherwise the first two are the strings (Strings golden).</summary>
    public static Variant InStr(ReadOnlySpan<Variant> arguments, CompareMode moduleMode)
    {
        if (arguments.Length >= 3 && !arguments[0].IsString && (arguments[0].IsNumeric || arguments[0].IsEmpty))
        {
            return InStr(arguments[0], arguments[1], arguments[2], arguments.Length > 3 ? arguments[3] : Variant.Missing, moduleMode);
        }

        return InStr(Variant.Missing, arguments[0], arguments[1], arguments.Length > 2 ? arguments[2] : Variant.Missing, moduleMode);
    }

    /// <summary>InStrB([Start,] String1, String2, [Compare]): the 1-based byte position, over the bytes under binary comparison and over the characters under text comparison.</summary>
    public static Variant InStrB(in Variant start, in Variant text, in Variant find, in Variant compare, CompareMode moduleMode)
    {
        var first = start.IsMissing ? 1 : Coerce.ToInt32(start);
        if (first < 1)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        var mode = CompareModeArgument(compare, moduleMode);
        if (text.IsNull || find.IsNull)
        {
            return Variant.Null;
        }

        var haystack = Coerce.ToText(text);
        var needle = Coerce.ToText(find);
        if (haystack.ByteLength == 0)
        {
            return Variant.FromInt32(0);
        }

        if (needle.ByteLength == 0)
        {
            return Variant.FromInt32(first);
        }

        if (first > haystack.ByteLength)
        {
            return Variant.FromInt32(0);
        }

        if (mode == CompareMode.Binary)
        {
            var index = haystack.Bytes[(first - 1)..].IndexOf(needle.Bytes);
            return Variant.FromInt32(index < 0 ? 0 : index + first);
        }

        var charIndex = IndexOf(haystack.Chars, needle.Chars, (first - 1) / 2, mode);
        return Variant.FromInt32(charIndex < 0 ? 0 : charIndex * 2 + 1);
    }

    public static Variant InStrB(ReadOnlySpan<Variant> arguments, CompareMode moduleMode)
    {
        if (arguments.Length >= 3 && !arguments[0].IsString && (arguments[0].IsNumeric || arguments[0].IsEmpty))
        {
            return InStrB(arguments[0], arguments[1], arguments[2], arguments.Length > 3 ? arguments[3] : Variant.Missing, moduleMode);
        }

        return InStrB(Variant.Missing, arguments[0], arguments[1], arguments.Length > 2 ? arguments[2] : Variant.Missing, moduleMode);
    }

    /// <summary>The $ form of a function (Left$, Error$): a String result, where Null raises error 94 (MS-VBAL 6.1.2).</summary>
    public static VbaString DollarForm(in Variant value) => value.IsNull ? throw VbaErrors.InvalidUseOfNull() : Coerce.ToText(value);

    /// <summary>Mid statement (MS-VBAL 5.4.3.5) on a String variable or a Variant holding one: writes the code units in place (D20), never changing the length.</summary>
    public static void MidAssign(VbaString target, int start, int? length, VbaString value)
    {
        var chars = Bstr.MutableChars(target.Pointer);
        if (start < 1 || start > chars.Length)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        var source = value.Chars;
        var count = Math.Min(source.Length, chars.Length - start + 1);
        if (length is { } limit)
        {
            if (limit < 0)
            {
                throw VbaErrors.InvalidProcedureCall();
            }

            count = Math.Min(count, limit);
        }

        source[..count].CopyTo(chars[(start - 1)..]);
    }

    /// <summary>Mid statement on a fixed-length string, which is still a .NET string until M7 lands its inline storage.</summary>
    public static string MidAssign(string target, int start, int? length, VbaString value)
    {
        ArgumentNullException.ThrowIfNull(target);
        if (start < 1 || start > target.Length)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        var count = Math.Min(value.Length, target.Length - start + 1);
        if (length is { } limit)
        {
            if (limit < 0)
            {
                throw VbaErrors.InvalidProcedureCall();
            }

            count = Math.Min(count, limit);
        }

        return string.Concat(target.AsSpan(0, start - 1), value.Chars[..count], target.AsSpan(start - 1 + count));
    }

    /// <summary>Fits a value into a fixed-length string (MS-VBAL 5.5.1.2.4): truncated or padded with spaces on the right.</summary>
    public static string ToFixed(string value, int length)
    {
        ArgumentNullException.ThrowIfNull(value);
        return value.Length >= length ? value[..length] : value.PadRight(length);
    }

    public static string ToFixed(VbaString value, int length) => ToFixed(value.ToString(), length);

    /// <summary>A fixed-length string member of a record, which lies inline in it as its characters (ROADMAP.md M7 C5), read as the String * n value.</summary>
    public static string FromChars<TChars>(in TChars chars)
        where TChars : struct =>
        new(MemoryMarshal.CreateReadOnlySpan(ref Unsafe.As<TChars, char>(ref Unsafe.AsRef(in chars)), Unsafe.SizeOf<TChars>() / sizeof(char)));

    /// <summary>A store into a fixed-length string member of a record: the value, already padded or truncated to the member's length, over its characters.</summary>
    public static void SetChars<TChars>(ref TChars chars, string value)
        where TChars : struct
    {
        ArgumentNullException.ThrowIfNull(value);
        var target = MemoryMarshal.CreateSpan(ref Unsafe.As<TChars, char>(ref chars), Unsafe.SizeOf<TChars>() / sizeof(char));
        value.AsSpan(0, Math.Min(value.Length, target.Length)).CopyTo(target);
    }

    /// <summary>LSet on a String variable: a new string of the target's current length (D20).</summary>
    public static VbaString LSet(VbaString target, VbaString value) => VbaString.Temporary(ToFixed(value, target.Length));

    /// <summary>RSet on a String variable.</summary>
    public static VbaString RSet(VbaString target, VbaString value) => VbaString.Temporary(RSet(target.ToString(), value));

    /// <summary>LSet target = value: the value left-justified in the target's current length, padded with spaces or cut on the right; an empty target stays empty (MS-VBAL 5.4.3.6, Types golden).</summary>
    public static string LSet(string target, VbaString value)
    {
        ArgumentNullException.ThrowIfNull(target);
        return ToFixed(value, target.Length);
    }

    /// <summary>RSet target = value: the value right-justified in the target's current length, padded with spaces on the left; a longer value keeps its leftmost characters (MS-VBAL 5.4.3.7, Types golden).</summary>
    public static string RSet(string target, VbaString value)
    {
        ArgumentNullException.ThrowIfNull(target);
        var text = value.ToString();
        return text.Length >= target.Length ? text[..target.Length] : text.PadLeft(target.Length);
    }

    /// <summary>InStrRev(StringCheck, StringMatch, [Start = -1], [Compare]).</summary>
    public static Variant InStrRev(in Variant text, in Variant find, in Variant start, in Variant compare, CompareMode moduleMode)
    {
        var last = start.IsMissing ? -1 : Coerce.ToInt32(start);
        if (last == 0 || last < -1)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        var mode = CompareModeArgument(compare, moduleMode);

        // Unlike InStr, the strings are declared String, so Null raises 94 (Strings golden).
        var haystack = Coerce.ToText(DollarText(text)).Chars;
        var needle = Coerce.ToText(DollarText(find)).Chars;
        if (last == -1)
        {
            last = haystack.Length;
        }

        if (last > haystack.Length)
        {
            return Variant.FromInt32(0);
        }

        if (needle.Length == 0)
        {
            return Variant.FromInt32(last);
        }

        for (var position = last - needle.Length; position >= 0; position--)
        {
            if (TextCompare.Compare(haystack.Slice(position, needle.Length), needle, mode) == 0)
            {
                return Variant.FromInt32(position + 1);
            }
        }

        return Variant.FromInt32(0);
    }

    /// <summary>Replace(Expression, Find, Replace, [Start = 1], [Count = -1], [Compare]): the result starts at Start.</summary>
    public static VbaString Replace(in Variant expression, in Variant find, in Variant replacement, in Variant start, in Variant count, in Variant compare, CompareMode moduleMode)
    {
        var first = start.IsMissing ? 1 : Coerce.ToInt32(start);
        var limit = count.IsMissing ? -1 : Coerce.ToInt32(count);
        if (first < 1 || limit < -1)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        var mode = CompareModeArgument(compare, moduleMode);
        var text = Coerce.ToText(expression).Chars;
        var needle = Coerce.ToText(find).Chars;
        var replaceWith = Coerce.ToText(replacement).Chars;
        if (first > text.Length)
        {
            return VbaString.Temporary(string.Empty);
        }

        text = text[(first - 1)..];
        if (needle.Length == 0 || limit == 0)
        {
            return VbaString.Temporary(VbaString.Alloc(text));
        }

        var result = new Bstr.Builder(text.Length);
        var position = 0;
        var replaced = 0;
        while (limit < 0 || replaced < limit)
        {
            var index = IndexOf(text, needle, position, mode);
            if (index < 0)
            {
                break;
            }

            result.Append(text[position..index]);
            result.Append(replaceWith);
            position = index + needle.Length;
            replaced++;
        }

        result.Append(text[position..]);
        return VbaString.Temporary(result.ToVbaString());
    }

    public static Variant UCase(in Variant text) => text.IsNull ? Variant.Null : MapCase(Coerce.ToText(text).Chars, upper: true);

    public static Variant LCase(in Variant text) => text.IsNull ? Variant.Null : MapCase(Coerce.ToText(text).Chars, upper: false);

    /// <summary>
    /// Case mapping the way VBA does it: the regional LCMapString without linguistic casing, so
    /// the dotless i and dotted I are left alone and "ß" stays "ß" (Strings goldens).
    /// </summary>
    public static string MapCase(string value, bool upper)
    {
        ArgumentNullException.ThrowIfNull(value);
        if (value.Length == 0)
        {
            return value;
        }

        var buffer = new char[value.Length];
        var mapped = MapCase(value.AsSpan(), buffer, upper);
        return mapped == value.Length ? new string(buffer) : (upper ? value.ToUpper(CultureInfo.CurrentCulture) : value.ToLower(CultureInfo.CurrentCulture));
    }

    /// <summary>Trim removes leading and trailing spaces only, not tabs or line breaks.</summary>
    public static Variant Trim(in Variant text) => text.IsNull ? Variant.Null : Result(Coerce.ToText(text).Chars.Trim(' '));

    public static Variant LTrim(in Variant text) => text.IsNull ? Variant.Null : Result(Coerce.ToText(text).Chars.TrimStart(' '));

    public static Variant RTrim(in Variant text) => text.IsNull ? Variant.Null : Result(Coerce.ToText(text).Chars.TrimEnd(' '));

    public static Variant Space(in Variant count)
    {
        var n = Coerce.ToInt32(count);
        return n < 0 ? throw VbaErrors.InvalidProcedureCall() : Repeat(' ', n);
    }

    /// <summary>
    /// String(Number, Character): the first character of a string, or the low byte of an Integer
    /// character code through the system code page (8364 gives "¬", -1 gives "ÿ"); an empty string
    /// raises 5, a Null character gives Null (Strings goldens).
    /// </summary>
    public static Variant String(in Variant count, in Variant character)
    {
        var n = Coerce.ToInt32(count);
        if (n < 0)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        if (character.IsNull)
        {
            return Variant.Null;
        }

        char c;
        if (character.IsString)
        {
            var text = character.AsVbaString().Chars;
            if (text.Length == 0)
            {
                throw VbaErrors.InvalidProcedureCall();
            }

            c = text[0];
        }
        else
        {
            var code = Coerce.ToInt16(character);
            c = Ansi.GetString([(byte)(code & 0xFF)])[0];
        }

        return Repeat(c, n);
    }

    public static Variant StrReverse(in Variant text)
    {
        var value = Coerce.ToText(text).Chars;
        var result = VbaString.Alloc(value);
        Bstr.MutableChars(result.Pointer).Reverse();
        return Variant.FromVbaString(result);
    }

    /// <summary>StrComp(String1, String2, [Compare]): -1, 0, or 1 as an Integer; Null gives Null; a bad compare mode raises 5.</summary>
    public static Variant StrComp(in Variant left, in Variant right, in Variant compare, CompareMode moduleMode)
    {
        var mode = CompareModeArgument(compare, moduleMode);
        if (left.IsNull || right.IsNull)
        {
            return Variant.Null;
        }

        return Variant.FromInt16((short)TextCompare.Compare(Coerce.ToText(left), Coerce.ToText(right), mode));
    }

    /// <summary>Split(Expression, [Delimiter = " "], [Limit = -1], [Compare]): a zero-based String array; "" gives an empty array.</summary>
    public static VbaArray Split(in Variant expression, in Variant delimiter, in Variant limit, in Variant compare, CompareMode moduleMode)
    {
        var text = Coerce.ToText(expression).Chars;
        var separator = delimiter.IsMissing ? " ".AsSpan() : Coerce.ToText(delimiter).Chars;
        var max = limit.IsMissing ? -1 : Coerce.ToInt32(limit);
        var mode = CompareModeArgument(compare, moduleMode);
        if (max < -1)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        // A limit of 0 gives an empty array, as does an empty string (Strings golden).
        if (text.Length == 0 || max == 0)
        {
            return ObjectRefs.Owned(VbaArray.Empty(VarType.String));
        }

        var parts = new List<Variant>();
        if (separator.Length == 0)
        {
            parts.Add(Result(text));
        }
        else
        {
            var position = 0;
            while (max < 0 || parts.Count < max - 1)
            {
                var index = IndexOf(text, separator, position, mode);
                if (index < 0)
                {
                    break;
                }

                parts.Add(Result(text[position..index]));
                position = index + separator.Length;
            }

            parts.Add(Result(text[position..]));
        }

        return VbaArray.FromValues(VarType.String, CollectionsMarshal.AsSpan(parts));
    }

    /// <summary>Join(SourceArray, [Delimiter = " "]): String and Variant arrays only (error 5 for others, 13 for a non-array).</summary>
    public static VbaString Join(in Variant source, in Variant delimiter)
    {
        var separator = delimiter.IsMissing ? " ".AsSpan() : Coerce.ToText(delimiter).Chars;
        if (!source.IsArray)
        {
            throw VbaErrors.TypeMismatch();
        }

        var array = source.AsArray();
        if (array.ElementType is not (VarType.String or VarType.Variant))
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        var result = new Bstr.Builder(64);
        var first = true;
        for (var index = 0; index < array.Count; index++)
        {
            var element = array.ElementAt(index);
            if (!first)
            {
                result.Append(separator);
            }

            result.Append(Coerce.ToText(element).Chars);
            first = false;
        }

        return VbaString.Temporary(result.ToVbaString());
    }

    /// <summary>Filter(SourceArray, Match, [Include = True], [Compare]): the elements containing Match, as a zero-based String array.</summary>
    public static VbaArray Filter(in Variant source, in Variant match, in Variant include, in Variant compare, CompareMode moduleMode)
    {
        var mode = CompareModeArgument(compare, moduleMode);
        var keep = include.IsMissing || Coerce.ToBoolean(include);
        var needle = Coerce.ToText(match).Chars;
        if (!source.IsArray)
        {
            throw VbaErrors.TypeMismatch();
        }

        var array = source.AsArray();
        if (array.Rank != 1)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        var matches = new List<Variant>();
        for (var index = 0; index < array.Count; index++)
        {
            var element = array.ElementAt(index);
            var text = Coerce.ToText(element).Chars;
            var found = needle.Length == 0 || IndexOf(text, needle, 0, mode) >= 0;
            if (found == keep)
            {
                matches.Add(Result(text));
            }
        }

        return VbaArray.FromValues(VarType.String, CollectionsMarshal.AsSpan(matches));
    }

    /// <summary>Chr(CharCode): 0 to 255 through the system code page; anything else raises 5.</summary>
    public static Variant Chr(in Variant code)
    {
        var value = Coerce.ToInt32(code);
        if (value is < 0 or > 255)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        return Result(Ansi.GetString([(byte)value]));
    }

    /// <summary>ChrB(CharCode): a string of the one byte; the argument is a Byte, so 256 overflows with error 6 (Strings golden).</summary>
    public static Variant ChrB(in Variant code) => ResultBytes([Coerce.ToByte(code)]);

    /// <summary>ChrW(CharCode): -32768 to 65535 as a UTF-16 code unit; anything else raises 5.</summary>
    public static Variant ChrW(in Variant code)
    {
        var value = Coerce.ToInt32(code);
        if (value is < -32768 or > 65535)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        return Repeat((char)(value & 0xFFFF), 1);
    }

    /// <summary>Asc(String): the code of the first character in the system code page; "" raises 5.</summary>
    public static short Asc(in Variant text)
    {
        var value = Coerce.ToText(text).Chars;
        if (value.Length == 0)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        var bytes = Ansi.GetBytes(value[..1].ToArray());
        return bytes.Length == 1 ? bytes[0] : (short)((bytes[0] << 8) | bytes[1]);
    }

    /// <summary>AscB(String): the first byte of the string as a Byte; "" raises 5 (Strings golden: AscB(ChrW(8364)) is 172).</summary>
    public static byte AscB(in Variant text)
    {
        var bytes = Coerce.ToText(text).Bytes;
        return bytes.Length == 0 ? throw VbaErrors.InvalidProcedureCall() : bytes[0];
    }

    /// <summary>AscW(String): the first UTF-16 code unit as a signed Integer; "" raises 5.</summary>
    public static short AscW(in Variant text)
    {
        var value = Coerce.ToText(text).Chars;
        return value.Length == 0 ? throw VbaErrors.InvalidProcedureCall() : (short)value[0];
    }

    /// <summary>StrConv(String, Conversion): case conversions, and the Unicode/ANSI byte reinterpretations; Null gives Null.</summary>
    public static Variant StrConv(in Variant text, in Variant conversion)
    {
        var mode = Coerce.ToInt32(conversion);
        if (text.IsNull)
        {
            return Variant.Null;
        }

        const int caseMask = 3;
        var caseMode = mode & caseMask;
        var rest = mode & ~caseMask;
        if (rest is not (0 or 64 or 128) || (rest == 64 && caseMode != 0) || (rest == 128 && caseMode != 0))
        {
            // vbWide, vbNarrow, vbKatakana, vbHiragana, and mixed flags raise 5 outside East Asian locales.
            throw VbaErrors.InvalidProcedureCall();
        }

        ReadOnlySpan<byte> bytes;
        VbaString value = default;
        if (text.IsArray && text.AsArray().ElementType == VarType.Byte)
        {
            bytes = text.AsArray().ToBytes();
        }
        else
        {
            value = Coerce.ToText(text);
            bytes = value.Bytes;
        }

        switch (rest)
        {
            case 64:
                // vbUnicode: each byte of the ANSI form becomes a character.
                return Result(Ansi.GetString(bytes));
            case 128:
                // vbFromUnicode: the ANSI bytes of the string, kept as bytes, so an odd count survives (Strings golden).
                return ResultBytes(Ansi.GetBytes(value.Chars.ToArray()));
        }

        var chars = text.IsArray ? Encoding.Unicode.GetString(bytes).AsSpan() : value.Chars;
        return caseMode switch
        {
            1 => MapCase(chars, upper: true),
            2 => MapCase(chars, upper: false),
            3 => ProperCase(chars),
            _ => Result(chars),
        };
    }

    /// <summary>The compare mode argument of a string function: omitted uses the module's Option Compare; 1 is text, anything else binary (Strings goldens).</summary>
    public static CompareMode CompareModeArgument(in Variant compare, CompareMode moduleMode)
    {
        if (compare.IsMissing)
        {
            return moduleMode;
        }

        return Coerce.ToInt32(compare) == 1 ? CompareMode.Text : CompareMode.Binary;
    }

    /// <summary>The 0-based index of <paramref name="needle"/> in <paramref name="text"/> at or after <paramref name="start"/>, under the compare mode.</summary>
    public static int IndexOf(string text, string needle, int start, CompareMode mode)
    {
        ArgumentNullException.ThrowIfNull(text);
        ArgumentNullException.ThrowIfNull(needle);
        return IndexOf(text.AsSpan(), needle.AsSpan(), start, mode);
    }

    public static int IndexOf(ReadOnlySpan<char> text, ReadOnlySpan<char> needle, int start, CompareMode mode)
    {
        if (start > text.Length)
        {
            return -1;
        }

        var index = mode == CompareMode.Binary
            ? text[start..].IndexOf(needle, StringComparison.Ordinal)
            : CultureInfo.CurrentCulture.CompareInfo.IndexOf(text[start..], needle, CompareOptions.IgnoreCase | CompareOptions.IgnoreKanaType | CompareOptions.IgnoreWidth);
        return index < 0 ? -1 : index + start;
    }

    /// <summary>A new BSTR the statement owns, holding the code units.</summary>
    private static Variant Result(ReadOnlySpan<char> chars) => Variant.FromVbaString(VbaString.Alloc(chars));

    /// <summary>A new BSTR the statement owns, holding exactly these bytes.</summary>
    private static Variant ResultBytes(ReadOnlySpan<byte> bytes) => Variant.FromVbaString(VbaString.AllocBytes(bytes));

    private static Variant Repeat(char c, int count)
    {
        var result = VbaString.AllocLength(count);
        Bstr.MutableChars(result.Pointer).Fill(c);
        return Variant.FromVbaString(result);
    }

    /// <summary>A String-declared argument: Null raises 94 where a Variant one would pass it through (MS-VBAL 6.1.2).</summary>
    private static Variant DollarText(in Variant value) => value.IsNull ? throw VbaErrors.InvalidUseOfNull() : value;

    private static Variant MapCase(ReadOnlySpan<char> value, bool upper)
    {
        var result = VbaString.AllocLength(value.Length);
        var buffer = Bstr.MutableChars(result.Pointer);
        var mapped = MapCase(value, buffer, upper);
        if (mapped != value.Length)
        {
            (upper ? value.ToString().ToUpper(CultureInfo.CurrentCulture) : value.ToString().ToLower(CultureInfo.CurrentCulture)).AsSpan().CopyTo(buffer);
        }

        return Variant.FromVbaString(result);
    }

    /// <summary>LCMapStringEx into the buffer; the count mapped, or 0 when the call is unavailable, in which case the caller falls back.</summary>
    private static unsafe int MapCase(ReadOnlySpan<char> value, Span<char> buffer, bool upper)
    {
        if (value.Length == 0)
        {
            return 0;
        }

        if (!OperatingSystem.IsWindows())
        {
            return 0;
        }

        const uint LcMapLowerCase = 0x00000100;
        const uint LcMapUpperCase = 0x00000200;
        fixed (char* source = value)
        fixed (char* destination = buffer)
        {
            return LCMapStringEx(null, upper ? LcMapUpperCase : LcMapLowerCase, source, value.Length, destination, buffer.Length, 0, 0, 0);
        }
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
    private static extern unsafe int LCMapStringEx(string? localeName, uint flags, char* source, int sourceLength, char* destination, int destinationLength, nint versionInformation, nint reserved, nint sortHandle);

    private static Variant ProperCase(ReadOnlySpan<char> value)
    {
        var lowered = MapCase(value, upper: false);
        var characters = Bstr.MutableChars(lowered.AsVbaString().Pointer);
        var startOfWord = true;
        Span<char> one = stackalloc char[1];
        for (var i = 0; i < characters.Length; i++)
        {
            var c = characters[i];
            if (c is '\0' or '\t' or '\n' or '\v' or '\f' or '\r' or ' ')
            {
                startOfWord = true;
            }
            else if (startOfWord)
            {
                if (MapCase(characters.Slice(i, 1), one, upper: true) == 1)
                {
                    characters[i] = one[0];
                }

                startOfWord = false;
            }
        }

        return lowered;
    }

    private static Encoding CreateAnsi()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        return Encoding.GetEncoding(CultureInfo.CurrentCulture.TextInfo.ANSICodePage);
    }
}
