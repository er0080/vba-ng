using System.Collections.Frozen;

using VbaNg.Runtime;

namespace VbaNg.Compiler.Binding;

/// <summary>
/// The VBA standard library as the binder sees it (MS-VBAL 6.1): the intrinsic functions with the
/// C# each compiles to, and the built-in constants with the type and value Excel records for
/// them (Constants golden). Template placeholders:
/// <list type="bullet">
/// <item><c>{0}</c>: argument 0 as a Variant, or <c>Variant.Missing</c> when omitted.</item>
/// <item><c>{0|text}</c>: argument 0 as a Variant, or the literal text when omitted.</item>
/// <item><c>{0:d}</c>, <c>{0:i}</c>: argument 0 as a double or int, 0 when omitted; <c>{0:d|0.1}</c> sets the default.</item>
/// <item><c>{rest:n}</c>: arguments n and up as a Variant collection expression.</item>
/// <item><c>{cmp}</c>: the module's Option Compare mode; <c>{base}</c>: its Option Base; <c>{rnd}</c>: the shared Rnd generator.</item>
/// </list>
/// </summary>
public static class StandardLibrary
{
    private static readonly FrozenSet<string> EnumNames = new[]
    {
        "VbVarType", "VbCallType", "VbCompareMethod", "VbDayOfWeek", "VbFirstWeekOfYear", "VbDateTimeFormat",
        "VbTriState", "VbStrConv", "VbMsgBoxStyle", "VbMsgBoxResult", "VbQueryClose", "VbAppWinStyle",
        "VbFileAttribute", "VbIMEStatus",
    }.ToFrozenSet(StringComparer.OrdinalIgnoreCase);

    private static readonly FrozenSet<string> LibraryModules = new[]
    {
        "Math", "Strings", "Conversion", "DateTime", "Information", "Interaction", "FileSystem", "Financial", "Constants",
    }.ToFrozenSet(StringComparer.OrdinalIgnoreCase);

    private const string R = "global::VbaNg.Runtime.";
    private const string L = "global::VbaNg.Runtime.Library.";
    private const string Missing = R + "Variant.Missing";

    private static readonly Dictionary<string, IntrinsicSymbol> Functions = new(StringComparer.OrdinalIgnoreCase);
    private static readonly Dictionary<string, ConstantSymbol> Constants = new(StringComparer.OrdinalIgnoreCase);

    static StandardLibrary()
    {
        // Conversion (MS-VBAL 6.1.2.3)
        Add("CBool", VbaType.Boolean, 1, 1, L + "Conversion.CBool({0})");
        Add("CByte", VbaType.Byte, 1, 1, L + "Conversion.CByte({0})");
        Add("CCur", VbaType.Currency, 1, 1, L + "Conversion.CCur({0})");
        Add("CDate", VbaType.Date, 1, 1, L + "Conversion.CDate({0})");
        Add("CDbl", VbaType.Double, 1, 1, L + "Conversion.CDbl({0})");
        Add("CDec", VbaType.Variant, VbaType.Decimal, 1, 1, L + "Conversion.CDec({0})");
        Add("CInt", VbaType.Integer, 1, 1, L + "Conversion.CInt({0})");
        Add("CLng", VbaType.Long, 1, 1, L + "Conversion.CLng({0})");
        Add("CLngLng", VbaType.LongLong, 1, 1, L + "Conversion.CLngLng({0})");
        Add("CLngPtr", VbaType.LongLong, 1, 1, L + "Conversion.CLngPtr({0})");
        Add("CSng", VbaType.Single, 1, 1, L + "Conversion.CSng({0})");
        Add("CStr", VbaType.String, 1, 1, L + "Conversion.CStr({0})");
        Add("CVar", VbaType.Variant, 1, 1, L + "Conversion.CVar({0})");
        Add("CVErr", VbaType.Variant, 1, 1, R + "Variant.FromError(" + L + "Conversion.CVErr({0}))");
        Add("Error", VbaType.Variant, 0, 1, R + "Variant.ViewString(" + R + "Err.Current.Description)", L + "Conversion.Error({0})");
        Add("Fix", VbaType.Variant, 1, 1, L + "Conversion.Fix({0})");
        Add("Int", VbaType.Variant, 1, 1, L + "Conversion.Int({0})");
        Add("Hex", VbaType.Variant, 1, 1, L + "Conversion.Hex({0})");
        Add("Oct", VbaType.Variant, 1, 1, L + "Conversion.Oct({0})");
        Add("Str", VbaType.Variant, 1, 1, L + "Conversion.Str({0})");
        Add("Val", VbaType.Double, 1, 1, L + "Conversion.Val({0})");

        // Information (MS-VBAL 6.1.2.5)
        Add("RGB", VbaType.Long, 3, 3, L + "Information.RGB({0}, {1}, {2})");
        Add("QBColor", VbaType.Long, 1, 1, L + "Information.QBColor({0})");
        Add("IsArray", VbaType.Boolean, 1, 1, L + "Information.IsArray({0})");
        Add("IsDate", VbaType.Boolean, 1, 1, L + "Information.IsDate({0})");
        Add("IsEmpty", VbaType.Boolean, 1, 1, L + "Information.IsEmpty({0})");
        Add("IsError", VbaType.Boolean, 1, 1, L + "Information.IsError({0})");
        Add("IsMissing", VbaType.Boolean, 1, 1, L + "Information.IsMissing({0})");
        Add("IsNull", VbaType.Boolean, 1, 1, L + "Information.IsNull({0})");
        Add("IsNumeric", VbaType.Boolean, 1, 1, L + "Information.IsNumeric({0})");
        Add("IsObject", VbaType.Boolean, 1, 1, L + "Information.IsObject({0})");
        Add("LBound", VbaType.Long, 1, 2, L + "Information.LBound({0}, {1:i|1})");
        Add("UBound", VbaType.Long, 1, 2, L + "Information.UBound({0}, {1:i|1})");
        Add("TypeName", VbaType.String, 1, 1, L + "Information.TypeName({0})");
        Add("VarType", VbaType.Long, 1, 1, L + "Information.VarType({0})");

        // Interaction (MS-VBAL 6.1.2.6)
        Add("IIf", VbaType.Variant, 3, 3, L + "Interaction.IIf({0}, {1}, {2})");
        Add("Choose", VbaType.Variant, 2, int.MaxValue, L + "Interaction.Choose({0}, {rest:1})");
        Add("Switch", VbaType.Variant, 0, int.MaxValue, L + "Interaction.Switch({rest:0})");
        Add("DoEvents", VbaType.Integer, 0, 0, L + "Interaction.DoEvents()");
        Add("Partition", VbaType.Variant, 4, 4, L + "Interaction.Partition({0}, {1}, {2}, {3})");
        Add("Environ", VbaType.String, 1, 1, L + "Interaction.Environ({0})");

        // Financial (MS-VBAL 6.1.2.4)
        Add("Pmt", VbaType.Double, 3, 5, L + "Financial.Pmt({0:d}, {1:d}, {2:d}, {3:d}, {4:d})");
        Add("FV", VbaType.Double, 3, 5, L + "Financial.FV({0:d}, {1:d}, {2:d}, {3:d}, {4:d})");
        Add("PV", VbaType.Double, 3, 5, L + "Financial.PV({0:d}, {1:d}, {2:d}, {3:d}, {4:d})");
        Add("NPer", VbaType.Double, 3, 5, L + "Financial.NPer({0:d}, {1:d}, {2:d}, {3:d}, {4:d})");
        Add("IPmt", VbaType.Double, 4, 6, L + "Financial.IPmt({0:d}, {1:d}, {2:d}, {3:d}, {4:d}, {5:d})");
        Add("PPmt", VbaType.Double, 4, 6, L + "Financial.PPmt({0:d}, {1:d}, {2:d}, {3:d}, {4:d}, {5:d})");
        Add("DDB", VbaType.Double, 4, 5, L + "Financial.DDB({0:d}, {1:d}, {2:d}, {3:d}, {4:d|2})");
        Add("SLN", VbaType.Double, 3, 3, L + "Financial.SLN({0:d}, {1:d}, {2:d})");
        Add("SYD", VbaType.Double, 4, 4, L + "Financial.SYD({0:d}, {1:d}, {2:d}, {3:d})");
        Add("NPV", VbaType.Double, 2, 2, L + "Financial.NPV({0:d}, " + L + "Financial.Flows({1}))");
        Add("IRR", VbaType.Double, 1, 2, L + "Financial.IRR(" + L + "Financial.Flows({0}), {1:d|0.1})");
        Add("MIRR", VbaType.Double, 3, 3, L + "Financial.MIRR(" + L + "Financial.Flows({0}), {1:d}, {2:d})");
        Add("Rate", VbaType.Double, 3, 6, L + "Financial.Rate({0:d}, {1:d}, {2:d}, {3:d}, {4:d}, {5:d|0.1})");

        // Strings (MS-VBAL 6.1.2.8)
        Add("Len", VbaType.Variant, 1, 1, L + "Strings.Len({0})");
        Add("LenB", VbaType.Variant, 1, 1, L + "Strings.LenB({0})");
        Add("Left", VbaType.Variant, 2, 2, L + "Strings.Left({0}, {1})");
        Add("Right", VbaType.Variant, 2, 2, L + "Strings.Right({0}, {1})");
        Add("Mid", VbaType.Variant, 2, 3, L + "Strings.Mid({0}, {1}, {2})");
        Add("InStr", VbaType.Variant, 2, 4, L + "Strings.InStr({rest:0}, {cmp})");
        Add("LeftB", VbaType.Variant, 2, 2, L + "Strings.LeftB({0}, {1})");
        Add("RightB", VbaType.Variant, 2, 2, L + "Strings.RightB({0}, {1})");
        Add("MidB", VbaType.Variant, 2, 3, L + "Strings.MidB({0}, {1}, {2})");
        Add("InStrB", VbaType.Variant, 2, 4, L + "Strings.InStrB({rest:0}, {cmp})");
        Add("InStrRev", VbaType.Variant, 2, 4, L + "Strings.InStrRev({0}, {1}, {2}, {3}, {cmp})");
        Add("Replace", VbaType.String, 3, 6, L + "Strings.Replace({0}, {1}, {2}, {3}, {4}, {5}, {cmp})");
        Add("UCase", VbaType.Variant, 1, 1, L + "Strings.UCase({0})");
        Add("LCase", VbaType.Variant, 1, 1, L + "Strings.LCase({0})");
        Add("Trim", VbaType.Variant, 1, 1, L + "Strings.Trim({0})");
        Add("LTrim", VbaType.Variant, 1, 1, L + "Strings.LTrim({0})");
        Add("RTrim", VbaType.Variant, 1, 1, L + "Strings.RTrim({0})");
        Add("Space", VbaType.Variant, 1, 1, L + "Strings.Space({0})");
        Add("String", VbaType.Variant, 2, 2, L + "Strings.String({0}, {1})");
        Add("StrReverse", VbaType.Variant, 1, 1, L + "Strings.StrReverse({0})");
        Add("StrComp", VbaType.Variant, 2, 3, L + "Strings.StrComp({0}, {1}, {2}, {cmp})");
        Add("Split", VbaType.Variant, 1, 4, R + "Variant.FromArray(" + L + "Strings.Split({0}, {1}, {2}, {3}, {cmp}))");
        Add("Join", VbaType.String, 1, 2, L + "Strings.Join({0}, {1})");
        Add("Filter", VbaType.Variant, 2, 4, R + "Variant.FromArray(" + L + "Strings.Filter({0}, {1}, {2}, {3}, {cmp}))");
        Add("Chr", VbaType.Variant, 1, 1, L + "Strings.Chr({0})");
        Add("ChrW", VbaType.Variant, 1, 1, L + "Strings.ChrW({0})");
        Add("Asc", VbaType.Integer, 1, 1, L + "Strings.Asc({0})");
        Add("ChrB", VbaType.Variant, 1, 1, L + "Strings.ChrB({0})");
        Add("AscB", VbaType.Byte, 1, 1, L + "Strings.AscB({0})");
        Add("AscW", VbaType.Integer, 1, 1, L + "Strings.AscW({0})");
        Add("StrConv", VbaType.Variant, 2, 3, L + "Strings.StrConv({0}, {1})");
        Add("Format", VbaType.Variant, 1, 4, L + "VbaFormat.Format({0}, {1}, {2}, {3})");
        Add("FormatNumber", VbaType.String, 1, 5, L + "VbaFormat.FormatNumber({0}, {1}, {2}, {3}, {4})");
        Add("FormatPercent", VbaType.String, 1, 5, L + "VbaFormat.FormatPercent({0}, {1}, {2}, {3}, {4})");
        Add("FormatCurrency", VbaType.String, 1, 5, L + "VbaFormat.FormatCurrency({0}, {1}, {2}, {3}, {4})");
        Add("FormatDateTime", VbaType.String, 1, 2, L + "VbaFormat.FormatDateTime({0}, {1})");
        Add("MonthName", VbaType.String, 1, 2, L + "VbaFormat.MonthName({0}, {1})");
        Add("WeekdayName", VbaType.String, 1, 3, L + "VbaFormat.WeekdayName({0}, {1}, {2})");

        // Math (MS-VBAL 6.1.2.7)
        Add("Abs", VbaType.Variant, 1, 1, L + "VbaMath.Abs({0})");
        Add("Sgn", VbaType.Variant, 1, 1, L + "VbaMath.Sgn({0})");
        Add("Sqr", VbaType.Double, 1, 1, L + "VbaMath.Sqr({0})");
        Add("Exp", VbaType.Double, 1, 1, L + "VbaMath.Exp({0})");
        Add("Log", VbaType.Double, 1, 1, L + "VbaMath.Log({0})");
        Add("Sin", VbaType.Double, 1, 1, L + "VbaMath.Sin({0})");
        Add("Cos", VbaType.Double, 1, 1, L + "VbaMath.Cos({0})");
        Add("Tan", VbaType.Double, 1, 1, L + "VbaMath.Tan({0})");
        Add("Atn", VbaType.Double, 1, 1, L + "VbaMath.Atn({0})");
        Add("Round", VbaType.Variant, 1, 2, L + "VbaMath.Round({0}, {1})");
        Add("Rnd", VbaType.Single, 0, 1, "{rnd}.Next()", "{rnd}.Next({0})");
        AddSub("Randomize", 0, 1, "{rnd}.Randomize(" + Missing + ")", "{rnd}.Randomize({0})");

        // DateTime (MS-VBAL 6.1.2.2)
        Add("Now", VbaType.Date, 0, 0, L + "VbaDateTime.Now()");
        Add("Date", VbaType.Date, 0, 0, L + "VbaDateTime.Today()");
        Add("Time", VbaType.Date, 0, 0, L + "VbaDateTime.TimeOfDay()");
        Add("Timer", VbaType.Single, 0, 0, L + "VbaDateTime.Timer()");
        Add("DateSerial", VbaType.Date, 3, 3, L + "VbaDateTime.DateSerial({0}, {1}, {2})");
        Add("TimeSerial", VbaType.Date, 3, 3, L + "VbaDateTime.TimeSerial({0}, {1}, {2})");
        Add("DateAdd", VbaType.Variant, 3, 3, L + "VbaDateTime.DateAdd({0}, {1}, {2})");
        Add("DateDiff", VbaType.Variant, 3, 5, L + "VbaDateTime.DateDiff({0}, {1}, {2}, {3}, {4})");
        Add("DatePart", VbaType.Variant, 2, 4, L + "VbaDateTime.DatePart({0}, {1}, {2}, {3})");
        Add("Year", VbaType.Variant, 1, 1, L + "VbaDateTime.Year({0})");
        Add("Month", VbaType.Variant, 1, 1, L + "VbaDateTime.Month({0})");
        Add("Day", VbaType.Variant, 1, 1, L + "VbaDateTime.Day({0})");
        Add("Hour", VbaType.Variant, 1, 1, L + "VbaDateTime.Hour({0})");
        Add("Minute", VbaType.Variant, 1, 1, L + "VbaDateTime.Minute({0})");
        Add("Second", VbaType.Variant, 1, 1, L + "VbaDateTime.Second({0})");
        Add("Weekday", VbaType.Variant, 1, 2, L + "VbaDateTime.Weekday({0}, {1})");
        Add("DateValue", VbaType.Date, 1, 1, L + "VbaDateTime.DateValue({0})");
        Add("TimeValue", VbaType.Date, 1, 1, L + "VbaDateTime.TimeValue({0})");

        // COM objects (ARCHITECTURE.md section 6): created through the host's provider, reached late-bound by name.
        Add("CreateObject", VbaType.Variant, 1, 2, L + "Interaction.CreateObject({0}, {1})");
        Add("GetObject", VbaType.Variant, 0, 2, L + "Interaction.GetObject({0}, {1})");
        Add("CallByName", VbaType.Variant, 3, int.MaxValue, L + "Interaction.CallByName({0}, {1}, {2}, {rest:3})");

        // Host services (ARCHITECTURE.md D8): dialogs go through the host, Shell and AppActivate to the operating system.
        // The result is a VbMsgBoxResult, so the intrinsic is a Long and the emitted expression has
        // to be one: the host answers with a Variant, which is coerced here rather than wrapped as
        // a Variant by the caller, which would not compile (VBA0003).
        Add("MsgBox", VbaType.Long, 1, 5, R + "Coerce.ToInt32(" + L + "Interaction.MsgBox({0}, {1}, {2}))");
        Add("InputBox", VbaType.String, 1, 7, L + "Interaction.InputBox({0}, {1}, {2})");
        Add("Shell", VbaType.Double, 1, 2, L + "Interaction.Shell({0}, {1})");
        AddSub("AppActivate", 1, 2, L + "Interaction.AppActivate({0}, {1})");
        AddSub("Beep", 0, 0, L + "Interaction.Beep()");
        AddSub("SendKeys", 1, 2, L + "Interaction.SendKeys({0}, {1})");

        // The FileSystem module (MS-VBAL 6.1.2, FileSystem golden): files, folders, attributes, and the settings of Interaction.
        Add("CurDir", VbaType.String, 0, 1, L + "FileSystem.CurDir({0})");
        Add("Dir", VbaType.String, 0, 2, L + "FileSystem.Dir({0}, {1})");
        Add("FileLen", VbaType.Long, 1, 1, L + "FileSystem.FileLen({0})");
        Add("FileDateTime", VbaType.Date, 1, 1, L + "FileSystem.FileDateTime({0})");
        Add("GetAttr", VbaType.Long, 1, 1, L + "FileSystem.GetAttr({0})");
        Add("FreeFile", VbaType.Integer, 0, 1, L + "FileSystem.FreeFile({0})");
        Add("EOF", VbaType.Boolean, 1, 1, L + "FileSystem.Eof({0})");
        Add("LOF", VbaType.Long, 1, 1, L + "FileSystem.Lof({0})");
        Add("Loc", VbaType.Long, 1, 1, L + "FileSystem.Loc({0})");
        Add("Seek", VbaType.Long, 1, 1, L + "FileSystem.SeekPosition({0})");
        Add("Input", VbaType.String, 2, 2, L + "FileSystem.InputChars({0}, {1})");
        Add("GetSetting", VbaType.String, 3, 4, L + "FileSystem.GetSetting({0}, {1}, {2}, {3})");
        Add("GetAllSettings", VbaType.Variant, 2, 2, L + "FileSystem.GetAllSettings({0}, {1})");
        AddSub("Kill", 1, 1, L + "FileSystem.Kill({0})");
        AddSub("MkDir", 1, 1, L + "FileSystem.MkDir({0})");
        AddSub("RmDir", 1, 1, L + "FileSystem.RmDir({0})");
        AddSub("ChDir", 1, 1, L + "FileSystem.ChDir({0})");
        AddSub("ChDrive", 1, 1, L + "FileSystem.ChDrive({0})");
        AddSub("SetAttr", 2, 2, L + "FileSystem.SetAttr({0}, {1})");
        AddSub("FileCopy", 2, 2, L + "FileSystem.FileCopy({0}, {1})");
        AddSub("SaveSetting", 4, 4, L + "FileSystem.SaveSetting({0}, {1}, {2}, {3})");
        AddSub("DeleteSetting", 1, 3, L + "FileSystem.DeleteSetting({0}, {1}, {2})");

        // InputB reads bytes into a byte string (FileSystem golden); Command reports the command line of a standalone
        // program, which Excel does not have.
        Add("InputB", VbaType.String, 2, 2, L + "FileSystem.InputBytes({0}, {1})");
        Functions["Command"] = new IntrinsicSymbol("Command", VbaType.String, VbaType.String, 0, 0, [string.Empty]) { Pending = "never: it reports the command line of a standalone program, which Excel does not have" };

        // The address functions bind as pointers in Binder.Pointers.cs; these entries name them, and a call of
        // any other shape (a named argument, none at all) is reported as unsupported.
        foreach (var name in new[] { "VarPtr", "ObjPtr", "StrPtr" })
        {
            Functions[name] = new IntrinsicSymbol(name, VbaType.LongLong, VbaType.LongLong, 1, 1, [string.Empty])
            {
                Pending = "a direct call with one argument: only that shape binds as an address",
            };
        }

        // IMEStatus reports the Far East input method editor; it has no equivalent here.
        Functions["IMEStatus"] = new IntrinsicSymbol("IMEStatus", VbaType.Variant, VbaType.Variant, 0, 0, [string.Empty]) { Pending = "never: it reports the Far East input method editor, which has no equivalent here" };

        AddSub("Debug.Assert", 1, 1, R + "Debug.Assert({0})");

        // The Assert module of test procedures (ARCHITECTURE.md section 8): a vba-ng addition, bound only
        // when the project declares nothing named Assert. Equality follows the module's Option Compare.
        AddSub("Assert.AreEqual", 2, 3, L + "Assert.AreEqual({0}, {1}, {2}, {cmp})");
        AddSub("Assert.AreNotEqual", 2, 3, L + "Assert.AreNotEqual({0}, {1}, {2}, {cmp})");
        AddSub("Assert.IsTrue", 1, 2, L + "Assert.IsTrue({0}, {1})");
        AddSub("Assert.IsFalse", 1, 2, L + "Assert.IsFalse({0}, {1})");
        AddSub("Assert.IsNothing", 1, 2, L + "Assert.IsNothing({0}, {1})");
        AddSub("Assert.IsNotNothing", 1, 2, L + "Assert.IsNotNothing({0}, {1})");
        AddSub("Assert.Fail", 0, 1, L + "Assert.Fail({0})");
        AddConstants();
    }

    public static bool TryGetFunction(string name, out IntrinsicSymbol intrinsic)
    {
        ArgumentNullException.ThrowIfNull(name);
        return Functions.TryGetValue(name, out intrinsic!);
    }

    public static bool TryGetConstant(string name, out ConstantSymbol constant)
    {
        ArgumentNullException.ThrowIfNull(name);
        return Constants.TryGetValue(name, out constant!);
    }

    /// <summary>The modules of the VBA library (MS-VBAL 6.1), which may qualify its functions as VBA.Math.Abs.</summary>
    public static bool IsLibraryModule(string name)
    {
        ArgumentNullException.ThrowIfNull(name);
        return LibraryModules.Contains(name);
    }

    /// <summary>The enums the VBA library declares (MS-VBAL 6.1.1); their members are the vb constants and their values are Long.</summary>
    public static bool IsEnumName(string name)
    {
        ArgumentNullException.ThrowIfNull(name);
        return EnumNames.Contains(name);
    }

    /// <summary>The parameter names of a library function as MS-VBAL 6.1 documents them, for named arguments; empty for a function that takes none by name.</summary>
    public static List<string> ParameterNames(string function)
    {
        ArgumentNullException.ThrowIfNull(function);
        return Names.TryGetValue(function, out var names) ? names : [];
    }

    private static readonly Dictionary<string, List<string>> Names = new(StringComparer.OrdinalIgnoreCase)
    {
        // Strings (MS-VBAL 6.1.2.11)
        ["Left"] = ["String", "Length"],
        ["Right"] = ["String", "Length"],
        ["Mid"] = ["String", "Start", "Length"],
        ["InStrRev"] = ["StringCheck", "StringMatch", "Start", "Compare"],
        ["Replace"] = ["Expression", "Find", "Replace", "Start", "Count", "Compare"],
        ["Split"] = ["Expression", "Delimiter", "Limit", "Compare"],
        ["Join"] = ["SourceArray", "Delimiter"],
        ["Filter"] = ["SourceArray", "Match", "Include", "Compare"],
        ["StrConv"] = ["String", "Conversion", "LocaleID"],
        ["StrReverse"] = ["Expression"],
        ["Trim"] = ["String"],
        ["LTrim"] = ["String"],
        ["RTrim"] = ["String"],
        ["UCase"] = ["String"],
        ["LCase"] = ["String"],
        ["Len"] = ["Expression"],
        ["LenB"] = ["Expression"],
        ["Space"] = ["Number"],
        ["String"] = ["Number", "Character"],
        ["Chr"] = ["CharCode"],
        ["ChrW"] = ["CharCode"],
        ["Asc"] = ["String"],
        ["AscB"] = ["String"],
        ["ChrB"] = ["CharCode"],
        ["LeftB"] = ["String", "Length"],
        ["RightB"] = ["String", "Length"],
        ["MidB"] = ["String", "Start", "Length"],
        ["InStrB"] = ["Start", "String1", "String2", "Compare"],
        ["AscW"] = ["String"],
        ["Format"] = ["Expression", "Format", "FirstDayOfWeek", "FirstWeekOfYear"],
        ["FormatNumber"] = ["Expression", "NumDigitsAfterDecimal", "IncludeLeadingDigit", "UseParensForNegativeNumbers", "GroupDigits"],
        ["FormatCurrency"] = ["Expression", "NumDigitsAfterDecimal", "IncludeLeadingDigit", "UseParensForNegativeNumbers", "GroupDigits"],
        ["FormatPercent"] = ["Expression", "NumDigitsAfterDecimal", "IncludeLeadingDigit", "UseParensForNegativeNumbers", "GroupDigits"],
        ["FormatDateTime"] = ["Expression", "NamedFormat"],
        ["MonthName"] = ["Month", "Abbreviate"],
        ["WeekdayName"] = ["Weekday", "Abbreviate", "FirstDayOfWeek"],
        // Math (6.1.2.9)
        ["Round"] = ["Number", "NumDigitsAfterDecimal"],
        ["Abs"] = ["Number"],
        ["Sgn"] = ["Number"],
        ["Sqr"] = ["Number"],
        ["Exp"] = ["Number"],
        ["Log"] = ["Number"],
        ["Sin"] = ["Number"],
        ["Cos"] = ["Number"],
        ["Tan"] = ["Number"],
        ["Atn"] = ["Number"],
        ["Rnd"] = ["Number"],
        // Conversion (6.1.2.3)
        ["CBool"] = ["Expression"],
        ["CByte"] = ["Expression"],
        ["CCur"] = ["Expression"],
        ["CDate"] = ["Expression"],
        ["CDbl"] = ["Expression"],
        ["CDec"] = ["Expression"],
        ["CInt"] = ["Expression"],
        ["CLng"] = ["Expression"],
        ["CLngLng"] = ["Expression"],
        ["CLngPtr"] = ["Expression"],
        ["CSng"] = ["Expression"],
        ["CStr"] = ["Expression"],
        ["CVar"] = ["Expression"],
        ["CVErr"] = ["Expression"],
        ["Val"] = ["String"],
        ["Str"] = ["Number"],
        ["Hex"] = ["Number"],
        ["Oct"] = ["Number"],
        ["Int"] = ["Number"],
        ["Fix"] = ["Number"],
        ["Error"] = ["ErrorNumber"],
        // Information (6.1.2.8)
        ["IsArray"] = ["VarName"],
        ["IsDate"] = ["Expression"],
        ["IsEmpty"] = ["Expression"],
        ["IsError"] = ["Expression"],
        ["IsMissing"] = ["ArgName"],
        ["IsNull"] = ["Expression"],
        ["IsNumeric"] = ["Expression"],
        ["IsObject"] = ["Expression"],
        ["TypeName"] = ["VarName"],
        ["VarType"] = ["VarName"],
        ["RGB"] = ["Red", "Green", "Blue"],
        ["QBColor"] = ["Color"],
        // Interaction (6.1.2.7)
        ["IIf"] = ["Expression", "TruePart", "FalsePart"],
        ["Choose"] = ["Index"],
        ["Partition"] = ["Number", "Start", "Stop", "Interval"],
        ["MsgBox"] = ["Prompt", "Buttons", "Title", "HelpFile", "Context"],
        ["InputBox"] = ["Prompt", "Title", "Default", "XPos", "YPos", "HelpFile", "Context"],
        ["Shell"] = ["PathName", "WindowStyle"],
        ["Environ"] = ["Expression"],
        ["CreateObject"] = ["Class", "ServerName"],
        ["GetObject"] = ["PathName", "Class"],
        ["CallByName"] = ["Object", "ProcName", "CallType"],
        ["GetSetting"] = ["AppName", "Section", "Key", "Default"],
        ["SaveSetting"] = ["AppName", "Section", "Key", "Setting"],
        ["GetAllSettings"] = ["AppName", "Section"],
        ["DeleteSetting"] = ["AppName", "Section", "Key"],
        ["AppActivate"] = ["Title", "Wait"],
        ["SendKeys"] = ["String", "Wait"],
        // DateTime (6.1.2.4)
        ["DateAdd"] = ["Interval", "Number", "Date"],
        ["DateDiff"] = ["Interval", "Date1", "Date2", "FirstDayOfWeek", "FirstWeekOfYear"],
        ["DatePart"] = ["Interval", "Date", "FirstDayOfWeek", "FirstWeekOfYear"],
        ["DateSerial"] = ["Year", "Month", "Day"],
        ["DateValue"] = ["Date"],
        ["TimeSerial"] = ["Hour", "Minute", "Second"],
        ["TimeValue"] = ["Time"],
        ["Day"] = ["Date"],
        ["Month"] = ["Date"],
        ["Year"] = ["Date"],
        ["Hour"] = ["Time"],
        ["Minute"] = ["Time"],
        ["Second"] = ["Time"],
        ["Weekday"] = ["Date", "FirstDayOfWeek"],
        // FileSystem (6.1.2.5)
        ["Dir"] = ["PathName", "Attributes"],
        ["FileLen"] = ["PathName"],
        ["FileDateTime"] = ["PathName"],
        ["GetAttr"] = ["PathName"],
        ["SetAttr"] = ["PathName", "Attributes"],
        ["Kill"] = ["PathName"],
        ["MkDir"] = ["Path"],
        ["RmDir"] = ["Path"],
        ["ChDir"] = ["Path"],
        ["ChDrive"] = ["Drive"],
        ["CurDir"] = ["Drive"],
        ["FileCopy"] = ["Source", "Destination"],
        ["FreeFile"] = ["RangeNumber"],
        ["EOF"] = ["FileNumber"],
        ["LOF"] = ["FileNumber"],
        ["Loc"] = ["FileNumber"],
        ["Seek"] = ["FileNumber"],
        ["Input"] = ["Number", "FileNumber"],
        // Arrays
        ["LBound"] = ["ArrayName", "Dimension"],
        ["UBound"] = ["ArrayName", "Dimension"],
    };

    private static void Add(string name, VbaType returnType, int min, int max, params string[] templates) =>
        Functions[name] = new IntrinsicSymbol(name, returnType, returnType, min, max, templates);

    private static void Add(string name, VbaType returnType, VbaType resultType, int min, int max, params string[] templates) =>
        Functions[name] = new IntrinsicSymbol(name, returnType, resultType, min, max, templates);

    private static void AddSub(string name, int min, int max, params string[] templates) =>
        Functions[name] = new IntrinsicSymbol(name, VbaType.Variant, VbaType.Variant, min, max, templates) { IsSub = true };

    private static void AddConstants()
    {
        // Strings (MS-VBAL 6.1.1 Constants module).
        Constant("vbCrLf", "\r\n");
        Constant("vbCr", "\r");
        Constant("vbLf", "\n");
        Constant("vbNewLine", "\r\n");
        Constant("vbTab", "\t");
        Constant("vbBack", "\b");
        Constant("vbFormFeed", "\f");
        Constant("vbVerticalTab", "\v");
        Constant("vbNullChar", "\0");
        Constant("vbNullString", "");
        Constant("vbObjectError", -2147221504);

        // VbVarType.
        Constant("vbEmpty", 0);
        Constant("vbNull", 1);
        Constant("vbInteger", 2);
        Constant("vbLong", 3);
        Constant("vbSingle", 4);
        Constant("vbDouble", 5);
        Constant("vbCurrency", 6);
        Constant("vbDate", 7);
        Constant("vbString", 8);
        Constant("vbObject", 9);
        Constant("vbError", 10);
        Constant("vbBoolean", 11);
        Constant("vbVariant", 12);
        Constant("vbDataObject", 13);
        Constant("vbDecimal", 14);
        Constant("vbByte", 17);
        Constant("vbLongLong", 20);
        Constant("vbUserDefinedType", 36);
        Constant("vbArray", 8192);

        // VbTriState, VbCompareMethod, VbDayOfWeek, VbFirstWeekOfYear, VbDateTimeFormat, VbStrConv.
        Constant("vbTrue", -1);
        Constant("vbFalse", 0);
        Constant("vbUseDefault", -2);
        Constant("vbBinaryCompare", 0);
        Constant("vbTextCompare", 1);
        Constant("vbDatabaseCompare", 2);
        Constant("vbSunday", 1);
        Constant("vbMonday", 2);
        Constant("vbTuesday", 3);
        Constant("vbWednesday", 4);
        Constant("vbThursday", 5);
        Constant("vbFriday", 6);
        Constant("vbSaturday", 7);
        Constant("vbUseSystemDayOfWeek", 0);
        Constant("vbUseSystem", 0);
        Constant("vbFirstJan1", 1);
        Constant("vbFirstFourDays", 2);
        Constant("vbFirstFullWeek", 3);
        Constant("vbGeneralDate", 0);
        Constant("vbLongDate", 1);
        Constant("vbShortDate", 2);
        Constant("vbLongTime", 3);
        Constant("vbShortTime", 4);
        Constant("vbUpperCase", 1);
        Constant("vbLowerCase", 2);
        Constant("vbProperCase", 3);
        Constant("vbWide", 4);
        Constant("vbNarrow", 8);
        Constant("vbKatakana", 16);
        Constant("vbHiragana", 32);
        Constant("vbUnicode", 64);
        Constant("vbFromUnicode", 128);

        // ColorConstants.
        Constant("vbBlack", 0);
        Constant("vbRed", 255);
        Constant("vbGreen", 65280);
        Constant("vbYellow", 65535);
        Constant("vbBlue", 16711680);
        Constant("vbMagenta", 16711935);
        Constant("vbCyan", 16776960);
        Constant("vbWhite", 16777215);

        // VbMsgBoxStyle and VbMsgBoxResult.
        Constant("vbOKOnly", 0);
        Constant("vbOKCancel", 1);
        Constant("vbAbortRetryIgnore", 2);
        Constant("vbYesNoCancel", 3);
        Constant("vbYesNo", 4);
        Constant("vbRetryCancel", 5);
        Constant("vbCritical", 16);
        Constant("vbQuestion", 32);
        Constant("vbExclamation", 48);
        Constant("vbInformation", 64);
        Constant("vbDefaultButton1", 0);
        Constant("vbDefaultButton2", 256);
        Constant("vbDefaultButton3", 512);
        Constant("vbDefaultButton4", 768);
        Constant("vbApplicationModal", 0);
        Constant("vbSystemModal", 4096);
        Constant("vbMsgBoxHelpButton", 16384);
        Constant("vbMsgBoxSetForeground", 65536);
        Constant("vbMsgBoxRight", 524288);
        Constant("vbMsgBoxRtlReading", 1048576);
        Constant("vbOK", 1);
        Constant("vbCancel", 2);
        Constant("vbAbort", 3);
        Constant("vbRetry", 4);
        Constant("vbIgnore", 5);
        Constant("vbYes", 6);
        Constant("vbNo", 7);

        // VbFileAttribute, VbAppWinStyle.
        Constant("vbNormal", 0);
        Constant("vbReadOnly", 1);
        Constant("vbHidden", 2);
        Constant("vbSystem", 4);
        Constant("vbVolume", 8);
        Constant("vbDirectory", 16);
        Constant("vbArchive", 32);
        Constant("vbAlias", 64);
        Constant("vbHide", 0);
        Constant("vbNormalFocus", 1);
        Constant("vbMinimizedFocus", 2);
        Constant("vbMaximizedFocus", 3);
        Constant("vbNormalNoFocus", 4);
        Constant("vbMinimizedNoFocus", 6);

        // KeyCodeConstants are Integer (Constants golden).
        Constant("vbKeyLButton", (short)1);
        Constant("vbKeyReturn", (short)13);
        Constant("vbKeyEscape", (short)27);
        Constant("vbKeySpace", (short)32);
        Constant("vbKey0", (short)48);
        Constant("vbKeyA", (short)65);
        Constant("vbKeyNumpad0", (short)96);
        Constant("vbKeyF1", (short)112);

        // VbCallType, VbCalendar, VbIMEStatus.
        Constant("vbGet", 2);
        Constant("vbLet", 4);
        Constant("vbSet", 8);
        Constant("vbMethod", 1);
        Constant("vbCalGreg", 0);
        Constant("vbCalHijri", 1);
        Constant("vbIMEModeNoControl", 0);
    }

    private static void Constant(string name, Variant value) =>
        Constants[name] = new ConstantSymbol(name, VbaType.FromVarType(value.Type), value) { IsPublic = true };

    /// <summary>A String constant of the library lives as long as the compiler does, so its BSTR is a literal rather than a temporary of one compilation (ARCHITECTURE.md D20).</summary>
    private static void Constant(string name, string text) => Constant(name, Variant.ViewString(VbaString.Literal(text)));
}
