using System.Globalization;
using System.Runtime.InteropServices;

namespace VbaNg.Runtime;

/// <summary>
/// A VBA run-time error: the number, description, source, and help context that <c>Err</c> exposes
/// (MS-VBAL 6.1.3.1 ErrObject). Every runtime error the library raises is one of these, with the
/// number and message text VBA uses (CLAUDE.md R5).
/// </summary>
public sealed class VbaException : Exception
{
    private string source;

    /// <summary>
    /// An omitted description is the number's message, an omitted source is the project name, and
    /// an omitted help file is VBA's own help file with context 1000000 plus the number, which is
    /// what Err reports for errors the runtime raises (Errors golden).
    /// </summary>
    public VbaException(int number, string? description = null, string? source = null, string? helpFile = null, int? helpContext = null)
        : base(description ?? VbaErrors.Message(number))
    {
        Number = number;
        Description = description ?? VbaErrors.Message(number);
        this.source = source ?? VbaErrors.DefaultSource;
        HelpFile = helpFile ?? VbaErrors.HelpFile;
        HelpContext = helpContext ?? (helpFile is null && number > 0 ? VbaErrors.HelpContextBase + number : 0);
    }

    public VbaException()
        : this(VbaErrors.ApplicationDefined)
    {
    }

    public VbaException(string message)
        : this(VbaErrors.ApplicationDefined, message)
    {
    }

    public VbaException(string message, Exception innerException)
        : base(message, innerException)
    {
        Number = VbaErrors.ApplicationDefined;
        Description = message;
        source = VbaErrors.DefaultSource;
        HelpFile = string.Empty;
    }

    public int Number { get; }

    /// <summary>The VBA error an exception stands for: a VbaException itself, or the runtime error a .NET exception maps to (MS-VBAL 5.4.4 run-time errors); null for exceptions VBA code cannot handle.</summary>
    public static VbaException? From(Exception exception)
    {
        ArgumentNullException.ThrowIfNull(exception);
        return exception switch
        {
            VbaException error => error,
            DivideByZeroException => VbaErrors.DivisionByZero(),
            OverflowException => VbaErrors.Overflow(),
            IndexOutOfRangeException or ArgumentOutOfRangeException => VbaErrors.SubscriptOutOfRange(),
            NullReferenceException => VbaErrors.ObjectVariableNotSet(),
            InvalidCastException => VbaErrors.TypeMismatch(),
            OutOfMemoryException => new VbaException(VbaErrors.OutOfMemory),
            _ => null,
        };
    }

    public string Description { get; }

    /// <summary><c>Err.Source</c>: the project name unless <c>Err.Raise</c> supplied one.</summary>
    public override string? Source
    {
        get => source;
        set => source = value ?? VbaErrors.DefaultSource;
    }

    public string HelpFile { get; }

    public int HelpContext { get; }
}

/// <summary>
/// VBA error numbers and their message text, recorded from Excel (tests/VbaNg.Golden/Goldens/Errors.json,
/// CLAUDE.md R5). Numbers without a message of their own describe themselves as
/// "Application-defined or object-defined error"; negative numbers are HRESULTs and describe
/// themselves as "Automation error" followed by the system's text for the code, when it has one.
/// </summary>
public static class VbaErrors
{
    public const int ReturnWithoutGoSub = 3;
    public const int InvalidProcedureCallNumber = 5;
    public const int OverflowNumber = 6;
    public const int OutOfMemory = 7;
    public const int SubscriptOutOfRangeNumber = 9;
    public const int ArrayFixedOrLocked = 10;
    public const int DivisionByZeroNumber = 11;
    public const int TypeMismatchNumber = 13;
    public const int OutOfStringSpace = 14;
    public const int ResumeWithoutError = 20;
    public const int UserInterrupt = 18;
    public const int ObjectVariableNotSetNumber = 91;
    public const int ForLoopNotInitialized = 92;
    public const int InvalidPatternStringNumber = 93;
    public const int InvalidUseOfNullNumber = 94;
    public const int ObjectRequired = 424;
    public const int CannotCreateObject = 429;
    public const int ClassDoesNotSupportAutomation = 430;
    public const int FileOrClassNameNotFound = 432;
    public const int ObjectDoesNotSupportMember = 438;
    public const int NoNamedArguments = 446;
    public const int UnsupportedLocale = 447;
    public const int NamedArgumentNotFound = 448;
    public const int ArgumentNotOptional = 449;
    public const int WrongNumberOfArguments = 450;
    public const int DuplicateCollectionKey = 457;
    public const int ApplicationDefined = 1000;

    /// <summary>The largest number <c>Err.Raise</c> and <c>Error$</c> accept.</summary>
    public const int MaxNumber = 65535;

    private const string ApplicationDefinedMessage = "Application-defined or object-defined error";

    /// <summary>
    /// <c>Err.Source</c> of errors the runtime raises: the VBA project name. The host sets it when a
    /// project loads; "VBAProject" is what a fresh Excel project is called.
    /// </summary>
    public static string DefaultSource { get; set; } = "VBAProject";

    /// <summary><c>Err.HelpFile</c> of errors the runtime raises: the VBA language reference help file, as Excel reports it.</summary>
    public static string HelpFile { get; set; } = @"C:\Program Files\Common Files\Microsoft Shared\VBA\VBA7.1\1033\VbLR6.chm";

    /// <summary><c>Err.HelpContext</c> of a runtime error is this plus the error number.</summary>
    public const int HelpContextBase = 1000000;

    public static VbaException InvalidProcedureCall() => new(InvalidProcedureCallNumber);

    public static VbaException Overflow() => new(OverflowNumber);

    public static VbaException SubscriptOutOfRange() => new(SubscriptOutOfRangeNumber);

    public static VbaException DivisionByZero() => new(DivisionByZeroNumber);

    public static VbaException TypeMismatch() => new(TypeMismatchNumber);

    public static VbaException ObjectVariableNotSet() => new(ObjectVariableNotSetNumber);

    public static VbaException InvalidPatternString() => new(InvalidPatternStringNumber);

    public static VbaException InvalidUseOfNull() => new(InvalidUseOfNullNumber);

    /// <summary>The message VBA gives error <paramref name="number"/>, as <c>Error$(number)</c> returns it.</summary>
    public static string Message(int number)
    {
        if (number < 0)
        {
            return AutomationErrorMessage(number);
        }

        return number switch
        {
            3 => "Return without GoSub",
            5 => "Invalid procedure call or argument",
            6 => "Overflow",
            7 => "Out of memory",
            9 => "Subscript out of range",
            10 => "This array is fixed or temporarily locked",
            11 => "Division by zero",
            13 => "Type mismatch",
            14 => "Out of string space",
            16 => "Expression too complex",
            17 => "Can't perform requested operation",
            18 => "User interrupt occurred",
            20 => "Resume without error",
            28 => "Out of stack space",
            35 => "Sub or Function not defined",
            47 => "Too many DLL application clients",
            48 => "Error in loading DLL",
            49 => "Bad DLL calling convention",
            51 => "Internal error",
            52 => "Bad file name or number",
            53 => "File not found",
            54 => "Bad file mode",
            55 => "File already open",
            57 => "Device I/O error",
            58 => "File already exists",
            59 => "Bad record length",
            61 => "Disk full",
            62 => "Input past end of file",
            63 => "Bad record number",
            67 => "Too many files",
            68 => "Device unavailable",
            70 => "Permission denied",
            71 => "Disk not ready",
            74 => "Can't rename with different drive",
            75 => "Path/File access error",
            76 => "Path not found",
            91 => "Object variable or With block variable not set",
            92 => "For loop not initialized",
            93 => "Invalid pattern string",
            94 => "Invalid use of Null",
            96 => "Unable to sink events of object because the object is already firing events to the maximum number of event receivers that it supports",
            97 => "Can not call friend function on object which is not an instance of defining class",
            98 => "A property or method call cannot include a reference to a private object, either as an argument or as a return value",
            321 => "Invalid file format",
            322 => "Can't create necessary temporary file",
            325 => "Invalid format in resource file",
            380 => "Invalid property value",
            381 => "Invalid property array index",
            382 => "Set not supported at runtime",
            383 => "Set not supported (read-only property)",
            385 => "Need property array index",
            387 => "Set not permitted",
            393 => "Get not supported at runtime",
            394 => "Get not supported (write-only property)",
            422 => "Property not found",
            423 => "Property or method not found",
            424 => "Object required",
            429 => "ActiveX component can't create object",
            430 => "Class does not support Automation or does not support expected interface",
            432 => "File name or class name not found during Automation operation",
            438 => "Object doesn't support this property or method",
            440 => "Automation error",
            442 => "Connection to type library or object library for remote process has been lost. Press OK for dialog to remove reference.",
            443 => "Automation object does not have a default value",
            445 => "Object doesn't support this action",
            446 => "Object doesn't support named arguments",
            447 => "Object doesn't support current locale setting",
            448 => "Named argument not found",
            449 => "Argument not optional",
            450 => "Wrong number of arguments or invalid property assignment",
            451 => "Property let procedure not defined and property get procedure did not return an object",
            452 => "Invalid ordinal",
            453 => "Specified DLL function not found",
            454 => "Code resource not found",
            455 => "Code resource lock error",
            457 => "This key is already associated with an element of this collection",
            458 => "Variable uses an Automation type not supported in Visual Basic",
            459 => "Object or class does not support the set of events",
            460 => "Invalid clipboard format",
            461 => "Method or data member not found",
            462 => "The remote server machine does not exist or is unavailable",
            463 => "Class not registered on local machine",
            481 => "Invalid picture",
            482 => "Printer error",
            735 => "Can't save file to TEMP",
            744 => "Search text not found",
            746 => "Replacements too long",
            _ => ApplicationDefinedMessage,
        };
    }

    /// <summary>
    /// "Automation error", then on a new line the system's text for the HRESULT when it has one,
    /// with its trailing line break turned into a space, as Excel reports it (Errors golden).
    /// </summary>
    private static string AutomationErrorMessage(int hresult)
    {
        var system = SystemMessage(hresult);
        return system is null ? "Automation error" : "Automation error\n" + system;
    }

    private static unsafe string? SystemMessage(int hresult)
    {
        if (!OperatingSystem.IsWindows())
        {
            return null;
        }

        const uint FormatMessageFromSystem = 0x00001000;
        const uint FormatMessageIgnoreInserts = 0x00000200;
        var buffer = stackalloc char[1024];
        var length = FormatMessageW(FormatMessageFromSystem | FormatMessageIgnoreInserts, 0, (uint)hresult, 0, buffer, 1024, 0);
        if (length == 0)
        {
            return null;
        }

        return new string(buffer, 0, (int)length).Replace("\r\n", " ", StringComparison.Ordinal);
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
    private static extern unsafe uint FormatMessageW(uint flags, nint source, uint messageId, uint languageId, char* buffer, uint size, nint arguments);

    /// <summary><c>Error$(number)</c>: the message; 0 gives an empty string, a number above 65535 raises error 6, a negative number is application-defined (Information golden).</summary>
    public static string ErrorText(long number)
    {
        if (number > MaxNumber || number < int.MinValue)
        {
            throw Overflow();
        }

        if (number == 0)
        {
            return string.Empty;
        }

        return number < 0 ? ApplicationDefinedMessage : Message((int)number);
    }

    internal static string Invariant(FormattableString text) => text.ToString(CultureInfo.InvariantCulture);
}
