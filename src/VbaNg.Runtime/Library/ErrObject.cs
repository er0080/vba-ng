namespace VbaNg.Runtime.Library;

/// <summary>
/// The Err object (MS-VBAL 6.1.3.1 ErrObject): the properties of the most recent runtime error
/// and the Raise and Clear methods. One instance serves a running project; compiled error
/// handlers set it from the <see cref="VbaException"/> they catch.
/// </summary>
public sealed class ErrObject : RuntimeObject, IVbaObject
{
    public string TypeName => "ErrObject";

    /// <summary>A COM caller reaches Err's members: Number, the default member, at 0, the others at dispids of the runtime's own (ROADMAP.md M7 F).</summary>
    internal override IDispatchObject ComMembers =>
        new LateBoundMembers(this, TypeName, [(0, "Number"), (1, "Source"), (2, "Description"), (3, "HelpFile"), (4, "HelpContext"), (5, "Raise"), (6, "Clear"), (7, "LastDllError")]);

    /// <summary>Err has a COM identity like any object (ARCHITECTURE.md section 5, "Objects"); it lives as long as its thread, so a last reference going changes nothing.</summary>
    protected override void OnLastRelease()
    {
    }

    public int Number { get; set; }

    /// <summary>Err.Description as generated code reads and writes it: a String of the current statement (ARCHITECTURE.md D20); the runtime works with <see cref="DescriptionText"/>.</summary>
    public VbaString Description
    {
        get => VbaString.Temporary(DescriptionText);
        set => DescriptionText = value.ToString();
    }

    public VbaString Source
    {
        get => VbaString.Temporary(SourceText);
        set => SourceText = value.ToString();
    }

    public VbaString HelpFile
    {
        get => VbaString.Temporary(HelpFileText);
        set => HelpFileText = value.ToString();
    }

    public string DescriptionText { get; set; } = string.Empty;

    public string SourceText { get; set; } = string.Empty;

    public string HelpFileText { get; set; } = string.Empty;

    public int HelpContext { get; set; }

    public int LastDllError { get; set; }

    /// <summary>The line number of the last executed numbered line, which Erl reports (MS-VBAL 5.4.1.1 line numbers).</summary>
    public int Line { get; set; }

    /// <summary>Records an error that a handler caught.</summary>
    public void Set(VbaException error)
    {
        ArgumentNullException.ThrowIfNull(error);
        Number = error.Number;
        DescriptionText = error.Description;
        SourceText = error.Source ?? VbaErrors.DefaultSource;
        HelpFileText = error.HelpFile;
        HelpContext = error.HelpContext;
    }

    public void Clear()
    {
        Number = 0;
        DescriptionText = string.Empty;
        SourceText = string.Empty;
        HelpFileText = string.Empty;
        HelpContext = 0;
    }

    /// <summary>
    /// Err.Raise Number, [Source], [Description], [HelpFile], [HelpContext]: 0 and numbers above
    /// 65535 raise error 5, Null arguments raise 94, an omitted description is the number's message
    /// and an omitted source is the project name (Errors goldens).
    /// </summary>
    public VbaException Raise(Variant number, Variant source = default, Variant description = default, Variant helpFile = default, Variant helpContext = default)
    {
        var code = Coerce.ToInt32(number);
        if (code == 0 || code > VbaErrors.MaxNumber)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        // While Err already holds an error, omitted arguments keep its current properties: raising 7 inside a
        // handler for error 5 reports "Invalid procedure call or argument" (Procedures golden).
        var active = Number != 0;
        var text = IsOmitted(description) ? (active && DescriptionText.Length > 0 ? DescriptionText : null) : Coerce.ToString(description);
        var origin = IsOmitted(source) ? (active && SourceText.Length > 0 ? SourceText : null) : Coerce.ToString(source);
        var file = IsOmitted(helpFile) ? (active && HelpFileText.Length > 0 ? HelpFileText : null) : Coerce.ToString(helpFile);
        int? context = IsOmitted(helpContext) ? (active && HelpContext != 0 ? HelpContext : null) : Coerce.ToInt32(helpContext);
        var error = new VbaException(code, text, origin, file, context);
        Set(error);
        return error;
    }

    private static bool IsOmitted(in Variant value) => value.Type == VarType.Empty || value.IsMissing;
}
