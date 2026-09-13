namespace VbaNg.Runtime.Hosting;

/// <summary>
/// Marks the compiled form of a document module (ARCHITECTURE.md section 3): a class module with
/// a predeclared instance that the host binds to a workbook object by CodeName. <see cref="Kind"/>
/// is Workbook, Worksheet, or Chart; the class carries a static <c>Me</c> field the host sets.
/// </summary>
[AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = false)]
public sealed class VbaDocumentModuleAttribute(string kind) : Attribute
{
    public string Kind { get; } = kind ?? throw new ArgumentNullException(nameof(kind));

    /// <summary>The name of the static field holding the bound object.</summary>
    public const string MeField = "Me";
}

/// <summary>
/// Marks an event procedure (<c>Worksheet_Change</c>, <c>CommandButton1_Click</c>): the host
/// finds the source object by name, the event by name on its source interface, and advises a
/// sink that calls the method (ARCHITECTURE.md section 6, "Events"). The source is Workbook,
/// Worksheet, or Chart for the module's own object, otherwise the name of a control on it.
/// </summary>
[AttributeUsage(AttributeTargets.Method, AllowMultiple = false, Inherited = false)]
public sealed class VbaEventHandlerAttribute(string source, string eventName) : Attribute
{
    public string Source { get; } = source ?? throw new ArgumentNullException(nameof(source));

    public string EventName { get; } = eventName ?? throw new ArgumentNullException(nameof(eventName));
}
