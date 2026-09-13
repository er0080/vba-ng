using System.Globalization;
using System.Reflection;
using System.Runtime.CompilerServices;

namespace VbaNg.Runtime.Hosting;

/// <summary>
/// Excel's <c>Application.Run</c> as generated code calls it (ARCHITECTURE.md D23): a macro name
/// that names a procedure of the calling project runs in-process, in the loaded project the caller
/// belongs to and so with the caller's module state; any other name goes to Excel. The names and
/// arguments behave as Excel showed them (docs/vba-quirks.md, "Host, Application.Run"). Each
/// standard and document module carries a Run table the compiler emits: <c>__RunId</c> maps a
/// procedure name to an id, <c>__Run</c> calls the procedure with a late-bound argument list.
/// </summary>
public static class ApplicationRun
{
    /// <summary>The error Excel raises for a macro it cannot run.</summary>
    private const int CannotRunMacro = 1004;

    private const string ExcelSource = "Microsoft Excel";

    private static readonly ConditionalWeakTable<Assembly, RunTables> Tables = new();

    private delegate Variant RunProcedure(int id, ReadOnlySpan<Variant> arguments);

    /// <summary>
    /// <c>Run</c> on Excel's Application or Global (<paramref name="dispId"/> on <paramref name="target"/>),
    /// called from a module of <paramref name="caller"/>'s project: the macro name comes first, then
    /// the arguments.
    /// </summary>
    public static Variant Invoke(Type caller, IDispatchObject? target, int dispId, ReadOnlySpan<Variant> arguments)
    {
        ArgumentNullException.ThrowIfNull(caller);
        if (arguments.Length > 0 && arguments[0].Type == VarType.String && Resolve(caller.Assembly, Coerce.ToString(arguments[0])) is { } found)
        {
            // Excel drops the arguments left out at the end, so they count neither toward the procedure's parameters nor as Missing.
            var passed = arguments[1..];
            var count = passed.Length;
            while (count > 0 && passed[count - 1].IsMissing)
            {
                count--;
            }

            return found.Table.Run(found.Id, passed[..count]);
        }

        return EarlyBound.Invoke(target, dispId, InvokeKind.Method, arguments);
    }

    /// <summary>Error 450 when the call passes more arguments than the procedure takes; the caller can trap it.</summary>
    public static void Arity(ReadOnlySpan<Variant> arguments, int count)
    {
        if (arguments.Length > count)
        {
            throw new VbaException(VbaErrors.WrongNumberOfArguments);
        }
    }

    /// <summary>Whether an exception leaving the procedure Run called is a run-time error, which the caller's handlers must not see.</summary>
    public static bool Escapes(Exception exception) => VbaException.From(exception) is not null;

    /// <summary>The run-time error, as one no handler traps: VBA shows its run-time error dialog rather than raise it to the caller.</summary>
    public static UnhandledErrorException Unhandled(Exception exception)
    {
        ArgumentNullException.ThrowIfNull(exception);
        return new UnhandledErrorException(VbaException.From(exception) ?? throw new ArgumentException("Not a run-time error.", nameof(exception)), exception);
    }

    /// <summary>The procedure a macro name names in the project, or null when it names none and Excel is to run it.</summary>
    private static (RunTable Table, int Id)? Resolve(Assembly project, string macro)
    {
        var tables = Tables.GetValue(project, RunTables.Of);
        var name = macro;
        var bang = name.LastIndexOf('!');
        if (bang >= 0)
        {
            if (!tables.IsOwnWorkbook(Unquote(name[..bang])))
            {
                return null;
            }

            name = name[(bang + 1)..];
        }

        var dot = name.IndexOf('.', StringComparison.Ordinal);
        if (dot >= 0)
        {
            // Module.Procedure: a standard or document module; a class module's methods are not macros.
            return tables.Find(name[..dot]) is { } table && table.IdOf(name[(dot + 1)..]) is var id and >= 0 ? (table, id) : null;
        }

        (RunTable Table, int Id)? found = null;
        foreach (var table in tables.Standard)
        {
            var id = table.IdOf(name);
            if (id < 0)
            {
                continue;
            }

            if (found is not null)
            {
                // A bare name two standard modules define is a macro Excel cannot run.
                throw new VbaException(CannotRunMacro, string.Create(CultureInfo.InvariantCulture, $"Cannot run the macro '{macro}'. The macro may not be available in this workbook or all macros may be disabled."), ExcelSource);
            }

            found = (table, id);
        }

        return found;
    }

    /// <summary>A workbook name as a macro name writes it, <c>'Book 1.xlsm'</c> or <c>Book.xlsm</c>.</summary>
    private static string Unquote(string book) =>
        book.Length >= 2 && book[0] == '\'' && book[^1] == '\'' ? book[1..^1].Replace("''", "'", StringComparison.Ordinal) : book;

    /// <summary>One module's Run table.</summary>
    private sealed record RunTable(string Name, Func<string, int> IdOf, RunProcedure Run);

    /// <summary>The Run tables of a project's standard and document modules, found once per loaded assembly.</summary>
    private sealed class RunTables
    {
        private readonly Dictionary<string, RunTable> modules = new(StringComparer.OrdinalIgnoreCase);
        private readonly List<RunTable> standard = [];
        private FieldInfo? workbook;

        public IReadOnlyList<RunTable> Standard => standard;

        public static RunTables Of(Assembly project)
        {
            var tables = new RunTables();
            foreach (var type in project.GetTypes())
            {
                var idOf = type.GetMethod("__RunId", BindingFlags.Public | BindingFlags.Static, [typeof(string)]);
                var run = type.GetMethod("__Run", BindingFlags.Public | BindingFlags.Static, [typeof(int), typeof(ReadOnlySpan<Variant>)]);
                if (idOf is null || run is null)
                {
                    continue;
                }

                var table = new RunTable(type.Name, idOf.CreateDelegate<Func<string, int>>(), run.CreateDelegate<RunProcedure>());
                tables.modules[type.Name] = table;
                var document = type.GetCustomAttribute<VbaDocumentModuleAttribute>();
                if (document is null)
                {
                    tables.standard.Add(table);
                }
                else if (document.Kind == "Workbook")
                {
                    tables.workbook = type.GetField(VbaDocumentModuleAttribute.MeField, BindingFlags.Public | BindingFlags.Static);
                }
            }

            return tables;
        }

        public RunTable? Find(string module) => modules.GetValueOrDefault(module);

        /// <summary>Whether a workbook a macro name qualifies it with is the project's own; a project bound to no workbook takes every name as its own.</summary>
        public bool IsOwnWorkbook(string book)
        {
            if (workbook?.GetValue(null) is not IDispatchObject me)
            {
                return true;
            }

            var name = Coerce.ToString(me.Invoke(me.GetDispId("Name"), InvokeKind.PropertyGet, []));
            return string.Equals(name, book, StringComparison.OrdinalIgnoreCase);
        }
    }
}
