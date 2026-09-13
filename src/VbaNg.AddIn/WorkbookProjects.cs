using System.Globalization;
using System.Linq.Expressions;
using System.Reflection;

using ExcelDna.Integration;

using VbaNg.Interop;
using VbaNg.Runtime;
using VbaNg.Runtime.Hosting;

namespace VbaNg.AddIn;

/// <summary>
/// The workbook lifecycle of ARCHITECTURE.md section 7: on every WorkbookOpen (and for workbooks
/// already open when the add-in loads, D16) the project folder next to the workbook is found
/// (D4), built when stale, loaded, its document modules bound to the workbook's objects by
/// CodeName, their event procedures advised, Workbook_Open and Auto_Open run, its public Subs
/// registered as commands so OnAction and Application.Run find them, its public Functions
/// registered as UDFs so formulas find them, and its folder watched for hot reload; on
/// WorkbookBeforeClose the close handlers run and the binding, the event sinks, and the watcher
/// go. The names stay registered: Excel-DNA offers no way to unregister them, so they outlive the
/// workbook and are registered again after a rebuild.
/// </summary>
internal sealed class WorkbookProjects : IDisposable
{
    private readonly ProjectHost projects;
    private readonly Action<string> log;
    private readonly Dictionary<string, BoundWorkbook> bound = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, string> registered = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, ProjectWatcher> watchers = new(StringComparer.OrdinalIgnoreCase);

    // The folder of the project each registered procedure belongs to, for its log; weak, so a reloaded project's old assembly can go.
    private readonly System.Runtime.CompilerServices.ConditionalWeakTable<Assembly, string> projectDirs = [];
    private readonly VbaNg.Runtime.TypeLibraries.ComLibrary? excelLibrary;
    private ComObject? application;
    private EventSink? applicationSink;

    public WorkbookProjects(ProjectHost projects, Action<string> log)
    {
        this.projects = projects;
        this.log = log;

        // Excel's objects expose no class information, so their event interfaces come from the type library model.
        excelLibrary = TypeLibraryCache.Resolve("Excel", null, null);
        if (excelLibrary is null)
        {
            log("The Excel type library could not be read; workbook events will not be connected.");
        }
    }

    public IEnumerable<string> Summary => bound.Values.Select(b => b.Name + " -> " + b.Project.ProjectDir);

    /// <summary>How many times a watched project was rebuilt and swapped in.</summary>
    public int Reloads { get; private set; }

    /// <summary>Advises the application's events and binds every workbook that is already open.</summary>
    public void Start(ComObject excel)
    {
        // The wrappers this scan creates are temporaries of the call; what a binding keeps, it retains (ARCHITECTURE.md D18).
        using var frame = ObjectRefs.Frame();
        application = excel;
        var source = EventSource.Describe(excel) ?? (excelLibrary is null ? null : EventSource.FromLibrary(excelLibrary, "Application"));
        if (source is null)
        {
            log("Application exposes no event source; workbooks will not bind automatically.");
            return;
        }

        var open = source.Events.GetValueOrDefault("WorkbookOpen")?.DispId;
        var beforeClose = source.Events.GetValueOrDefault("WorkbookBeforeClose")?.DispId;
        applicationSink = new EventSink(source.InterfaceId, (dispId, raw) =>
        {
            var arguments = EventBinder.InDeclaredOrder(source, raw);
            if (dispId == open && arguments.Length > 0 && arguments[0].IsObject && arguments[0].AsObject() is ComObject workbook)
            {
                OnWorkbookOpen(workbook);
            }
            else if (dispId == beforeClose && arguments.Length > 0 && arguments[0].IsObject && arguments[0].AsObject() is ComObject closing)
            {
                OnWorkbookBeforeClose(closing);
            }
        })
        {
            Error = ex => log("Workbook lifecycle error: " + EventBinder.Describe(ex)),
        };
        excel.Advise(applicationSink);

        var workbooks = excel.Invoke(excel.GetDispId("Workbooks"), InvokeKind.PropertyGet, []);
        if (workbooks.AsObject() is ComObject collection)
        {
            using var enumerator = collection.Enumerate();
            while (enumerator.MoveNext())
            {
                if (enumerator.Current.AsObject() is ComObject workbook)
                {
                    OnWorkbookOpen(workbook);
                }
            }
        }
    }

    /// <summary>A project reloaded by a host command: rebinds the workbook it belongs to without rerunning Workbook_Open (D16).</summary>
    public void Refresh(LoadedProject project)
    {
        using var frame = ObjectRefs.Frame();
        foreach (var entry in bound.Values.ToList())
        {
            if (string.Equals(entry.Project.ProjectDir, project.ProjectDir, StringComparison.OrdinalIgnoreCase) && !ReferenceEquals(entry.Project, project))
            {
                entry.Unbind();
                entry.Bind(project);
                Register(project);
                log($"Rebound {entry.Name} to the rebuilt project.");
            }
        }
    }

    public void Dispose()
    {
        foreach (var entry in bound.Values)
        {
            entry.Dispose();
        }

        bound.Clear();
        foreach (var watcher in watchers.Values)
        {
            watcher.Dispose();
        }

        watchers.Clear();
        applicationSink?.Dispose();
        applicationSink = null;
    }

    private void OnWorkbookOpen(ComObject workbook)
    {
        var fullName = Coerce.ToString(workbook.Invoke(workbook.GetDispId("FullName"), InvokeKind.PropertyGet, []));
        if (bound.ContainsKey(fullName))
        {
            return;
        }

        var projectDir = Path.ChangeExtension(fullName, ProjectPaths.FolderSuffix);
        if (!Directory.Exists(projectDir))
        {
            return;
        }

        if (HasVBProject(workbook))
        {
            // A workbook that keeps its VBA project would run its own macros as well as the project's (ARCHITECTURE.md D17).
            var name = Path.GetFileName(fullName);
            ExcelHostServices.Instance.Notify($"{name} still has a VBA project, so vba-ng does not bind {Path.GetFileName(projectDir)} to it: both would run. Save a macro-free copy with \"vbang import --to-xlsx {name}\" and open that one.");
            return;
        }

        var outcome = ProjectBuilder.EnsureBuilt(projectDir);
        if (outcome == BuildOutcome.Failed)
        {
            log($"Build failed for {projectDir}:{Environment.NewLine}{ProjectBuilder.LastOutput}");
            return;
        }

        if (outcome == BuildOutcome.NoCli)
        {
            log(ProjectBuilder.LastOutput);
        }

        LoadedProject project;
        try
        {
            project = projects.Load(projectDir);
        }
        catch (Exception ex) when (ex is IOException or BadImageFormatException or FileNotFoundException)
        {
            log($"Could not load {projectDir}: {ex.Message}");
            return;
        }

        var entry = new BoundWorkbook(Path.GetFileName(fullName), workbook, log, excelLibrary);
        entry.Bind(project);
        bound[fullName] = entry;
        Register(project);
        Watch(projectDir);
        entry.RunOpenHandlers();
        log($"Bound {entry.Name} to {project.Name} ({entry.DocumentCount} document module(s), {entry.SinkCount} event source(s)).");
    }

    /// <summary>Workbook.HasVBProject; false when this Excel cannot tell.</summary>
    private static bool HasVBProject(ComObject workbook)
    {
        try
        {
            return Coerce.ToBoolean(workbook.Invoke(workbook.GetDispId("HasVBProject"), InvokeKind.PropertyGet, []));
        }
        catch (Exception ex) when (ex is VbaException or System.Runtime.InteropServices.COMException)
        {
            return false;
        }
    }

    private void OnWorkbookBeforeClose(ComObject workbook)
    {
        var fullName = Coerce.ToString(workbook.Invoke(workbook.GetDispId("FullName"), InvokeKind.PropertyGet, []));
        if (!bound.Remove(fullName, out var entry))
        {
            return;
        }

        entry.RunCloseHandlers();
        entry.Dispose();
        Unwatch(entry.Project.ProjectDir);
        log($"Unbound {entry.Name}.");
    }

    /// <summary>Starts hot reload for a bound project folder (ARCHITECTURE.md section 7).</summary>
    private void Watch(string projectDir)
    {
        if (watchers.ContainsKey(projectDir))
        {
            return;
        }

        try
        {
            watchers[projectDir] = new ProjectWatcher(projectDir, log, () => Reload(projectDir));
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException)
        {
            log($"Could not watch {projectDir} for changes: {ex.Message}");
        }
    }

    private void Unwatch(string projectDir)
    {
        if (watchers.Remove(projectDir, out var watcher))
        {
            watcher.Dispose();
        }
    }

    /// <summary>A rebuilt project, on Excel's thread: loads the new assembly and swaps it into its workbook.</summary>
    private void Reload(string projectDir)
    {
        LoadedProject project;
        try
        {
            project = projects.Load(projectDir);
        }
        catch (Exception ex) when (ex is IOException or BadImageFormatException or FileNotFoundException)
        {
            log($"Hot reload could not load {projectDir}: {ex.Message}");
            return;
        }

        Refresh(project);
        Reloads++;
        log($"Hot reload {Reloads.ToString(CultureInfo.InvariantCulture)}: {project.Name} rebuilt and swapped in.");
    }

    /// <summary>
    /// Registers the project's public procedures with Excel (ARCHITECTURE.md section 7): every
    /// public parameterless Sub of a standard module as a command and every public Function as a
    /// UDF, each under its own name and as Module.Name. A name another project already holds is
    /// registered as Project.Name instead, with a warning; a project may re-register its own names
    /// after a rebuild.
    /// </summary>
    private void Register(LoadedProject project)
    {
        projectDirs.AddOrUpdate(project.Assembly, project.ProjectDir);
        var delegates = new List<Delegate>();
        var attributes = new List<object>();
        var argumentAttributes = new List<List<object>>();
        foreach (var type in project.Assembly.GetTypes().Where(t => t.IsClass && t.IsAbstract && t.IsSealed && t.GetCustomAttribute<VbaDocumentModuleAttribute>() is null))
        {
            foreach (var method in type.GetMethods(BindingFlags.Public | BindingFlags.Static | BindingFlags.DeclaredOnly))
            {
                var parameters = method.GetParameters();
                var isFunction = method.ReturnType != typeof(void);
                if (!isFunction && parameters.Length != 0)
                {
                    continue;
                }

                var name = VbaNames.Of(method);
                foreach (var registeredName in Claim(project, type, name))
                {
                    var target = method;
                    if (isFunction)
                    {
                        var acceptsObjects = parameters.Select(AcceptsObjects).ToList();
                        delegates.Add(FunctionDelegate(target, parameters.Length));
                        attributes.Add(new ExcelFunctionAttribute
                        {
                            Name = registeredName,
                            Category = project.Name,
                            Description = type.Name + "." + name + " (" + project.Name + ")",
                            // Turning a reference into a Range needs the XLM reference functions, and so do
                            // Application.Volatile, Caller, and ThisCell; only a macro-type function may call
                            // them. VBA draws no such line, any Function may, so every UDF registers as one
                            // (ROADMAP.md WP4).
                            IsMacroType = true,
                        });
                        argumentAttributes.Add(parameters.Select((p, i) => (object)new ExcelArgumentAttribute { Name = p.Name ?? "value", AllowReference = acceptsObjects[i] }).ToList());
                    }
                    else
                    {
                        delegates.Add(() => RunCommand(target));
                        attributes.Add(new ExcelCommandAttribute { Name = registeredName });
                        argumentAttributes.Add([]);
                    }
                }
            }
        }

        if (delegates.Count == 0)
        {
            return;
        }

        // Registration needs an XLL context; an event handler is not one.
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                ExcelIntegration.RegisterDelegates(delegates, attributes, argumentAttributes);
                var names = attributes.Select(a => a is ExcelFunctionAttribute f ? f.Name + "()" : ((ExcelCommandAttribute)a).Name);
                log($"Registered {delegates.Count.ToString(CultureInfo.InvariantCulture)} name(s) for {project.Name}: {string.Join(", ", names)}.");
            }
            catch (Exception ex)
            {
                log("Registration failed: " + ex.GetType().Name + ": " + ex.Message);
            }
        });
    }

    /// <summary>The names a procedure registers under: its own name, or Project.Name when another project holds it, and Module.Name.</summary>
    private IEnumerable<string> Claim(LoadedProject project, Type module, string name)
    {
        if (Own(name, project))
        {
            yield return name;
        }
        else
        {
            var fallback = project.Name + "." + name;
            log($"{name} is already registered by another project; {project.Name} registers it as {fallback}.");
            if (Own(fallback, project))
            {
                yield return fallback;
            }
        }

        var qualified = module.Name + "." + name;
        if (Own(qualified, project))
        {
            yield return qualified;
        }
    }

    private bool Own(string name, LoadedProject project)
    {
        if (registered.TryGetValue(name, out var owner) && !string.Equals(owner, project.ProjectDir, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        registered[name] = project.ProjectDir;
        return true;
    }

    /// <summary>A Variant, Object, or Range parameter receives the Range a formula passes, as in VBA; a typed one receives the cell's value.</summary>
    private static bool AcceptsObjects(ParameterInfo parameter)
    {
        var type = parameter.ParameterType.IsByRef ? parameter.ParameterType.GetElementType()! : parameter.ParameterType;
        if (type.IsGenericType && type.GetGenericTypeDefinition() == typeof(ObjectSlot<>))
        {
            // A ByRef object parameter is a slot of the declared class (ROADMAP.md M7 E3).
            type = type.GetGenericArguments()[0];
        }

        return type == typeof(Variant) || type == typeof(object) || type == typeof(IDispatchObject);
    }

    /// <summary>A delegate with one object parameter per VBA parameter, as Excel-DNA marshals a UDF, that funnels into <see cref="RunFunction"/>.</summary>
    private Delegate FunctionDelegate(MethodInfo method, int parameterCount)
    {
        var parameters = Enumerable.Range(0, parameterCount).Select(i => Expression.Parameter(typeof(object), "argument" + i.ToString(CultureInfo.InvariantCulture))).ToArray();
        var run = typeof(WorkbookProjects).GetMethod(nameof(RunFunction), BindingFlags.NonPublic | BindingFlags.Instance)!;
        var body = Expression.Call(Expression.Constant(this), run, Expression.Constant(method), Expression.NewArrayInit(typeof(object), parameters));
        return Expression.Lambda(body, parameters).Compile();
    }

    /// <summary>A UDF call from a formula: the arguments become the Variants VBA would receive, the result what Excel shows; a run-time error shows as #VALUE!, as in VBA.</summary>
    private object RunFunction(MethodInfo method, object?[] arguments)
    {
        try
        {
            // Application.Volatile, Caller, and ThisCell answer from Excel while this call runs (ROADMAP.md WP4).
            using var running = ExcelApplication.RunningFunction();

            // The call's temporaries, the result included, go when the call ends (ARCHITECTURE.md D20).
            using var frame = ObjectRefs.Frame();
            var values = new Variant[arguments.Length];
            for (var i = 0; i < arguments.Length; i++)
            {
                values[i] = FromExcel(arguments[i]);
            }

            return ToExcel(EventBinder.InvokeFunction(method, values));
        }
        catch (EndStatementException end)
        {
            // End stops the call and resets the project (docs/vba-quirks.md); the cell shows #VALUE! as for an error.
            ProjectReset.AfterEnd(end);
            return ExcelError.ExcelErrorValue;
        }
        catch (Exception ex)
        {
            // A UDF's error shows #VALUE! in its cell and no dialog, as in VBA. It reaches the add-in's
            // recent output, which vbang status shows, not the project log: no Running scope is open here.
            ExcelHostServices.Instance.ReportUnhandled($"{method.DeclaringType!.Name}.{VbaNames.Of(method)}", ex, dialog: false);
            return ExcelError.ExcelErrorValue;
        }
    }

    private void RunCommand(MethodInfo method)
    {
        using var running = ExcelHostServices.Instance.Running(ProjectDirOf(method));
        try
        {
            using var frame = ObjectRefs.Frame();
            method.Invoke(null, null);
        }
        catch (TargetInvocationException ex) when (ex.InnerException is EndStatementException end)
        {
            // End stops the macro and resets the project; it is not an error (docs/vba-quirks.md).
            ProjectReset.AfterEnd(end);
        }
        catch (TargetInvocationException ex) when (ex.InnerException is not null)
        {
            ExcelHostServices.Instance.ReportUnhandled($"{method.DeclaringType!.Name}.{VbaNames.Of(method)}", ex.InnerException, dialog: true);
        }
    }

    /// <summary>The folder of the project a registered procedure belongs to; null for one this lifecycle did not register.</summary>
    private string? ProjectDirOf(MethodInfo method) =>
        method.DeclaringType is { } type && projectDirs.TryGetValue(type.Assembly, out var dir) ? dir : null;

    /// <summary>What Excel passes to a UDF, as VBA would see it: a number, text, a Boolean, an error value, Empty for an empty cell, Missing for a left-out argument, a Range for a reference, a 1-based two-dimensional array for an array.</summary>
    private Variant FromExcel(object? value)
    {
        switch (value)
        {
            case null or ExcelMissing:
                return Variant.Missing;
            case ExcelEmpty:
                return Variant.Empty;
            case double d:
                return Variant.FromDouble(d);
            case string s:
                return Variant.FromString(s);
            case bool b:
                return Variant.FromBoolean(b);
            case ExcelError error:
                return Variant.FromError(ErrorValue.FromNumber(2000 + (int)error));
            case ExcelReference reference:
                return RangeOf(reference);
            case object[,] cells:
                {
                    var rows = cells.GetLength(0);
                    var columns = cells.GetLength(1);
                    var array = VbaArray.Create(VarType.Variant, [(1, rows), (1, columns)]);
                    for (var row = 0; row < rows; row++)
                    {
                        for (var column = 0; column < columns; column++)
                        {
                            array.Set([row + 1, column + 1], FromExcel(cells[row, column]));
                        }
                    }

                    // The frame of the call releases the cells' strings when it ends (D20).
                    return Variant.FromArray(ObjectRefs.Owned(array));
                }

            default:
                return Variant.FromDouble(System.Convert.ToDouble(value, CultureInfo.InvariantCulture));
        }
    }

    /// <summary>The Range behind a reference, through the running Application; Empty when no Application is at hand. Also what Caller and ThisCell answer with inside a UDF (<see cref="ExcelApplication"/>).</summary>
    internal Variant RangeOf(ExcelReference reference)
    {
        if (application is null || XlCall.Excel(XlCall.xlfReftext, reference, true) is not string address)
        {
            return Variant.Empty;
        }

        return application.Invoke(application.GetDispId("Range"), InvokeKind.PropertyGet, [Variant.FromString(address)]);
    }

    /// <summary>What a formula shows for a function's result: numbers as Double, dates as serials, VBA error values as Excel errors, arrays as two-dimensional arrays (a one-dimensional array as a row), an object as its default value, Empty as an empty result.</summary>
    private object ToExcel(in Variant value)
    {
        switch (value.Type)
        {
            case VarType.Empty:
                return ExcelEmpty.Value;
            case VarType.Null:
                return ExcelError.ExcelErrorValue;
            case VarType.String:
                return value.AsString();
            case VarType.Boolean:
                return value.AsBoolean();
            case VarType.Error:
                {
                    var number = value.AsError().Number;
                    return number is >= 2000 and <= 2043 ? (ExcelError)(number - 2000) : ExcelError.ExcelErrorValue;
                }

            case VarType.Date:
                return value.AsDate().Serial;
            case VarType.Array:
                return ArrayToExcel(value.AsArray());
            case VarType.Object:
                return value.AsObject() is null ? ExcelError.ExcelErrorValue : ToExcel(Coerce.LetValue(value));
            case VarType.UserDefinedType:
                return ExcelError.ExcelErrorValue;
            default:
                return Coerce.ToDouble(value);
        }
    }

    private object ArrayToExcel(VbaArray array)
    {
        if (!array.IsAllocated || array.Rank > 2)
        {
            return ExcelError.ExcelErrorValue;
        }

        if (array.Rank == 1)
        {
            var lower = array.LBound(1);
            var count = array.UBound(1) - lower + 1;
            var row = new object[1, count];
            for (var i = 0; i < count; i++)
            {
                row[0, i] = ToExcel(array.Get([lower + i]));
            }

            return row;
        }

        var rowLower = array.LBound(1);
        var columnLower = array.LBound(2);
        var rows = array.UBound(1) - rowLower + 1;
        var columns = array.UBound(2) - columnLower + 1;
        var cells = new object[rows, columns];
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < columns; c++)
            {
                cells[r, c] = ToExcel(array.Get([rowLower + r, columnLower + c]));
            }
        }

        return cells;
    }

    /// <summary>One workbook's binding: its project, the objects its document modules are bound to, and their event sinks.</summary>
    private sealed class BoundWorkbook : IDisposable
    {
        private readonly List<(Type Module, ComObject Target)> documents = [];
        private readonly ComObject workbook;
        private readonly Action<string> log;
        private readonly VbaNg.Runtime.TypeLibraries.ComLibrary? excelLibrary;
        private EventBinder? events;

        public BoundWorkbook(string name, ComObject workbook, Action<string> log, VbaNg.Runtime.TypeLibraries.ComLibrary? excelLibrary)
        {
            Name = name;
            this.workbook = workbook;
            this.log = log;
            this.excelLibrary = excelLibrary;

            // The binding keeps the workbook and its sheets beyond the statement that fetched them (ARCHITECTURE.md D18).
            workbook.AddRef();
        }

        public string Name { get; }

        public LoadedProject Project { get; private set; } = null!;

        public int DocumentCount => documents.Count;

        public int SinkCount => events?.SinkCount ?? 0;

        public void Bind(LoadedProject project)
        {
            Project = project;
            events = new EventBinder(log, project.ProjectDir);
            foreach (var module in project.Assembly.GetTypes())
            {
                var attribute = module.GetCustomAttribute<VbaDocumentModuleAttribute>();
                if (attribute is null)
                {
                    continue;
                }

                var target = attribute.Kind.Equals("Workbook", StringComparison.OrdinalIgnoreCase) ? workbook : FindSheet(module.Name);
                if (target is null)
                {
                    log($"{Name}: no sheet with CodeName {module.Name}; its module is not bound.");
                    continue;
                }

                module.GetField(VbaDocumentModuleAttribute.MeField, BindingFlags.Public | BindingFlags.Static)?.SetValue(null, target);
                target.AddRef();
                documents.Add((module, target));

                var handlers = module.GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static | BindingFlags.DeclaredOnly)
                    .Where(m => m.GetCustomAttribute<VbaEventHandlerAttribute>() is not null)
                    .GroupBy(m => m.GetCustomAttribute<VbaEventHandlerAttribute>()!.Source, StringComparer.OrdinalIgnoreCase);
                foreach (var group in handlers)
                {
                    if (group.Key.Equals(attribute.Kind, StringComparison.OrdinalIgnoreCase))
                    {
                        var fallback = excelLibrary is null ? null : EventSource.FromLibrary(excelLibrary, attribute.Kind);
                        events.Bind(target, module.Name, group.ToList(), fallback);
                    }
                    else if (FindControl(target, group.Key) is { } control)
                    {
                        events.Bind(control, module.Name + "." + group.Key, group.ToList());
                    }
                    else
                    {
                        log($"{Name}: {module.Name} has no control named {group.Key}; {group.Count()} handler(s) not connected.");
                    }
                }
            }
        }

        public void Dispose()
        {
            Unbind();
            workbook.Release();
        }

        public void Unbind()
        {
            events?.Dispose();
            events = null;
            foreach (var (module, target) in documents)
            {
                module.GetField(VbaDocumentModuleAttribute.MeField, BindingFlags.Public | BindingFlags.Static)?.SetValue(null, null);
                target.Release();
            }

            documents.Clear();
        }

        /// <summary>Workbook_Open of ThisWorkbook, then every public Auto_Open (the WorkbookOpen event fired before the sink existed).</summary>
        public void RunOpenHandlers()
        {
            RunHandler("Workbook", "Open");
            RunPublicSubs("Auto_Open");
        }

        public void RunCloseHandlers() => RunPublicSubs("Auto_Close");

        private void RunHandler(string source, string eventName)
        {
            foreach (var (module, _) in documents)
            {
                foreach (var method in module.GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static | BindingFlags.DeclaredOnly))
                {
                    var attribute = method.GetCustomAttribute<VbaEventHandlerAttribute>();
                    if (attribute is not null && attribute.Source.Equals(source, StringComparison.OrdinalIgnoreCase) && attribute.EventName.Equals(eventName, StringComparison.OrdinalIgnoreCase))
                    {
                        Run(method, []);
                    }
                }
            }
        }

        private void RunPublicSubs(string name)
        {
            foreach (var type in Project.Assembly.GetTypes().Where(t => t.IsClass && t.IsAbstract && t.IsSealed && t.GetCustomAttribute<VbaDocumentModuleAttribute>() is null))
            {
                foreach (var method in type.GetMethods(BindingFlags.Public | BindingFlags.Static | BindingFlags.DeclaredOnly))
                {
                    if (method.GetParameters().Length == 0 && VbaNames.Of(method).Equals(name, StringComparison.OrdinalIgnoreCase))
                    {
                        Run(method, []);
                    }
                }
            }
        }

        private void Run(MethodInfo method, Variant[] arguments)
        {
            // The frame is the event sink's or the caller's: what a ByRef parameter leaves behind must outlive the write-back.
            using var running = ExcelHostServices.Instance.Running(Project.ProjectDir);
            try
            {
                EventBinder.Invoke(method, arguments);
            }
            catch (EndStatementException end)
            {
                // End in a handler stops it and resets the project (docs/vba-quirks.md).
                ProjectReset.AfterEnd(end);
            }
            catch (Exception ex) when (ex is VbaException or InvalidOperationException or TargetInvocationException)
            {
                ExcelHostServices.Instance.ReportUnhandled($"{method.DeclaringType!.Name}.{method.Name}", ex is TargetInvocationException { InnerException: { } inner } ? inner : ex, dialog: true);
            }
        }

        private ComObject? FindSheet(string codeName)
        {
            var sheets = workbook.Invoke(workbook.GetDispId("Sheets"), InvokeKind.PropertyGet, []);
            if (sheets.AsObject() is not ComObject collection)
            {
                return null;
            }

            using var enumerator = collection.Enumerate();
            while (enumerator.MoveNext())
            {
                if (enumerator.Current.AsObject() is ComObject sheet)
                {
                    var name = Coerce.ToString(sheet.Invoke(sheet.GetDispId("CodeName"), InvokeKind.PropertyGet, []));
                    if (name.Equals(codeName, StringComparison.OrdinalIgnoreCase))
                    {
                        return sheet;
                    }
                }
            }

            return null;
        }

        private static ComObject? FindControl(ComObject sheet, string name)
        {
            try
            {
                var oleObject = sheet.Invoke(sheet.GetDispId("OLEObjects"), InvokeKind.PropertyGet, [Variant.FromString(name)]);
                if (oleObject.AsObject() is not ComObject wrapper)
                {
                    return null;
                }

                return wrapper.Invoke(wrapper.GetDispId("Object"), InvokeKind.PropertyGet, []).AsObject() as ComObject;
            }
            catch (VbaException)
            {
                return null;
            }
        }
    }

    internal static string Invariant(FormattableString text) => text.ToString(CultureInfo.InvariantCulture);
}
