using System.Collections.Concurrent;
using System.Reflection;
using System.Runtime.ExceptionServices;

namespace VbaNg.Runtime.Hosting;

/// <summary>
/// Resets a project the way VBA resets one at <c>End</c> and when its workbook closes
/// (docs/vba-quirks.md; ARCHITECTURE.md section 5, value model): every module-level variable and
/// <c>Static</c> local returns to its initial value, and the objects the project held are
/// destroyed without their <c>Class_Terminate</c> running. Each module's generated
/// <see cref="MethodName"/> does the clearing; while a reset of the project is under way, a class
/// instance of the project whose last reference goes releases what it holds and skips
/// <c>Class_Terminate</c>. Objects of another project keep their events.
/// </summary>
public static class ProjectReset
{
    /// <summary>The generated method of a module that returns its storage to its initial values.</summary>
    public const string MethodName = "__Reset";

    private static readonly ConcurrentDictionary<Assembly, int> Resetting = new();

    /// <summary>Resets every module of the project; a module whose reset raises does not stop the others, and the first error is raised afterwards.</summary>
    public static void Run(Assembly project)
    {
        ArgumentNullException.ThrowIfNull(project);
        Begin(project);
        ExceptionDispatchInfo? failure = null;
        try
        {
            foreach (var module in project.GetTypes().OrderBy(t => t.FullName, StringComparer.Ordinal))
            {
                var reset = module.GetMethod(MethodName, BindingFlags.Public | BindingFlags.Static | BindingFlags.DeclaredOnly, Type.EmptyTypes);
                if (reset is null)
                {
                    continue;
                }

                try
                {
                    reset.Invoke(null, null);
                }
                catch (TargetInvocationException ex) when (ex.InnerException is not null)
                {
                    failure ??= ExceptionDispatchInfo.Capture(ex.InnerException);
                }
            }

            // Instances the modules no longer reach (a reference cycle, a COM caller's reference) go too, as VBA destroys every object of the project.
            RuntimeObject.DestroyInstances(project);
        }
        finally
        {
            Finish(project);
        }

        failure?.Throw();
    }

    /// <summary>
    /// The End statement (MS-VBAL 5.4.2.15) in a module of the project. The project is marked as
    /// being reset before the exception unwinds, so the locals its frames release go without
    /// Class_Terminate, as VBA destroys them (docs/vba-quirks.md). The host that catches the
    /// exception calls <see cref="AfterEnd"/>.
    /// </summary>
    public static EndStatementException Ending(Type module)
    {
        ArgumentNullException.ThrowIfNull(module);
        Begin(module.Assembly);
        return new EndStatementException(module.Assembly);
    }

    /// <summary>A host caught End: the files Open left open are closed, the project's module-level storage is reset, then the mark <see cref="Ending"/> set is lifted.</summary>
    public static void AfterEnd(EndStatementException end)
    {
        ArgumentNullException.ThrowIfNull(end);

        // End closes every file Open opened, as VBA does (MS-VBAL 5.4.2.15), before the storage
        // that held their file numbers goes. The channel table is the runtime's rather than one
        // project's, so a project ending closes another loaded project's files too; VBA gives
        // each project its own, which only shows when two of them hold files at once.
        Library.FileSystem.Reset();

        if (end.Project is not { } project)
        {
            return;
        }

        try
        {
            Run(project);
        }
        finally
        {
            Finish(project);
        }
    }

    /// <summary>True while the project that declares <paramref name="type"/> is being reset: its instances go without Class_Terminate.</summary>
    internal static bool IsResetting(Type type) => !Resetting.IsEmpty && Resetting.ContainsKey(type.Assembly);

    private static void Begin(Assembly project) => Resetting.AddOrUpdate(project, 1, (_, depth) => depth + 1);

    private static void Finish(Assembly project)
    {
        while (Resetting.TryGetValue(project, out var depth))
        {
            if (depth <= 1 ? Resetting.TryRemove(new KeyValuePair<Assembly, int>(project, depth)) : Resetting.TryUpdate(project, depth - 1, depth))
            {
                return;
            }
        }
    }
}
