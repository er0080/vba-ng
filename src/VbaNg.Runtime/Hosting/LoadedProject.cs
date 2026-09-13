using System.Reflection;
using System.Runtime.ExceptionServices;

namespace VbaNg.Runtime.Hosting;

/// <summary>A compiled project loaded into its own <see cref="ProjectLoadContext"/>.</summary>
public sealed class LoadedProject
{
    private ProjectLoadContext? context;

    internal LoadedProject(string name, string projectDir, Assembly assembly, ProjectLoadContext context, string stamp)
    {
        Name = name;
        ProjectDir = projectDir;
        Assembly = assembly;
        Stamp = stamp;
        this.context = context;
    }

    public string Name { get; }

    public string ProjectDir { get; }

    public Assembly Assembly { get; }

    /// <summary>Content hash of the assembly that was loaded; used to detect a rebuilt project.</summary>
    public string Stamp { get; }

    public bool IsLoaded => context is not null;

    /// <summary>
    /// Runs a public, parameterless procedure written as <c>Module.Procedure</c>.
    /// Names are case-insensitive, as in VBA. Exceptions thrown by the procedure propagate unwrapped.
    /// </summary>
    public void Run(string procedure)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(procedure);
        if (context is null)
        {
            throw new InvalidOperationException($"Project '{Name}' has been unloaded.");
        }

        var dot = procedure.LastIndexOf('.');
        if (dot <= 0 || dot == procedure.Length - 1)
        {
            throw new ArgumentException("Procedure must be written as Module.Procedure.", nameof(procedure));
        }

        var moduleName = procedure[..dot];
        var procedureName = procedure[(dot + 1)..];

        var module = Assembly.GetTypes().FirstOrDefault(t =>
            t.IsClass && string.Equals(t.Name, moduleName, StringComparison.OrdinalIgnoreCase))
            ?? throw new MissingMethodException($"Module '{moduleName}' was not found in project '{Name}'.");

        var method = module.GetMethods(BindingFlags.Public | BindingFlags.Static | BindingFlags.DeclaredOnly)
            .FirstOrDefault(m => m.GetParameters().Length == 0 && string.Equals(VbaNames.Of(m), procedureName, StringComparison.OrdinalIgnoreCase))
            ?? throw new MissingMethodException($"Procedure '{procedureName}' was not found in module '{module.Name}'.");

        try
        {
            method.Invoke(null, null);
        }
        catch (TargetInvocationException ex) when (ex.InnerException is EndStatementException end)
        {
            // End stops the macro and resets the project; it is not an error (docs/vba-quirks.md).
            ProjectReset.AfterEnd(end);
        }
        catch (TargetInvocationException ex) when (ex.InnerException is not null)
        {
            ExceptionDispatchInfo.Capture(ex.InnerException).Throw();
        }
        finally
        {
            Debug.Flush();
        }
    }

    /// <summary>
    /// Unloads the project as VBA unloads one when its workbook closes: its module-level storage is
    /// reset first, which destroys the objects it holds without their Class_Terminate
    /// (docs/vba-quirks.md), so nothing the project made keeps its load context alive.
    /// </summary>
    internal void Unload()
    {
        if (context is null)
        {
            return;
        }

        try
        {
            ProjectReset.Run(Assembly);
        }
        finally
        {
            context.Unload();
            context = null;
        }
    }
}
