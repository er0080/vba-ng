using System.Globalization;
using System.Reflection;
using System.Text;

using VbaNg.Compiler;
using VbaNg.Golden.Harness;
using VbaNg.Runtime;
using VbaNg.Runtime.Hosting;
using VbaNg.Runtime.Library;

namespace VbaNg.Golden.Replay;

/// <summary>
/// Replays a golden area through the real compiler (ARCHITECTURE.md section 11): every case
/// becomes a Sub of one module, the module is bound and compiled with Roslyn, loaded into a
/// collectible context, and each case runs in turn. The case's <c>?</c> lines call a
/// <c>GoldenRecord</c> Sub that collects values in a Collection, so a value crosses no boundary
/// the recorder in Excel did not cross either. A case the binder rejects is reported as
/// unsupported, named with the diagnostic, and the rest of the area still runs.
/// </summary>
public static class CompiledReplay
{
    private const string ResultsField = "GoldenResults";

    /// <summary>stdole, which every VBA project references, as the recording's project in Excel does (MS-VBAL 5.6.16.7): a copy of its type library model, so the replay needs no registry.</summary>
    private static readonly Lazy<VbaNg.Runtime.TypeLibraries.ComLibrary> Stdole = new(() =>
        VbaNg.Runtime.TypeLibraries.ComLibrary.Load(Path.Combine(GoldenPaths.RepositoryRoot, "tests", "VbaNg.Golden", "Fixtures", "TypeLibs", "stdole.json")));

    public static IReadOnlyList<CaseOutcome> Run(GoldenArea golden)
    {
        ArgumentNullException.ThrowIfNull(golden);

        // CreateObject reaches COM through the Interop provider, as it does in the add-in (ARCHITECTURE.md section 6).
        Com.Provider = new VbaNg.Interop.ComProvider();
        var excluded = new Dictionary<int, string>();
        IReadOnlyList<(string Module, string Text)>? generated = null;
        var moduleName = RecorderModule.ModuleName(golden.Area);
        var sourceDir = Path.Combine(Path.GetTempPath(), "vbang-golden");
        var sourcePath = Path.Combine(sourceDir, moduleName + ".bas");
        var classFiles = (golden.Classes ?? []).Select(c => new SourceFile(Path.Combine(sourceDir, c.Name + ".cls"), string.Join("\r\n", c.Source) + "\r\n"))
            .Concat((golden.Modules ?? []).Select(m => new SourceFile(Path.Combine(sourceDir, m.Name + ".bas"), string.Join("\r\n", m.Source) + "\r\n")))
            .ToList();

        // Bind until every remaining case binds; each round names the cases whose lines carry errors.
        for (var round = 0; round <= golden.Cases.Count; round++)
        {
            var module = Generate(golden, excluded);
            var diagnostics = new List<Diagnostic>();
            generated = ProjectCompiler.Generate(moduleName, [new SourceFile(sourcePath, module.Text), .. classFiles], diagnostics, libraries: [Stdole.Value]);
            if (generated is not null)
            {
                break;
            }

            var newlyExcluded = false;
            foreach (var diagnostic in diagnostics.Where(d => d.IsError))
            {
                if (!string.Equals(diagnostic.FilePath, sourcePath, StringComparison.OrdinalIgnoreCase))
                {
                    return golden.Cases.Select(c => new CaseOutcome(c.Name, CaseVerdict.Fail, "a class module of the area does not compile: " + diagnostic)).ToList();
                }

                var index = CaseAt(module.CaseLines, diagnostic.Line);
                if (index < 0)
                {
                    return golden.Cases.Select(c => new CaseOutcome(c.Name, CaseVerdict.Fail, "the area module does not compile: " + diagnostic)).ToList();
                }

                if (excluded.TryAdd(index, "does not compile: " + diagnostic.Id + ": " + diagnostic.Message))
                {
                    newlyExcluded = true;
                }
            }

            if (!newlyExcluded)
            {
                return golden.Cases.Select(c => new CaseOutcome(c.Name, CaseVerdict.Fail, "binding failed without a case to blame: " + string.Join("; ", diagnostics))).ToList();
            }
        }

        if (generated is null)
        {
            return golden.Cases.Select(c => new CaseOutcome(c.Name, CaseVerdict.Fail, "binding did not converge")).ToList();
        }

        var buildDir = Path.Combine(Path.GetTempPath(), "vbang-golden", Guid.NewGuid().ToString("N", CultureInfo.InvariantCulture));
        Directory.CreateDirectory(buildDir);
        var assemblyPath = Path.Combine(buildDir, moduleName + ".dll");
        var files = generated.Select(g => new GeneratedFile(Path.Combine(buildDir, g.Module + ".cs"), g.Text)).ToList();
        foreach (var file in files)
        {
            File.WriteAllText(file.Path, file.Text);
        }

        var compileErrors = CSharpCompiler.Compile(moduleName, files, assemblyPath, Path.Combine(buildDir, moduleName + ".pdb"));
        if (compileErrors.Count > 0)
        {
            var summary = string.Join("; ", compileErrors.Take(5));
            return golden.Cases.Select(c => new CaseOutcome(c.Name, CaseVerdict.Fail, "generated C# does not compile: " + summary)).ToList();
        }

        var context = new ProjectLoadContext(moduleName);
        Assembly? assembly = null;
        try
        {
            assembly = context.LoadFromAssemblyPath(assemblyPath);
            var type = assembly.GetType(moduleName) ?? throw new InvalidOperationException($"Generated class {moduleName} not found.");
            var outcomes = new List<CaseOutcome>();
            for (var i = 0; i < golden.Cases.Count; i++)
            {
                var golden1 = golden.Cases[i];
                outcomes.Add(excluded.TryGetValue(i, out var reason)
                    ? new CaseOutcome(golden1.Name, CaseVerdict.Unsupported, reason)
                    : RunCase(type, i + 1, golden1));
            }

            return outcomes;
        }
        finally
        {
            // The area's module-level storage goes as its workbook's would when it closes (docs/vba-quirks.md), on a thread of its own because Err is per thread.
            if (assembly is not null)
            {
                var reset = new Thread(() => ProjectReset.Run(assembly)) { IsBackground = true, Name = "golden reset" };
                reset.Start();
                reset.Join(CaseTimeout);
            }

            context.Unload();
        }
    }

    /// <summary>A case that runs longer than this is a loop that never ends; it is reported by name and its thread abandoned so the area still finishes.</summary>
    private static readonly TimeSpan CaseTimeout = TimeSpan.FromSeconds(30);

    /// <summary>Runs the case on its own thread (Err is per thread, so the whole case lives there) and gives up on it after <see cref="CaseTimeout"/>.</summary>
    private static CaseOutcome RunCase(Type module, int number, GoldenCase golden)
    {
        CaseOutcome? outcome = null;
        var worker = new Thread(() => outcome = RunCaseOnThisThread(module, number, golden), 16 * 1024 * 1024)
        {
            IsBackground = true,
            Name = "golden case " + number.ToString(CultureInfo.InvariantCulture),
        };
        worker.Start();
        if (!worker.Join(CaseTimeout))
        {
            return new CaseOutcome(golden.Name, CaseVerdict.Fail, $"did not finish within {CaseTimeout.TotalSeconds:F0} s (a loop that never ends); its thread was abandoned");
        }

        return outcome!;
    }

    private static CaseOutcome RunCaseOnThisThread(Type module, int number, GoldenCase golden)
    {
        var reset = module.GetMethod("GoldenReset", BindingFlags.Public | BindingFlags.Static)!;
        var run = module.GetMethod("GoldenCase" + number.ToString(CultureInfo.InvariantCulture), BindingFlags.Public | BindingFlags.Static)!;
        var results = module.GetField(ResultsField, BindingFlags.Public | BindingFlags.Static)!;

        // A fresh results collection, then the count of live BSTRs and arrays the case starts from; the
        // reset that follows the case releases the results it recorded, so what is left
        // allocated afterwards is the case's own (module-level storage it filled, or a leak).
        reset.Invoke(null, null);
        var liveBefore = Bstr.LiveCount + VbaArray.LiveCount;
        var outcome = Replay(run, results, golden);
        reset.Invoke(null, null);
        return outcome with { Leaked = Bstr.LiveCount + VbaArray.LiveCount - liveBefore };
    }

    private static CaseOutcome Replay(MethodInfo run, FieldInfo results, GoldenCase golden)
    {
        // The harness is a host: what a case hands back to it goes when the case ends (ARCHITECTURE.md D20).
        using var frame = ObjectRefs.Frame();
        Err.Current.Clear();
        Err.Current.Line = 0;
        RandomGenerator.Shared = new RandomGenerator();
        VbaException? failure = null;
        try
        {
            run.Invoke(null, null);
            if (Err.Current.Number != 0)
            {
                // An error skipped by Resume Next stays in Err when the procedure returns; the recorder reported it as the case's error.
                failure = new VbaException(Err.Current.Number, Err.Current.DescriptionText, Err.Current.SourceText, Err.Current.HelpFileText, Err.Current.HelpContext);
            }
        }
        catch (TargetInvocationException ex) when (ex.InnerException is VbaException error)
        {
            failure = error;
        }
        catch (TargetInvocationException ex) when (ex.InnerException is not null)
        {
            return new CaseOutcome(golden.Name, CaseVerdict.Fail, $"compiled case threw {ex.InnerException.GetType().Name}: {ex.InnerException.Message}");
        }
        finally
        {
            Err.Current.Clear();
        }

        // The module's Results variable is a typed object variable, so its storage is an object slot (ROADMAP.md M7 E3).
        var recorded = ((ObjectSlot<Collection>)results.GetValue(null)!).Target?.ToList() ?? [];
        return CaseRunner.Compare(golden, recorded, failure);
    }

    /// <summary>The module for an area: options, the recording helpers, then one Public Sub per case; excluded cases become empty Subs so numbering holds.</summary>
    private static GeneratedModule Generate(GoldenArea golden, Dictionary<int, string> excluded)
    {
        var builder = new StringBuilder();
        var lineNumber = 0;
        void Line(string text = "")
        {
            builder.Append(text).Append("\r\n");
            lineNumber++;
        }

        foreach (var option in golden.Options)
        {
            Line(option);
        }

        Line();
        Line($"Public {ResultsField} As Collection");
        Line();
        foreach (var declaration in golden.Declarations ?? [])
        {
            Line(declaration);
        }

        if (golden.Declarations is { Count: > 0 })
        {
            Line();
        }

        Line("Public Sub GoldenReset()");
        Line($"    Set {ResultsField} = New Collection");
        Line("End Sub");
        Line();
        Line($"Public Sub {RecorderModule.RecordProcedure}(ByVal caseIndex As Long, ByVal resultIndex As Long, ByVal value As Variant)");
        Line($"    {ResultsField}.Add value");
        Line("End Sub");

        var caseLines = new List<int>();
        for (var i = 0; i < golden.Cases.Count; i++)
        {
            var number = i + 1;
            var source = golden.Cases[i];
            Line();
            Line(string.Create(CultureInfo.InvariantCulture, $"' {source.Name}"));
            Line(string.Create(CultureInfo.InvariantCulture, $"Public Sub GoldenCase{number}()"));
            caseLines.Add(lineNumber + 1);
            if (!excluded.ContainsKey(i))
            {
                var resultIndex = 0;
                foreach (var line in source.Source)
                {
                    if (GoldenCaseSource.IsResultLine(line))
                    {
                        Line(string.Create(CultureInfo.InvariantCulture, $"    {RecorderModule.RecordProcedure} {number}, {resultIndex}, {GoldenCaseSource.ResultExpression(line)}"));
                        resultIndex++;
                    }
                    else
                    {
                        Line("    " + line.Trim());
                    }
                }
            }

            Line("End Sub");
        }

        return new GeneratedModule(builder.ToString(), caseLines);
    }

    private static int CaseAt(IReadOnlyList<int> caseLines, int moduleLine)
    {
        var index = -1;
        for (var i = 0; i < caseLines.Count && caseLines[i] <= moduleLine; i++)
        {
            index = i;
        }

        return index;
    }
}
