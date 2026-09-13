using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Text;
using System.Xml;

namespace VbaNg.Runtime.Hosting;

public enum TestOutcome
{
    Passed,

    /// <summary>An <c>Assert</c> call failed.</summary>
    Failed,

    /// <summary>The procedure ended with an unhandled run-time error.</summary>
    Error,
}

/// <summary>The outcome of one test procedure, with the <c>Debug.Print</c> lines it wrote.</summary>
public sealed record TestResult(string Module, string Name, TestOutcome Outcome, string? Message, IReadOnlyList<string> Output, TimeSpan Duration)
{
    public string FullName => Module + "." + Name;
}

/// <summary>Every test of a project, in module and declaration order, with the text and JUnit reports <c>vbang test</c> prints (ARCHITECTURE.md section 8).</summary>
public sealed class TestRunResult
{
    internal TestRunResult(string project, IReadOnlyList<TestResult> results, TimeSpan duration)
    {
        Project = project;
        Results = results;
        Duration = duration;
    }

    public string Project { get; }

    public IReadOnlyList<TestResult> Results { get; }

    public TimeSpan Duration { get; }

    public int Passed => Results.Count(r => r.Outcome == TestOutcome.Passed);

    public int Failed => Results.Count(r => r.Outcome == TestOutcome.Failed);

    public int Errors => Results.Count(r => r.Outcome == TestOutcome.Error);

    public bool AllPassed => Results.All(r => r.Outcome == TestOutcome.Passed);

    /// <summary>One line per test, the reason and output under each failure, then a summary line.</summary>
    public string Summary =>
        string.Create(
            CultureInfo.InvariantCulture,
            $"{Results.Count} {(Results.Count == 1 ? "test" : "tests")}: {Passed} passed, {Failed} failed, {Errors} {(Errors == 1 ? "error" : "errors")}. {Duration.TotalMilliseconds:0} ms.");

    public string ToText()
    {
        var text = new StringBuilder();
        foreach (var result in Results)
        {
            var label = result.Outcome switch
            {
                TestOutcome.Passed => "passed",
                TestOutcome.Failed => "FAILED",
                _ => "ERROR ",
            };
            text.Append(label).Append("  ").AppendLine(result.FullName);
            if (result.Outcome == TestOutcome.Passed)
            {
                continue;
            }

            if (result.Message is not null)
            {
                text.Append("        ").AppendLine(result.Message);
            }

            foreach (var line in result.Output)
            {
                text.Append("        | ").AppendLine(line);
            }
        }

        text.Append(Summary);
        return text.ToString();
    }

    /// <summary>The JUnit XML report: one testsuite per module, one testcase per test, with the test's output as system-out.</summary>
    public string ToJUnitXml()
    {
        var settings = new XmlWriterSettings { Indent = true, Encoding = new UTF8Encoding(false) };
        using var stream = new MemoryStream();
        using (var xml = XmlWriter.Create(stream, settings))
        {
            xml.WriteStartDocument();
            xml.WriteStartElement("testsuites");
            xml.WriteAttributeString("name", Project);
            WriteCounts(xml, Results, Duration);
            foreach (var module in Results.GroupBy(r => r.Module, StringComparer.Ordinal))
            {
                var tests = module.ToList();
                xml.WriteStartElement("testsuite");
                xml.WriteAttributeString("name", module.Key);
                WriteCounts(xml, tests, TimeSpan.FromTicks(tests.Sum(t => t.Duration.Ticks)));
                foreach (var test in tests)
                {
                    xml.WriteStartElement("testcase");
                    xml.WriteAttributeString("classname", test.Module);
                    xml.WriteAttributeString("name", test.Name);
                    xml.WriteAttributeString("time", Seconds(test.Duration));
                    if (test.Outcome != TestOutcome.Passed)
                    {
                        xml.WriteStartElement(test.Outcome == TestOutcome.Failed ? "failure" : "error");
                        xml.WriteAttributeString("message", test.Message ?? string.Empty);
                        xml.WriteAttributeString("type", test.Outcome == TestOutcome.Failed ? "AssertFailed" : "RuntimeError");
                        xml.WriteString(test.Message ?? string.Empty);
                        xml.WriteEndElement();
                    }

                    if (test.Output.Count > 0)
                    {
                        xml.WriteStartElement("system-out");
                        xml.WriteString(string.Join("\n", test.Output));
                        xml.WriteEndElement();
                    }

                    xml.WriteEndElement();
                }

                xml.WriteEndElement();
            }

            xml.WriteEndElement();
            xml.WriteEndDocument();
        }

        return Encoding.UTF8.GetString(stream.ToArray());
    }

    private static void WriteCounts(XmlWriter xml, IReadOnlyList<TestResult> tests, TimeSpan duration)
    {
        xml.WriteAttributeString("tests", tests.Count.ToString(CultureInfo.InvariantCulture));
        xml.WriteAttributeString("failures", tests.Count(t => t.Outcome == TestOutcome.Failed).ToString(CultureInfo.InvariantCulture));
        xml.WriteAttributeString("errors", tests.Count(t => t.Outcome == TestOutcome.Error).ToString(CultureInfo.InvariantCulture));
        xml.WriteAttributeString("time", Seconds(duration));
    }

    private static string Seconds(TimeSpan duration) => duration.TotalSeconds.ToString("0.000", CultureInfo.InvariantCulture);
}

/// <summary>
/// Runs the test procedures of a built project (ARCHITECTURE.md section 8): every public static
/// method marked <see cref="VbaTestAttribute"/>, in module and declaration order. Each test starts
/// with a clean <c>Err</c>; an <see cref="AssertFailedException"/> is a failure, any other
/// exception an error. <c>Debug.Print</c> output still reaches the current host and is also
/// kept per test for the reports.
/// </summary>
public static class TestRunner
{
    public static TestRunResult Run(LoadedProject project)
    {
        ArgumentNullException.ThrowIfNull(project);
        return Run(project.Name, project.Assembly);
    }

    public static TestRunResult Run(string projectName, Assembly assembly)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(projectName);
        ArgumentNullException.ThrowIfNull(assembly);
        var total = Stopwatch.StartNew();
        var results = new List<TestResult>();
        foreach (var method in Discover(assembly))
        {
            results.Add(RunOne(method));
        }

        return new TestRunResult(projectName, results, total.Elapsed);
    }

    /// <summary>The test methods of an assembly in the order their modules and procedures were declared.</summary>
    public static IReadOnlyList<MethodInfo> Discover(Assembly assembly)
    {
        ArgumentNullException.ThrowIfNull(assembly);
        return assembly.GetTypes()
            .Where(t => t.IsClass && t.IsAbstract && t.IsSealed)
            .OrderBy(t => t.MetadataToken)
            .SelectMany(t => t.GetMethods(BindingFlags.Public | BindingFlags.Static | BindingFlags.DeclaredOnly)
                .Where(m => m.GetCustomAttribute<VbaTestAttribute>() is not null && m.GetParameters().Length == 0)
                .OrderBy(m => m.MetadataToken))
            .ToList();
    }

    private static TestResult RunOne(MethodInfo method)
    {
        var module = method.DeclaringType!.Name;
        var name = VbaNames.Of(method);
        var capture = new CapturingHostServices();
        var previous = Host.Current;
        Host.Current = new TeeHostServices(previous, capture);
        Err.Current.Clear();
        Err.Current.Line = 0;
        var clock = Stopwatch.StartNew();
        try
        {
            try
            {
                method.Invoke(null, null);
            }
            catch (TargetInvocationException ex) when (ex.InnerException is not null)
            {
                throw ex.InnerException;
            }

            return new TestResult(module, name, TestOutcome.Passed, null, capture.Lines, clock.Elapsed);
        }
        catch (AssertFailedException ex)
        {
            return new TestResult(module, name, TestOutcome.Failed, ex.Message, capture.Lines, clock.Elapsed);
        }
        catch (EndStatementException end)
        {
            // End resets the project (docs/vba-quirks.md); the test that ran it is reported as an error.
            ProjectReset.AfterEnd(end);
            return new TestResult(module, name, TestOutcome.Error, "End statement executed.", capture.Lines, clock.Elapsed);
        }
        catch (Exception ex)
        {
            var error = ex is UnhandledErrorException { Error: { } unhandled } ? unhandled : VbaException.From(ex);
            var message = error is null
                ? ex.GetType().Name + ": " + ex.Message
                : string.Create(CultureInfo.InvariantCulture, $"Run-time error {error.Number}: {error.Description}");
            return new TestResult(module, name, TestOutcome.Error, message, capture.Lines, clock.Elapsed);
        }
        finally
        {
            VbaNg.Runtime.Debug.Flush();
            Host.Current = previous;
            Err.Current.Clear();
            Err.Current.Line = 0;
        }
    }

    private sealed class TeeHostServices(IHostServices inner, CapturingHostServices capture) : IHostServices
    {
        public void Print(string text)
        {
            capture.Print(text);
            inner.Print(text);
        }
    }
}
