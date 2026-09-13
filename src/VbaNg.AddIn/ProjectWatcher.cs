using ExcelDna.Integration;

namespace VbaNg.AddIn;

/// <summary>
/// Hot reload (ARCHITECTURE.md section 7): watches a bound project folder for changes to its
/// sources and manifest, waits for the editor to finish writing, builds on a worker thread so
/// Excel stays responsive, and hands the rebuilt project to Excel's thread to be swapped in
/// between macros. A change during a build queues one more build; a failed build is logged and
/// the loaded project stays as it was.
/// </summary>
internal sealed class ProjectWatcher : IDisposable
{
    private static readonly TimeSpan Debounce = TimeSpan.FromMilliseconds(500);

    private readonly FileSystemWatcher watcher;
    private readonly Action<string> log;
    private readonly Action reload;
    private readonly Lock gate = new();
    private System.Threading.Timer? timer;
    private bool building;
    private bool dirty;
    private bool disposed;

    public ProjectWatcher(string projectDir, Action<string> log, Action reload)
    {
        ProjectDir = projectDir;
        this.log = log;
        this.reload = reload;
        watcher = new FileSystemWatcher(projectDir)
        {
            IncludeSubdirectories = false,
            NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.Size,
        };
        watcher.Changed += OnChanged;
        watcher.Created += OnChanged;
        watcher.Deleted += OnChanged;
        watcher.Renamed += OnRenamed;
        watcher.Error += (_, e) => this.log($"Watcher error for {projectDir}: {e.GetException().Message}");
        watcher.EnableRaisingEvents = true;
    }

    public string ProjectDir { get; }

    public void Dispose()
    {
        lock (gate)
        {
            disposed = true;
            timer?.Dispose();
            timer = null;
        }

        watcher.EnableRaisingEvents = false;
        watcher.Dispose();
    }

    /// <summary>The files a build reads: the sources and the manifest, at the top of the folder (ARCHITECTURE.md section 3); out/ is never watched.</summary>
    private static bool IsSource(string? name) =>
        name is not null
        && (name.EndsWith(".bas", StringComparison.OrdinalIgnoreCase)
            || name.EndsWith(".cls", StringComparison.OrdinalIgnoreCase)
            || name.EndsWith(".frm", StringComparison.OrdinalIgnoreCase)
            || name.Equals("vbang.json", StringComparison.OrdinalIgnoreCase));

    private void OnChanged(object sender, FileSystemEventArgs e)
    {
        if (IsSource(e.Name))
        {
            Schedule();
        }
    }

    private void OnRenamed(object sender, RenamedEventArgs e)
    {
        if (IsSource(e.Name) || IsSource(e.OldName))
        {
            Schedule();
        }
    }

    private void Schedule()
    {
        lock (gate)
        {
            if (disposed)
            {
                return;
            }

            if (building)
            {
                dirty = true;
                return;
            }

            timer ??= new System.Threading.Timer(_ => Build(), null, Timeout.InfiniteTimeSpan, Timeout.InfiniteTimeSpan);
            timer.Change(Debounce, Timeout.InfiniteTimeSpan);
        }
    }

    private void Build()
    {
        lock (gate)
        {
            if (disposed || building)
            {
                return;
            }

            building = true;
            dirty = false;
        }

        BuildOutcome outcome;
        try
        {
            outcome = ProjectBuilder.EnsureBuilt(ProjectDir);
        }
        catch (Exception ex)
        {
            log($"Hot reload: build of {ProjectDir} threw {ex.GetType().Name}: {ex.Message}");
            outcome = BuildOutcome.Failed;
        }

        switch (outcome)
        {
            case BuildOutcome.Failed:
                log($"Hot reload: build failed for {ProjectDir}:{Environment.NewLine}{ProjectBuilder.LastOutput}");
                break;
            case BuildOutcome.NoCli:
                log(ProjectBuilder.LastOutput);
                break;
            default:
                // The swap touches Excel objects and registrations, so it runs on Excel's thread, after any running macro.
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    try
                    {
                        reload();
                    }
                    catch (Exception ex)
                    {
                        log($"Hot reload of {ProjectDir} failed: {ex.GetType().Name}: {ex.Message}");
                    }
                });
                break;
        }

        lock (gate)
        {
            building = false;
            if (dirty && !disposed)
            {
                dirty = false;
                timer?.Change(Debounce, Timeout.InfiniteTimeSpan);
            }
        }
    }
}
