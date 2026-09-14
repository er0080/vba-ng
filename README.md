# vba-ng

**VBA, compiled. Your macros in git, your workbook macro-free, your editor whichever one you
already use.**

vba-ng is an Excel add-in — an Excel-DNA `.xll` — plus a `vbang` command line tool that replaces
the VBA engine. Your code lives in ordinary text files next to the workbook, is compiled to .NET
IL by a real compiler, and runs inside Excel as VBA does.

This is not a new language, a subset, or a modernized dialect. It is VBA, bug for bug, with the
tooling VBA never had; [docs/vba-quirks.md](docs/vba-quirks.md) lists where vba-ng still differs.

> **Pre-release.** [ROADMAP.md](ROADMAP.md) lists the versions; `vbang --version` reports the build
> you have.

## Why

VBA's problems are not in the language. They are in everything around it: code sealed inside a
binary workbook, an editor from 1998, no diff, no merge, no branches, no test runner, no
continuous integration, and a macro-enabled file that mail gateways and security policies
increasingly refuse to carry. Rewriting a working spreadsheet in Python or C# means rewriting the
business logic too, which is where the risk is.

**Text files, one folder per workbook.** `Sales.xlsx` binds to `Sales.vbang/` in the same
directory, holding `.bas` and `.cls` files in exactly the format the VBA editor already exports.
Git sees ordinary text: diffs, merges, blame, code review, with no export step.

**The workbook stays macro-free.** A plain `.xlsx`: nothing to warn about, strip at a gateway, or
prompt on open.

**A real compiler, not a translator.** A hand-written lexer, parser, binder and code generator
built from the MS-VBAL specification, emitting C# with `#line` directives that Roslyn compiles to
IL and a portable PDB. Transpiling to VB.NET was rejected: there, `Integer` is 32 bits,
`Currency` does not exist, arrays are zero-based, and error numbers differ. Owning the front end
and the runtime is the only route to exact compatibility.

**Compatibility is measured, not claimed.** Every semantic rule is backed by a golden test
generated from real VBA in Excel, then replayed against the runtime with no Excel present. Where
Microsoft's specification and Excel disagree, Excel wins and the discrepancy is written down in
[docs/vba-quirks.md](docs/vba-quirks.md).

**Debugging that already works.** The PDB maps IL back to your `.bas` files, so any .NET debugger
attaches to `EXCEL.EXE` and breaks on a line of VBA, with working locals, stepping and watches.

**Built for a terminal, and for agents.** The interface is the CLI, files and stdout. Diagnostics
come out as `file(line,col): error VBA0001: message`, which every editor turns into clickable
problems, and `build`, `run`, `test` and `status` take `--json`. A coding agent drives vba-ng with
the same commands a person does, with no plugin or protocol in between.

**The compiler never runs inside Excel.** `vbang build` is a separate process producing a DLL and
a PDB; the add-in only loads the output. A compiler bug cannot take down a workbook, and the build
works on a machine with no Excel installed.

## Quickstart

[docs/user-manual.md](docs/user-manual.md) has the rest: every command and option, starting a
project beside a workbook of your own, importing an existing VBA project, the test runner.

**Prerequisites.** Windows with 64-bit Excel 2016 or later, and the
[.NET 10 Desktop Runtime](https://dotnet.microsoft.com/download/dotnet/10.0), x64. The add-in's
bitness must match Excel's: on a Click-to-Run install this prints `x64`, and otherwise
**File > Account > About Excel** says 64-bit. `dotnet --list-runtimes` lists
`Microsoft.WindowsDesktop.App 10.0` once the runtime is installed.

```powershell
(Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration').Platform
```

**1. Install.** Download `vbang-<version>-win-x64.zip` and `SHA256SUMS` from
[Releases](https://github.com/er0080/vba-ng/releases). In PowerShell, in the download folder: check
that the two hashes printed match, unblock the zip so Excel does not refuse an add-in from the
internet, unpack it, put the folder on your `PATH`, and copy its path:

```powershell
(Get-FileHash .\vbang-*-win-x64.zip).Hash.ToLower(); Get-Content .\SHA256SUMS
Unblock-File .\vbang-*-win-x64.zip
Expand-Archive .\vbang-*-win-x64.zip "$env:LOCALAPPDATA\vba-ng"
$dir = (Get-Item "$env:LOCALAPPDATA\vba-ng\vbang-*-win-x64").FullName
[Environment]::SetEnvironmentVariable('Path', [Environment]::GetEnvironmentVariable('Path', 'User') + ";$dir", 'User')
$env:Path += ";$dir"; $dir | Set-Clipboard
vbang --version
```

**2. Load the add-in.** In Excel, **File > Options > Add-ins > Manage: Excel Add-ins > Go >
Browse**, paste the copied path into the file name box, pick `vba-ng.xll`, and accept the notice
Excel shows for an unsigned add-in. If Excel offers to copy the add-in into its own Add-ins folder,
answer **No**: the add-in finds `vbang.exe` beside itself. Excel loads it at every start from then
on.

**3. Click the button.** Open the sample from the same terminal:

```powershell
Invoke-Item "$dir\samples\Quickstart\Quickstart.xlsx"
```

It is a plain `.xlsx` with one button, and beside it `Quickstart.vbang\`, the code bound to it by
name:

```vba
Attribute VB_Name = "Greeting"
Option Explicit

Public Sub SayHello()
    Worksheets("Sheet1").Range("B2").Value = "Hello, world"
    Debug.Print "Hello, world"
End Sub
```

The add-in builds the project as the workbook opens, and Excel does not respond for those few
seconds. Click **Say hello** and `Hello, world` lands in **B2**. If Excel says it cannot run the
macro, `vbang status` says why: the workbook missing from its bound workbooks, or a build error.

**4. Drive it from the terminal.** Leave Excel running:

```powershell
cd "$dir\samples\Quickstart"
vbang run Greeting.SayHello --project Quickstart.vbang
vbang status
```

`run` builds the project, runs the procedure in the Excel you have open, and prints
`Hello, world`. `status` reports the add-in's version, its loaded projects, the workbooks they are
bound to, `Quickstart.xlsx` among them, and the most recent output.

**5. Edit while Excel stays open.** In `Quickstart.vbang\Greeting.bas`, change `"Hello, world"` on
the `Range` line to `"Hello again"` and save. The add-in rebuilds the project half a second later and
reloads it between macros: click the button and **B2** changes, and `vbang status` counts one
reload. An edit that does not compile leaves the loaded version in place, and `vbang status` shows
the error.

## Building from source

With the [.NET 10 SDK](https://dotnet.microsoft.com/download) and git:

```bash
git clone https://github.com/er0080/vba-ng.git
cd vba-ng
dotnet build                       # everything; warnings are errors
dotnet test                        # unit and golden tests; never needs Excel
dotnet format --verify-no-changes  # style gate
```

The CLI and the add-in land where
[docs/user-manual.md](docs/user-manual.md#getting-the-add-in-and-the-cli) says. Excel-dependent
tests run only with `VBANG_E2E` set:

```powershell
$env:VBANG_E2E = "1"; dotnet test tests/VbaNg.E2E
```

**Break in the VBA.** With [VS Code](https://code.visualstudio.com/) and the C# extension, choose
**Debug Quickstart sample in Excel** in the Run panel. One keypress builds the solution, compiles
the sample, starts Excel with the add-in and the workbook, and runs `Greeting.SayHello`. Put a
breakpoint beside the `Worksheets(...)` line in `Greeting.bas` and click the button: execution stops
on that line inside `EXCEL.EXE`, with locals, stepping and a call stack. (If the debugger cannot
find Excel, correct the `program` path in `.vscode/launch.json`.)

## Documentation

| | |
|---|---|
| [docs/user-manual.md](docs/user-manual.md) | Using it: every command, option, workflow and error |
| [ARCHITECTURE.md](ARCHITECTURE.md) | The design: structure, components, data flow |
| [ROADMAP.md](ROADMAP.md) | What is planned, a row per version |
| [docs/vba-quirks.md](docs/vba-quirks.md) | Where Excel's VBA and the specification disagree |
| [docs/diagnostics.md](docs/diagnostics.md) | Every diagnostic id, and what triggers it |
| [docs/build-log.md](docs/build-log.md) | Why the code is the way it is |
| [docs/measurements.md](docs/measurements.md) | What the suites measure: goldens, corpus, benchmarks |
| [CLAUDE.md](CLAUDE.md) | The rules this repository is built under |

## License

[MIT](LICENSE). No GPL-derived code enters this repository; the VBA grammar and semantics here are
built from Microsoft's own specifications (MS-VBAL, MS-OVBA, MS-OFORMS) and from observing Excel.
