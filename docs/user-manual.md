# vba-ng user manual

## Requirements

| | |
|---|---|
| Excel | desktop Excel for Windows, 64-bit; the add-in is x64 only |
| .NET Desktop Runtime | 10.0, x64, to run a release |
| .NET SDK | 10.0.400 or later, to build from source |
| VS Code | optional, with the C# extension, for breakpoints in VBA source |

Check Excel's bitness before anything else — an `.xll` of the wrong bitness will not load. For a
Click-to-Run install:

```powershell
(Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration').Platform
```

`x64` is what you want. No part of vba-ng needs "Trust access to the VBA project object model" or
administrator rights.

## Getting the add-in and the CLI

**From a release.** `vbang-<version>-win-x64.zip` on the GitHub releases page unpacks to one folder:
`vbang.exe`, the add-in as a single `vba-ng.xll`, `samples\Quickstart`, and these docs. Check the
zip against `SHA256SUMS` and run `Unblock-File` on it before unpacking, or Excel refuses an add-in
that came from the internet. Keep `vba-ng.xll` beside `vbang.exe`: the add-in builds projects with
the CLI next to it. README.md's Quickstart has the commands, putting the folder on `PATH` included.

To upgrade, unpack the new zip beside the old folder, load its `vba-ng.xll` in place of the old one,
and change the `PATH` entry; each bound project builds again the next time its workbook opens. To
remove vba-ng, untick the add-in, then delete the folder, `%LOCALAPPDATA%\vbang`, and the `PATH`
entry.

**From source.**

```powershell
git clone https://github.com/er0080/vba-ng.git
cd vba-ng
dotnet build
```

| | |
|---|---|
| add-in | `src\VbaNg.AddIn\bin\Debug\net10.0-windows\VbaNg.AddIn-AddIn64.xll` |
| CLI | `src\VbaNg.Cli\bin\Debug\net10.0-windows\vbang.exe` |

Put the CLI folder on `PATH` and `vbang` works anywhere; without it, run
`dotnet run --project src/VbaNg.Cli -- <command>`. A release build also writes a single-file
add-in, `bin\Release\net10.0-windows\publish\VbaNg.AddIn-AddIn64-packed.xll`, which the release
zip ships as `vba-ng.xll`.

### Loading the add-in

**File > Options > Add-ins > Manage: Excel Add-ins > Go > Browse**, pick the `.xll`, and accept the
notice Excel shows for an unsigned add-in. If Excel offers to copy the add-in into its own Add-ins
folder, answer **No**: the add-in finds `vbang.exe` beside itself. From source, this script starts a
throwaway Excel with the add-in registered through `Application.RegisterXLL`, which skips the
notice, and prints the process id and the add-in's status:

```powershell
powershell -ExecutionPolicy Bypass -File tools/Start-ExcelWithAddIn.ps1 [-AddInPath <xll>] [-Hidden]
```

An Excel with the add-in loaded holds its files open, so quit it before rebuilding.

## A vba-ng project

### The folder

A project is a folder named after its workbook, in the same directory. That name is the whole
binding: `Sales.xlsx` binds to `Sales.vbang/`.

```
Sales.xlsx              the workbook, macro-free
Sales.vbang/
  vbang.json            manifest: references and document-module kinds (optional)
  .gitignore            out/
  Module1.bas           standard module
  Customer.cls          class module
  Sheet1.cls            document module of the sheet whose CodeName is Sheet1
  ThisWorkbook.cls      document module of the workbook
  out/                  build output, generated
```

A project needs no workbook if you only build, run and test it from the CLI; `samples/Classes` and
`samples/Storage` in the source repository have none.

Source files are exactly what the VBA editor exports: a `.bas` starts with
`Attribute VB_Name = "Module1"`, a `.cls` with the four-line `VERSION 1.0 CLASS` header and then
its attributes. A module is named by its `VB_Name`, or by its file name when the attribute is
absent. The `.bas` and `.cls` files at the top of the folder are compiled in file-name order, then
every `.frm` in its own; subfolders are not read. Files are read as UTF-8 unless a byte order mark
says otherwise, and the VBA editor exports in the system ANSI code page, so re-save an export that
has accented text. A `.frm` compiles as a class and warns `VBA0002`, but only its signatures are
real: the form has no controls and every one of its procedures raises error 438 when called —
`Show`, an event handler and your own helpers alike.

### vbang.json

Optional. A project without one has no references.

| Field | Type | Meaning |
|---|---|---|
| `name` | string | The project's own name. Informational: the compiler and the add-in take the name from the folder |
| `references` | array | What **Tools > References** holds in the VBA editor. Order is kept |
| `documents` | object | CodeName to document kind: `"Workbook"`, `"Worksheet"` or `"Chart"` |

A reference is `{ "name": ..., "guid": ..., "version": ... }`. With a `guid` the library is read
from the registry at that `major.minor` version. With a name alone it must be one vba-ng knows:
`Excel`, `Office`, `Scripting`, `stdole`, `VBIDE`, `MSForms`, or `vbang`, the vba-ng library that
brings the `Assert` module. Anything else is `VBA0023`. `VBA` and `stdole` need no entry.

```json
{
  "name": "Sales",
  "references": [
    { "name": "Excel" },
    { "name": "vbang" },
    { "name": "Scripting", "guid": "420b2830-e718-11cf-893d-00a0c9054228", "version": "1.0" }
  ],
  "documents": { "ThisWorkbook": "Workbook", "Sheet1": "Worksheet" }
}
```

Comments and trailing commas are accepted; field names are case-insensitive. A manifest that is
not valid JSON, or a reference with no name, is `VBA0022` and stops the build.

`documents` matters when the folder travels without its workbook, since the editor's attributes
cannot tell a sheet module from a `PublicNotCreatable` class with a predeclared instance. Present,
the map is the complete list and a `.cls` missing from it is a class; absent, the workbook beside
the folder decides by CodeName; with neither, `VB_PredeclaredId = True` plus `VB_Exposed = True`
does.

### out/

Everything under `out/` is generated by `vbang build` and can be deleted at any time.

| | |
|---|---|
| `<Project>.dll`, `<Project>.pdb` | the compiled project and the symbols breakpoints use |
| `gen/<Module>.cs` | the generated C#, one file per module, with `#line` back to your source |
| `build.json` | the input hashes the add-in compares to decide a rebuild is needed |
| `output.log`, `output.log.1` | what a command, an event handler or a CLI run printed, and the errors they left unhandled |
| `response.json` | a host command's reply when it was too long to return through Excel |

### What goes in git

Commit the sources and `vbang.json`. `out/` is derived: `vbang init` and `vbang import` write a
`.gitignore` containing `out/` for you.

## Commands

```
vbang init   [<workbook>] [--vscode]
vbang import <workbook> [<folder>] [--to-xlsx]
vbang build  [<folder>] [--json]
vbang run    <Module.Procedure> [--project <folder>] [--ui] [--json]
vbang test   [--project <folder>] [--junit <file>] [--ui] [--json]
vbang status [--json]
vbang logs   [--project <folder>] [--follow]
vbang report [<folder>] [--project <folder>]
vbang --version
```

`vbang`, `vbang help`, `vbang -h` and `vbang --help` print that usage, and `vbang --version` (also
`-v`, `version`) prints this build's version alone. Options take
`--name value` or `--name=value`; an unknown option is an error and none is abbreviated. `--json`
is accepted everywhere but only changes the output of `build`, `run`, `test` and `status`.

**Finding the project.** `build`, `run`, `test`, `logs` and `report` need a project folder. All
take `--project <folder>`; `build` and `report` also take it as the first positional argument, and
`--project` wins if both are given. With neither, the current folder is used when its name ends in
`.vbang`, otherwise the single `*.vbang` folder inside it — zero or several, and you are asked to
name one.

**Where each runs.** `init`, `import` and `build` need no Excel. `run`, `test` and `status` call
into an Excel that already has the add-in loaded, found in the Running Object Table or by walking
its workbook window, which needs at least one open workbook. There is no way to choose between two
running instances. A call Excel rejects as busy is retried 40 times, a quarter second apart.

### init

Starts a project beside a workbook — the named one, or the only `.xlsx`, `.xlsm`, `.xlsb`, `.xlam`,
`.xltx` or `.xltm` in the current folder. Writes `<Workbook>.vbang/vbang.json` referencing Excel
and Office, and a `.gitignore` for `out/`. It never overwrites an existing manifest. `--vscode`
also writes, next to the workbook, `.vscode/tasks.json` with a default build task running
`vbang build` under the `$msCompile` problem matcher, and `.vscode/launch.json` with an **Attach to
Excel** configuration.

### import

Reads the VBA project out of a macro workbook and writes it as source files. See
[From a macro workbook to a vba-ng project](#from-a-macro-workbook-to-a-vba-ng-project).

### build

Compiles the project. Diagnostics print one per line in the format every editor
turns into a clickable problem, then a final line:

```
C:\work\Sales.vbang\Module1.bas(14,5): error VBA0005: Variable not defined: 'totl'.
Build failed: 1 error(s).
```

On success the last line is `Build succeeded: <path to the dll>`. Warnings do not fail the build.
Every id is listed in [diagnostics.md](diagnostics.md). `--json` prints
`{ success, project, assembly, diagnostics: [{ id, severity, message, file, line, column }] }`.

### run

Builds the project, then runs one procedure in the running Excel and prints its `Debug.Print`
output. The procedure must be public and take no parameters; module and procedure names are matched
case-insensitively. The call blocks until the macro returns — Ctrl+C ends the CLI, not the macro.

```
> vbang run Greeting.SayHello --project Quickstart.vbang
Hello, world
```

A run-time error the code leaves unhandled prints to stderr, followed by the stack, whose frames name
the module file and line, reaches the project's log, and exits 4. With `--ui` it also shows VBA's
run-time error dialog.

```
vbang: Run-time error '11': Division by zero
```

`--ui` keeps dialogs interactive; see [Dialogs](#dialogs-errors-and-the-log). `--json` prints the
host's reply: `{ ok, output, error, detail }`.

### test

Builds the project, then runs its `'@Test` procedures in the running Excel. See [Tests](#tests).

### status

Asks the add-in what it is doing, on one line (wrapped here): its version, the collation the
process runs with, the loaded projects, the bound workbooks, how many hot reloads have happened,
how the startup run went, and the last 50 lines printed — which is where output from a button click
or a sheet event shows up.

```
vba-ng <version>; collation: NLS; loaded projects: Quickstart (C:\work\Quickstart.vbang);
bound workbooks: Quickstart.xlsx -> C:\work\Quickstart.vbang; reloads: 0;
startup run: ok (Greeting.SayHello); recent output: Hello, world
```

### logs

Prints the project's `out/output.log`. `--follow` keeps printing what arrives until Ctrl+C, and
starts over when the log rolls. With no log yet, it says so and exits 0.

### report

Writes one zip for a bug report. See [Reporting a bug](#reporting-a-bug).

### Exit codes

| Code | Meaning |
|---|---|
| 0 | success |
| 1 | compile errors; also an import that could not read the workbook |
| 2 | usage error: bad option, missing argument, no project folder, no workbook; an import that would overwrite |
| 3 | Excel not reachable, or the add-in did not answer |
| 4 | run-time error in the macro, or the host command itself failed |
| 5 | a test failed |

## From a macro workbook to a vba-ng project

`vbang import` reads `vbaProject.bin` directly, with no Excel, no COM, and no trust setting. The
workbook is never modified.

```
> vbang import Book.xlsm --to-xlsx
Imported 4 file(s) into C:\work\Book.vbang
  Module1.bas
  ThisWorkbook.cls
  Sheet1.cls
  vbang.json
Wrote C:\work\Book.xlsx
```

It writes one file per module into `Book.vbang/`, or the folder you name, in the editor's export
format with CRLF line ends, plus a `vbang.json` (the project's name, `Excel`, every other referenced
library with a GUID, and the `documents` map) and a `.gitignore` for `out/`. Each `.frm` warns that
only its code came over; the layout and `.frx` resources stay in the workbook. `.xls` and `.xla`
workbooks are read too.

An import never overwrites. It refuses a folder that already holds a project, and with `--to-xlsx`
a `Book.xlsx` that already exists, before writing anything; to import again, import into another
folder and compare.

`--to-xlsx` also saves the macro-free copy beside the workbook, through a hidden Excel of its own
with macros disabled, so none of the workbook's code runs. It is the only part of an import that
needs Excel installed, and exit 3 when it cannot be done.

Then `vbang build Book.vbang`, fix what the compiler reports, and open the `.xlsx` in an Excel with
the add-in loaded. Keep the original `.xlsm` out of the way: a workbook that still has its own VBA
project is **not** bound, because both would run. vba-ng says so, in a dialog when Excel is visible,
and says to save a copy as an Excel Workbook (.xlsx) instead.

## Inside Excel

### Binding

When the add-in loads, and on every workbook that opens afterwards, it looks for a folder named
after the workbook. With no folder, nothing happens. With one, it:

1. runs `vbang build` out of process if `out/build.json` does not match the sources or another version
   of vba-ng wrote it, and waits, with Excel not responding meanwhile;
2. loads `out/<Project>.dll` into a load context of its own;
3. sets `Me` in each document module to the workbook or the sheet whose CodeName matches the module
   name, and connects the event procedures;
4. registers the project's public procedures with Excel;
5. watches the folder for changes;
6. runs `Workbook_Open` and every public `Auto_Open`.

On close it runs every public `Auto_Close`, disconnects the events and stops watching the folder.
The assembly stays loaded until the add-in shuts down, and the names the project registered with
Excel are never unregistered, so they outlive the workbook ([vba-quirks.md](vba-quirks.md)).

Step 1 needs a `vbang.exe`: `VBANG_CLI`, else one beside the `.xll`, else this repository's build
output. Without one the add-in says in `vbang status` that it is using whatever is in `out/`.

### Buttons and events

Document modules work as in VBA. A `.cls` named `Sheet1` is the module behind the sheet whose
CodeName is `Sheet1`: `Me` is the sheet, unqualified `Range` and `Cells` mean that sheet's, and
`Worksheet_Change` and its kin run. `ThisWorkbook.cls` is the workbook's module, and
`Private WithEvents` variables work for anything with events. A document module whose CodeName the
workbook does not have is `VBA0024`: it compiles as a class and its event procedures never run.
An ActiveX control's handlers are named `<ControlName>_<Event>` in the sheet's module, as in VBA.

A Form control button runs its macro through `OnAction`. Set it to `Proc` or `Module.Proc` — public
parameterless Subs of standard modules are registered under both names, so `OnAction`,
`Application.Run`, `Application.OnTime` and `OnKey` resolve them. If two loaded projects claim the
same bare name, the second registers `Project.Name` instead and logs a warning.

The runtime resolves `Application.Run` of the project's own procedures in process, arguments and
all, `Book.xlsx!Module.Proc` included when the prefix names the project's own workbook. A prefix
naming another workbook does not resolve.

### Worksheet functions

Every public `Function` of a standard module registers as a worksheet function, under its own name
and `Module.Name`, in the function wizard under a category named after the project.

| A formula passes | The function receives |
|---|---|
| a number, text, a Boolean | `Double`, `String`, `Boolean` |
| an empty cell | `Empty` |
| a left-out argument | `Missing` |
| an error value | `CVErr` of 2000 plus Excel's error code |
| a reference | the `Range` itself for a `Variant`, `Object` or `Range` parameter; the cell's value for a typed one |
| an array | a 1-based two-dimensional `Variant` array |

What comes back: a number, text, a Boolean, a date as its serial, an error value of 2000 plus an
Excel error code as that error, a two-dimensional array as a block of cells, a one-dimensional array
as a row, an object as its default value, `Empty` as an empty value. `Nothing`, `Null`, a
user-defined type, an unallocated array, an array of three dimensions or more, any other error
value, and an unhandled run-time error all show `#VALUE!`, with no dialog. A function's
`Debug.Print` output and its unhandled error reach `vbang status`, not `out/output.log`.
`Application.Volatile`, `Application.Caller` and `Application.ThisCell` answer inside a function
call; every function runs on Excel's main thread and may touch the object model.

### Hot reload

Save a `.bas`, `.cls`, `.frm` or `vbang.json` at
the top of the folder and half a second later the add-in rebuilds on a worker thread, then swaps
the new assembly in on Excel's thread, between macros. `Workbook_Open` does not run again, and
module-level variables and `Static` locals reset, as they do in VBA after an edit. A failed build
shows in `vbang status` and the loaded version stays, so the workbook keeps working; an edit during
a build queues one more.

### Dialogs, errors and the log

Code triggered from inside Excel — a button, an event, `OnTime` — shows real dialogs. Under
`vbang run` and `vbang test` the same code is non-interactive so an unattended run cannot hang
behind a modal window: `MsgBox` prints its text and returns `vbOK`, `InputBox` returns its default,
`Beep` and `SendKeys` print what they would have done. `--ui` restores all of them.

An error left unhandled by a command or an event procedure gets VBA's own dialog — title
**Microsoft Visual Basic**, text `Run-time error '5':` and the description, with **End** as the
only working button — and a line in the log:

```
2026-09-12 11:04:19.286 Run-time error in Module1.Fails: '5': custom
```

Each line of `out/output.log` is stamped with the local time; a worksheet function's output, and the
add-in's own binding and reload lines, go to `vbang status` instead. Past one megabyte the log moves
to `output.log.1`, replacing the previous one. A log that cannot be written is skipped rather than
failing the run.

## Tests

Mark a public Sub with no parameters with a `'@Test` comment above it (`Rem @Test` works too;
anything after the word is free text). The annotation anywhere else is `VBA0021`.

`Assert` needs `{ "name": "vbang" }` in the manifest's references. Without that reference the name
`Assert` means whatever it means in your own code, so existing code never changes meaning; with it,
a module or variable of your own called `Assert` still wins. Both `Assert.X` and `vbang.Assert.X`
are accepted.

| Assertion | |
|---|---|
| `Assert.AreEqual expected, actual[, message]` | passes when `expected = actual` under the calling module's `Option Compare`, when both are `Null`, or when both are the same object |
| `Assert.AreNotEqual notExpected, actual[, message]` | the opposite |
| `Assert.IsTrue condition[, message]`, `Assert.IsFalse` | as `If` would take the condition; `Null` counts as False |
| `Assert.IsNothing value[, message]`, `Assert.IsNotNothing` | `IsNotNothing` also fails for a value that is not an object reference |
| `Assert.Fail [message]` | fails the test outright |

A failed assertion ends the test at once and is not a VBA error: no `On Error` handler in the test
can absorb it, so a test may run `On Error Resume Next` to inspect `Err.Number` and still fail on
its assertion. Each test starts with `Err` cleared. Tests run in module order, then declaration
order.

```
> vbang test --junit out/tests.xml
passed  Shapes.MovePoint_AddsTheOffset
passed  Shapes.Perimeter_SumsEverySide
FAILED  Shapes.Distance_IsEuclidean
        Assert.AreEqual failed. Expected:<5>. Actual:<4.9>.
        | probe: 3, 4
3 tests: 2 passed, 1 failed, 0 errors. 63 ms.
```

A failed assertion is reported as a failure, an unhandled run-time error as an error, each with the
test's own `Debug.Print` lines under it behind `|`. The exit code is 5 when any test fails.

`--junit <file>` writes JUnit XML — one `testsuite` per module, one `testcase` per test, a
`failure` of type `AssertFailed` or an `error` of type `RuntimeError`, and the test's output as
`system-out`. Excel writes the file itself, so the path is on the same machine as Excel.

## VS Code

`vbang init --vscode` writes the two files a debugging session needs: a build task and an **Attach
to Excel** configuration. Add `.vscode/settings.json` yourself — VS Code allows breakpoints only in
languages that register a debugger, and nothing registers one for VBA:

```json
{
  "debug.allowBreakpointsEverywhere": true,
  "files.associations": { "*.bas": "vb", "*.cls": "vb", "*.frm": "vb" },
  "[vb]": { "files.eol": "\r\n" }
}
```

Then start Excel with the add-in, open the workbook, attach to `EXCEL.EXE`, and set a breakpoint in
the gutter of a `.bas`. It binds: the PDB `vbang build` writes maps IL back to your VBA lines, so
any .NET debugger gives you breakpoints, stepping, a call stack, and a locals window that reads
`Variant`, arrays, objects, records and `Date` as VBA's own does.

To have Excel start and run a procedure under one keypress, write a `coreclr` launch configuration
that runs `EXCEL.EXE` with the `.xll` and a workbook as its arguments, and sets `VBANG_PROJECT` to
the project folder and `VBANG_RUN` to `Module.Procedure`. The add-in waits for Excel to be
ready — main window visible, splash screen gone, `Application.Ready`, a workbook open — polling for
up to a minute, then runs it, with `Debug.Print` reaching the debugger's console. Pass a workbook:
without one Excel sits on its Start screen and the queued run waits. `vbang status` reports how the
startup run went, and the source repository's `.vscode/launch.json` has working examples.

Highlighting is Visual Basic's grammar; there are no completions, hover, rename, or go-to-definition.

## Troubleshooting

| What you see | What to do |
|---|---|
| `no running Excel instance found` | Start Excel, load the add-in, open a workbook |
| `Excel could not run 'vbang.Run'. Is the vba-ng add-in loaded?`, or a host command returning nothing | The instance answered but has no add-in: check **File > Options > Add-ins**. With `Call was rejected by callee` in parentheses, Excel is busy: leave cell editing with Esc, or close the dialog it shows |
| Excel refuses `vba-ng.xll`, or the add-in is unticked after a restart | Excel must be 64-bit, and `dotnet --list-runtimes` must list `Microsoft.WindowsDesktop.App 10.0`. Run `Unblock-File` on the `.xll`; in **File > Options > Trust Center > Add-ins**, add-ins must not be disabled or required to be signed. Excel runs one .NET version, so another .NET add-in can keep this one out |
| A button says `Cannot run the macro` | The workbook is not bound. `vbang status` lists the bound workbooks and, under recent output, a `Build failed for` or `Could not load` line; `vbang build <folder>` prints the diagnostics |
| A workbook in a OneDrive or SharePoint folder is never bound | Excel can report such a workbook's path as a web address, where no folder is beside it. Open it from a local folder |
| `No project folder found` / `Several *.vbang folders found here` | Run inside the `.vbang` folder or pass `--project <folder>` |
| `Could not load ...` / `Project has not been built; expected ...\out\Sales.dll` | Nothing is in `out/`. Run `vbang build`, or make `vbang.exe` reachable so the add-in can |
| `Sales.xlsm still has a VBA project, so vba-ng does not bind Sales.vbang` | **File > Save As**, Excel Workbook (.xlsx), and open that copy instead |
| `no sheet with CodeName Sheet3; its module is not bound` | The `.cls` names a CodeName the workbook does not have. Rename the module, or list it in `documents` |
| `VBA0024` | Same, at build time |
| `SayHello is already registered by another project` | Two loaded projects claim one name; the second answers to `Project.SayHello` |
| `vbang CLI not found (set VBANG_CLI)` | The add-in cannot rebuild a stale project. Put `vbang.exe` beside the `.xll`, or set `VBANG_CLI` with `[Environment]::SetEnvironmentVariable('VBANG_CLI', '<folder>\vbang.exe', 'User')` and restart Excel |
| `The Excel type library could not be read; workbook events will not be connected` | The Excel type library is not registered for this Excel; reinstall or repair Office |
| `VBA0023 Can't find project or library` | The manifest names a library that is not registered, or a bare name vba-ng does not know. Add its `guid` |
| `already exists; init never overwrites a project` | Edit the manifest by hand, or delete it first |
| A rebuilt add-in will not overwrite | An Excel still has the old one loaded. Quit it |
| A UDF shows `#VALUE!` | A run-time error inside it, as in VBA. `vbang status` has the error; the project log does not |
| Stale type-library behavior after an Office update | Delete the file for that library under `%LOCALAPPDATA%\vbang\typelibs` and rebuild |

Surprising behavior, and where vba-ng still differs from Excel, is in [vba-quirks.md](vba-quirks.md).

## Reporting a bug

`vbang report` writes `vbang-report-<Project>-<date>-<time>.zip` in the current folder and prints
its path, its contents, and that the workbook is not among them:

| Entry | |
|---|---|
| `versions.txt` | vbang and .NET versions, OS and bitness, culture, the project path, whether the build succeeded, and the add-in's status line |
| `build.log` | the diagnostics of a build run right then, not whatever was left over |
| `vbang.json`, `source/*` | the manifest and every `.bas`, `.cls` and `.frm` |
| `out/build.json`, `out/gen/*.cs`, `out/output.log` | the build record, the generated C#, and the project's log |
| `status.json` | the add-in's status, when Excel was reachable |

File it at https://github.com/er0080/vba-ng/issues: attach the zip, give your Excel version and build
(**File > Account > About Excel**), say what you expected, and say what real Excel does with the
same code. Read the zip first: it holds all your source, the log, and the project's path.

## Files and environment variables

| | |
|---|---|
| `<Project>.vbang/out/` | build output; safe to delete |
| `%LOCALAPPDATA%\vbang\typelibs\<guid>-<major>.<minor>.json` | cached type library models, one per library version |
| `%TEMP%\vbang\shadow\` | copies of loaded assemblies, so a rebuild can overwrite `out/` |
| `VBANG_CLI` | the `vbang.exe` the add-in rebuilds a stale project with |
| `VBANG_PROJECT`, `VBANG_RUN` | the project folder and `Module.Procedure` the add-in runs at startup |

