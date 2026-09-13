# vba-ng

**VBA, compiled. Your macros in git, your workbook macro-free, your editor whichever one you
already use.**

vba-ng is an Excel add-in — an Excel-DNA `.xll` — plus a `vbang` command line tool that replaces
the VBA engine. Your code lives in ordinary text files next to the workbook, is compiled to .NET
IL by a real compiler, and runs inside Excel exactly as VBA does: buttons, sheet events, worksheet
functions, `MsgBox`, `Range`, COM references, error numbers, and all the rest. The workbook itself
stays a plain `.xlsx` with no VBA project in it.

Existing VBA runs unchanged. This is not a new language, a subset, or a modernized dialect. It is
VBA, bug for bug, with the tooling VBA never had.

> **Pre-release.** This is the first general testing release. [ROADMAP.md](ROADMAP.md) lists the
> versions; `vbang --version` reports the build you have.

## Why

VBA's problems are not in the language. They are in everything around it: code sealed inside a
binary workbook, an editor from 1998, no diff, no merge, no branches, no test runner, no
continuous integration, and a macro-enabled file that mail gateways and security policies
increasingly refuse to carry. Rewriting a working spreadsheet in Python or C# means rewriting the
business logic too, which is where the risk is.

vba-ng keeps the language and replaces everything around it.

**Text files, one folder per workbook.** `Sales.xlsx` binds to `Sales.vbang/` in the same
directory, holding `.bas` and `.cls` files in exactly the format the VBA editor already exports.
Git sees ordinary text: real diffs, real merges, real blame, real code review. No export step, no
synchronization dance, no code buried in a binary.

**The workbook stays macro-free.** Code lives beside the workbook rather than inside it, so the
file stays a plain `.xlsx`. Nothing to warn about, nothing to strip at a gateway, nothing to
prompt the user on open.

**A real compiler, not a translator.** A hand-written lexer, parser, binder and code generator
built from the MS-VBAL specification, emitting C# with `#line` directives that Roslyn compiles to
IL and a portable PDB. Transpiling to VB.NET was rejected outright: there, `Integer` is 32 bits,
`Currency` does not exist, arrays are zero-based, and error numbers differ. Owning the front end
and the runtime is the only route to exact compatibility.

**Compatibility is measured, not claimed.** Every semantic rule is backed by a golden test
generated from real VBA running in real Excel, then replayed against the runtime with no Excel
present. Where Microsoft's specification and Excel disagree, Excel wins and the discrepancy is
written down in [docs/vba-quirks.md](docs/vba-quirks.md).

**Debugging that already works.** The PDB maps IL back to your `.bas` files, so any .NET debugger
attaches to `EXCEL.EXE` and breaks on a line of VBA, with working locals, stepping and watches. No
new debugger to build, none to learn.

**Built for a terminal, and for agents.** The interface is the CLI, files and stdout. Diagnostics
come out as `file(line,col): error VBA0001: message`, which every editor already turns into
clickable problems, and `build`, `run`, `test` and `status` take `--json`. A coding agent drives
vba-ng with the same commands a person does, with no plugin or protocol in between.

**The compiler never runs inside Excel.** `vbang build` is a separate process producing a DLL and
a PDB; the add-in only loads the output. A compiler bug cannot take down a workbook, and the build
works on a machine with no Excel installed at all.

## Quickstart

This ends with a button in a workbook running VBA under vba-ng and a breakpoint you hit in the
`.bas` source. [docs/user-manual.md](docs/user-manual.md) has the rest: every command and option,
starting a project beside a workbook of your own, importing an existing VBA project, the test
runner.

**Prerequisites.** Windows with 64-bit Excel 2016 or later, the
[.NET 10 SDK](https://dotnet.microsoft.com/download), git, and
[VS Code](https://code.visualstudio.com/) with the C# extension for F5 debugging. The add-in's
bitness must match Excel's, so check it:

```powershell
(Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration').Platform
```

It should print `x64`. (32-bit Excel is planned, not supported yet.)

**1. Build.**

```bash
git clone <this repository> vba-ng
cd vba-ng
dotnet build
```

That builds the compiler, the runtime, the `vbang` CLI and the `.xll` add-in. No Excel is needed,
here or for `dotnet test`. The CLI lands in
`src\VbaNg.Cli\bin\Debug\net10.0-windows\vbang.exe`; put that folder on your `PATH`, or use
`dotnet run --project src/VbaNg.Cli --` as below.

**2. Read the sample.** `samples/Quickstart` is a workbook with one button plus the project bound
to it by name:

```
samples/Quickstart/
  Quickstart.xlsx           a normal .xlsx: no VBA project, one Form button on Sheet1
  Quickstart.vbang/         the code for that workbook, found by the matching name
    vbang.json              the manifest: this project's references, as Tools > References
    Greeting.bas            the module, in the format the VBA editor exports
```

```vba
Attribute VB_Name = "Greeting"
Option Explicit

Public Sub SayHello()
    Worksheets("Sheet1").Range("B2").Value = "Hello, world"
    Debug.Print "Hello, world"
End Sub
```

The button's macro is `Greeting.SayHello`, set the way Excel sets any Form control's macro.

**3. Press F5.** In VS Code, choose **Debug Quickstart sample in Excel** in the Run panel. One
keypress builds the solution, compiles the sample, starts Excel with the add-in and the workbook,
and runs `Greeting.SayHello`: `Hello, world` lands in **B2** and in the **Debug Console**. Click
the **Say hello** button on the sheet and the same procedure runs, exactly as it would in VBA.
(If the debugger cannot find Excel, correct the `program` path in `.vscode/launch.json`; it
assumes `C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE`.)

**4. Break in the VBA.** Put a breakpoint in the gutter beside the `Worksheets(...)` line in
`Greeting.bas` and click the button again. Execution stops on that line inside `EXCEL.EXE`, with
locals, stepping and a call stack.

**5. Drive it from the terminal.** Leave Excel running and open a terminal in the repository:

```bash
dotnet run --project src/VbaNg.Cli -- build samples/Quickstart/Quickstart.vbang
dotnet run --project src/VbaNg.Cli -- run Greeting.SayHello --project samples/Quickstart/Quickstart.vbang
dotnet run --project src/VbaNg.Cli -- status
```

`build` compiles the project with no Excel involved. `run` builds, then runs the procedure in the
Excel you have open and prints its `Debug.Print` output. `status` reports the add-in's version,
its loaded projects, the workbooks they are bound to and the most recent output.

**6. Edit while Excel stays open.** Change `Greeting.bas` and save. The add-in watches the project
folder, rebuilds it half a second later, and reloads it between macros; click the button again and
the new code runs. If the edit does not compile, the error is reported and the loaded version stays
in place, so the workbook keeps working.

## Building from source

```bash
dotnet build                       # everything; warnings are errors
dotnet test                        # unit and golden tests; never needs Excel
dotnet format --verify-no-changes  # style gate
```

Excel-dependent tests sit behind an environment variable, so a normal test run never opens Excel:

```powershell
$env:VBANG_E2E = "1"; dotnet test tests/VbaNg.E2E
```

One .NET project per component: `VbaNg.Compiler` (lexer, parser, binder, emitter), `VbaNg.Runtime`
(`Variant`, arrays, strings, the intrinsic library), `VbaNg.Interop` (type libraries, `IDispatch`),
`VbaNg.Import` (the MS-OVBA reader), `VbaNg.Cli` and `VbaNg.AddIn`. The compiler and the runtime
never reference Excel or COM, which is what keeps the build and the tests independent of Office.

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
