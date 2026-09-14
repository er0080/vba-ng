# CLAUDE.md

vba-ng is an Excel-DNA add-in plus a `vbang` CLI that replaces native VBA with a compiled,
git-native, 100% compatible implementation. ARCHITECTURE.md is the design reference: the current
design, in the present tense. Read it before any non-trivial change.

Every doc has one job, and writing in the wrong one is a defect: see Documentation below. Status is
ROADMAP.md's table; the reasoning behind a change is docs/build-log.md.

## Commands

```
dotnet build                          build everything; warnings are errors
dotnet test                           unit + golden replay tests; never needs Excel
$env:VBANG_CORPUS = "C:\path\to\vba"; dotnet test    also runs the corpus tests: the corpus parse rate and the compile scorecard; the reports land next to the test dll as corpus-report.txt and corpus-compile-report.txt
$env:VBANG_E2E = "1"; dotnet test tests/VbaNg.E2E     end-to-end tests (Category=E2E): starts hidden throwaway Excel instances with the add-in from the build output and quits them; skipped without the variable; with VBANG_CORPUS set too, the corpus spec workbooks are imported and their suites run under vbang.Test; the benchmarks run only in the Release build (`-c Release`) and land next to the test dll as benchmark-report.txt
dotnet format --verify-no-changes     style gate
dotnet run --project tests/VbaNg.Golden.Regen -- [Area ...]     regenerate goldens from real VBA in a hidden throwaway Excel; review the diff before committing. --validate checks that the cases parse without starting Excel; --bisect names a case the VBE rejects
dotnet test tests/VbaNg.Golden        replay the goldens against the runtime; the pass rate per area lands next to the test dll as golden-report.txt, known gaps live in tests/VbaNg.Golden/Expected
dotnet run --project src/VbaNg.Cli -- build samples/Hello/Hello.vbang      the CLI, from source
dotnet run --project src/VbaNg.Cli -- import Book.xlsm     read a workbook's VBA project into Book.vbang next to it, with no Excel; --to-xlsx also writes the macro-free copy
dotnet run --project src/VbaNg.Cli -- run Hello.Main --project samples/Hello/Hello.vbang
dotnet run --project src/VbaNg.Cli -- test --project samples/Procedural/Procedural.vbang --junit out/tests.xml     run the '@Test procedures in the running Excel; exit code 5 when one fails
powershell -ExecutionPolicy Bypass -File tools/Start-ExcelWithAddIn.ps1     start a throwaway Excel with the add-in loaded
powershell -ExecutionPolicy Bypass -File tools/New-SheetsSampleWorkbook.ps1     recreate samples/Sheets/Sheets.xlsx (a Form button and an ActiveX button) in a throwaway Excel
powershell -ExecutionPolicy Bypass -File tools/New-QuickstartWorkbook.ps1     recreate samples/Quickstart/Quickstart.xlsx (one Form button), the workbook the README quickstart walks through
powershell -ExecutionPolicy Bypass -File tools/New-ImportFixtureWorkbook.ps1     recreate tests/VbaNg.Import.Tests/Fixtures/Legacy.xlsm (a VBA project the importer is tested against)
powershell -ExecutionPolicy Bypass -File tools/Invoke-Gate.ps1     the whole gate: what CI runs, then the Excel half it cannot. -SkipExcel stops where CI stops, -Benchmarks adds the Release benchmarks, -Package builds the release zip and installs it in a hidden Excel
powershell -ExecutionPolicy Bypass -File tools/New-ReleasePackage.ps1     the release zip and its SHA256SUMS under artifacts/, as release.yml builds them; -Version sets what a tag would, -NoBuild packages the last Release build
powershell -ExecutionPolicy Bypass -File tools/Test-VersionSingleSource.ps1     check that the version is written down in one place only; CI passes -Expected the version it built
powershell -ExecutionPolicy Bypass -File tools/Update-Measurements.ps1     write the golden numbers in docs/measurements.md from golden-report.txt; -Verify fails instead, which is what CI runs
```

While an Excel instance has the add-in loaded, `dotnet build` of the solution fails to overwrite
the add-in's output. Quit that Excel first, or build only the project you changed. The VS Code
F5 configurations run tools/Prepare-Debug.ps1 first: it closes throwaway Excel instances, meaning
any that loaded the add-in from src/VbaNg.AddIn/bin, and never touches other Excel instances.

Tests run under Microsoft.Testing.Platform (global.json "test.runner"). Never pass MSBuild-style
flags such as `-nologo` to `dotnet test`: they are forwarded to the test host, which rejects them
and reports "Zero tests ran".

## Repository layout

```
README.md            landing page: what vba-ng is, why, the quickstart, links out
ARCHITECTURE.md      the current design, present tense; no history
ROADMAP.md           one table of past and planned work, a row per version
CLAUDE.md            this file: rules and commands
src/VbaNg.Compiler   lexer, parser, binder, lowering, C# emitter, Roslyn driver. No COM, no Excel-DNA.
src/VbaNg.Runtime    Variant, arrays, strings, intrinsic library, Err, host services. No Excel-DNA.
src/VbaNg.Interop    type library reader, IDispatch invoke, VARIANT marshaling, event sinks.
src/VbaNg.AddIn      the .xll. Depends on everything above.
src/VbaNg.Cli        vbang. Depends on Compiler, Import, and Interop (COM automation only).
src/VbaNg.Import     vbaProject.bin reader (MS-OVBA) and the source writer; no COM, no Excel.
tests/               one test project per src project, plus VbaNg.Golden (cases, goldens, replay), VbaNg.Golden.Regen (records goldens in Excel), and VbaNg.E2E
samples/             sample projects used by E2E tests and docs
docs/                user-manual.md, build-log.md, measurements.md, vba-quirks.md, diagnostics.md
```

Dependency direction is enforced by project references. Compiler and Runtime never reference
Interop, AddIn, or Excel-DNA. Generated code references only Runtime's public surface.

## Documentation

Each file has one job. Detail that does not fit its job goes to docs/build-log.md — never into
ARCHITECTURE.md or ROADMAP.md, which is how both became build logs.

| File | Holds | Never holds |
|---|---|---|
| `CLAUDE.md` | This file: the coding rules and conventions, and the commands | design detail, status, history |
| `README.md` | What vba-ng is, why it exists, the quickstart, and links out | a list of what works — the docs assume it all works, and vba-quirks.md says where it does not |
| `ARCHITECTURE.md` | The current design only: structure, components, data flow, the decisions in force | history, dated notes, anything that reads as a build log |
| `ROADMAP.md` | One table: a row per `MAJOR.MINOR.PATCH`, from `1.0.0-alpha1` to the last planned item | prose, rationale, per-item notes, status paragraphs |
| `docs/user-manual.md` | End-user usage in full: every command, option, workflow, and error a user meets | design rationale, internals |
| `docs/build-log.md` | The record: what changed, why, what was decided and what was rejected. Newest first | anything a user needs, which is the manual's; a current number, which is measurements.md's |
| `docs/measurements.md` | What the suites measure and last measured, each number with the command and report file behind it | reasoning about a number, which is the build log's |
| `docs/vba-quirks.md` | Where Excel's VBA and MS-VBAL disagree, with how each was verified, and the known gaps where vba-ng and Excel's VBA still differ | — |
| `docs/diagnostics.md` | Every diagnostic id, severity, meaning, and a triggering snippet | — |

- **A doc's scope is its contract.** Write each thing where the table says it goes. Detail
  worth keeping that fits nowhere else is build-log material, not a new heading elsewhere.
- **Terse, or it is not done.** A compiler is judged on its docs; a wordy doc is a broken one.
  Cut every sentence that restates the last one, every hedge, every phrase that survives its own
  deletion. This applies hardest to docs/build-log.md, where volume is the failure mode: an entry
  is a few lines, not a page.
- **ARCHITECTURE.md is present tense.** A reader must not be able to tell from it what
  changed when, or in which version. A decision in force is stated as design; the dated
  reasoning for it lives in docs/build-log.md.
- **ROADMAP.md stays a table.** A version's row says in a phrase what it delivers, and links out
  for the background rather than explaining it.
- **vba-quirks.md and diagnostics.md are updated in the same commit** as the behavior they
  describe: a quirk with how it was verified, a diagnostic with the snippet that triggers it.

## Rules

Every rule is meant to be checkable with a yes or no. Cite one by its name — "bug-for-bug", "a
failing test first" — never by a number.

The `R1`/`D12`/`WP4`/`M7` labels are discontinued from 1.0.0-alpha1: they turned every doc and
comment into a lookup. Do not create new ones. The ones already in source comments stay rather than
churn a thousand files, and docs/build-log.md keeps the old identifiers as historical anchors so
those comments still resolve to something.

### Compatibility

- **Real VBA is the specification.** MS-VBAL is the map. When the map and Excel's VBA
  disagree, Excel wins, and the discrepancy is recorded in docs/vba-quirks.md.
- **No semantic claim without a golden.** Every runtime behavior (operators, coercions,
  intrinsic functions, error numbers, error messages) is backed by a golden case generated from
  real VBA in tests/VbaNg.Golden. If a golden cannot be generated right now, write the case,
  mark the test pending-golden, and do not guess.
- **Bug-for-bug.** Never improve, fix, or modernize VBA semantics. Every valid VBA program
  means exactly what it means in Excel. Extensions (vba-mp) are opt-in per project and never
  change the meaning of existing code.
- **Cite the spec.** Code implementing a semantic rule carries a comment naming the
  MS-VBAL section, for example `// MS-VBAL 5.6.9.3 let-coercion`.
- **Error numbers and message text match VBA exactly.** Code in the wild checks both.

### Clean room

- **No GPL-derived code or grammar.** Do not read Rubberduck or other GPL sources for
  implementation guidance. Allowed references: MS-VBAL, MS-OVBA, MS-OFORMS, Excel-DNA docs,
  Microsoft's Excel object model docs.
- **Dependencies are minimal and permissively licensed** (MIT, Apache-2.0, BSD, MS-PL).
  Adding one needs a sentence of justification in the commit message and an entry in
  Directory.Packages.props. Current set: Excel-DNA, Microsoft.CodeAnalysis (Roslyn), xUnit,
  Verify. FluentAssertions is excluded (license).

### Architecture

- **ARCHITECTURE.md is authoritative.** Code that must deviate from a decision changes the
  doc in the same commit, so that it states the design now in force, and docs/build-log.md
  records what changed and why. Decided questions are not re-decided in code comments or chat.
- **`dotnet build` and `dotnet test` never require Excel**, Office type libraries, or
  admin rights. Excel-dependent tests live in tests/VbaNg.E2E and carry Category=E2E.
- **Generated C# is deterministic and readable.** Same inputs give byte-identical output.
  Identifiers stay recognizable. Every statement carries a `#line`.
- **Diagnostics have stable ids** (`VBA0001` and up), are never renumbered or reused,
  and are listed in docs/diagnostics.md with a snippet that triggers each one. These ids stay:
  they are a user-facing contract, not internal shorthand.

### Testing

- **Three tiers.** Unit tests (no COM) per project; golden replay tests (no Excel);
  E2E (Excel). A change is not done until the first two pass locally.
- **Bug fixes start with a failing test.**
- **Parser coverage.** Every grammar production has a round-trip test: parse then print
  equals the input byte-for-byte. Every diagnostic has a test that triggers it.
- **Snapshot tests (Verify) for generated C#** of representative modules. Snapshot diffs
  are read and explained, never accepted blind.
- **Golden regeneration is deliberate.** Run the regen script, read the diff, explain the
  changes in the commit. Goldens are committed.
- **COM tests run on an STA thread.** Tests that open Excel start a fresh hidden instance,
  use scratch workbooks under the temp directory, never touch a workbook the user has open,
  and always quit the instance they started.
- **A probe must exercise what it claims to test.** `Application.Run` compiles the procedure it
  runs, not the module, so a construct sitting in a procedure nothing calls is never bound and
  looks accepted whatever it is. Put it in, or call it from, the procedure being run.
- **VBA fixture files (.bas, .cls, .frm) are stored with CRLF** and marked in
  .gitattributes so git never normalizes them. Round-trip tests depend on this.

### Code

- **C# latest, .NET 10, x64, nullable enabled, warnings as errors,** analyzers at the
  recommended level, file-scoped namespaces, `dotnet format` clean. Shared settings live in
  Directory.Build.props and .editorconfig.
- **Hot paths do not allocate for scalar cases:** Variant operators, coercion, and
  IDispatch invoke. Everything else is measured before it is optimized.
- **Names.** Namespaces are `VbaNg.*`. The CLI is `vbang`. The Runtime public API is the
  contract with generated code; changes to it are called out explicitly in the commit.
- **A comment says the thing, not where the thing is written down.** Explain the behavior or the
  design in the sentence itself. Naming a doc section is fine; a bare label a reader has to go
  look up is not.
- **A version is read, never written down.** Directory.Build.props declares it once, as
  `VersionPrefix` and `VersionSuffix`, and a release tag overrides it with `-p:Version=<tag>`. Code
  that reports a version reads it off an assembly, and no doc spells one out except ROADMAP.md's
  table, this file, and the build log's record of what shipped. tools/Test-VersionSingleSource.ps1
  fails when either stops being true. NuGet writes the version into every packages.lock.json, so
  after changing the declaration run `dotnet restore` and commit the lock files with it, or CI's
  locked restore fails.
- **Tooling is dotnet-based where possible.** When a script is unavoidable it is
  PowerShell and runs under Windows PowerShell 5.1.

### Workflow

- **Trunk-based on main, small focused commits,** each of which builds and passes unit
  and golden tests. Message format: `scope: summary`, scope in {compiler, runtime, interop,
  addin, cli, import, golden, e2e, docs, build}. The body says why, naming the rule or the
  design it follows in words.
- **Work follows the current version in ROADMAP.md.** No scaffolding or building ahead of it,
  and no speculative abstractions.
- **A version is done when** its exit criterion is met and demonstrated, the tests are in place,
  each doc is updated within its own scope, and the compatibility scorecard is current.
- **Status is ROADMAP.md's table, never ARCHITECTURE.md.** The detail behind a row is
  docs/build-log.md's. After the 1.0.0-rc1 release, GitHub issues and pull requests take over
  from the table for planned work.
- **CI runs everything that does not need Excel; tools/Invoke-Gate.ps1 runs the rest.** GitHub
  runners have no Office, so .github/workflows/ci.yml is build, format, the unit and golden tests,
  the version guard and the measurements check. The E2E tests and the benchmarks are local: run the
  gate script before pushing, and before tagging.
- **A release is a tag.** `v1.0.0-alpha1` pushed to the remote builds that version, gates it, and
  attaches the zip and its checksum to a GitHub release. The tag is the only input; no file is
  edited to cut a release, and nothing is deployed.

### Working with Claude Code in this repo

- **Say what the work serves** before starting anything non-trivial: the version in ROADMAP.md and
  the part of the design in ARCHITECTURE.md it belongs to.
- **Before reporting anything as done:** run `dotnet build`, `dotnet test`, and
  `dotnet format --verify-no-changes`, and report the results verbatim. Failing tests are
  reported as failing.
- **Ask before:** adding a dependency, changing the Runtime public API, regenerating
  goldens, or changing a decision in ARCHITECTURE.md.
- **Commit at natural checkpoints without asking.** Pushing, rewriting history, and anything
  touching a remote still require asking.
- **Excel may be opened autonomously for E2E tests and golden regeneration**, in the throwaway
  hidden instances the COM-test rule describes. Kill only `/automation -Embedding` instances,
  never an Excel the user has open.
- **Prefer editing existing files.** No drive-by refactors or reformatting outside the
  change. No new docs unless the Documentation table names them.
- **Keep this file short.** Rationale and detail go to docs/build-log.md.

## Gotchas

- The VBE exports with CRLF and the system ANSI code page; this project's own files are UTF-8.
- XLL bitness must match Excel's. Verify the dev machine's Office bitness before the first build.
- One .NET runtime version per Excel process. Other .NET add-ins loaded in Excel can conflict.
- Excel rejects COM calls while busy (`RPC_E_SERVERCALL_RETRYLATER`, `RPC_E_CALL_REJECTED`). The
  CLI retries 40 times at 250 ms in VbaNg.Cli/ExcelAutomation.cs. There is no COM message filter;
  tests must expect the same retry.
- `Application.Run` blocks the caller until the macro returns, so a long macro blocks `vbang run`.
  Ctrl+C ends the CLI process without cancelling the macro or killing Excel. Only `vbang logs
  --follow` handles Ctrl+C itself.
- Type library models are cached as JSON under `%LOCALAPPDATA%\vbang\typelibs`, one file per
  library version. Delete a file there to force `vbang build` to read the library again.
- In Windows PowerShell, never pipe a native command (dotnet, vbang, a script) into
  `Select-Object -First`. It stops the pipeline early, kills the process, and reports exit
  code -1. Capture the output in a variable first, then filter it.
