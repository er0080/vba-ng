# Build log

What changed, why, and what was decided or rejected. Newest first. This is the only file in the
project that is dated or written in the past tense: the design in force is
[ARCHITECTURE.md](../ARCHITECTURE.md), the plan is [ROADMAP.md](../ROADMAP.md), and the current
numbers are [measurements.md](measurements.md).

The `R1`/`D20`/`WP4`/`M7` identifiers are discontinued from 1.0.0-alpha1, but some 570 citations in
source comments still use them, most as `ROADMAP.md WP4` or `ARCHITECTURE.md D19`. Those labels
resolve here, not in the files they name: the legend below for a work or rule label, Decisions for
a `D` label.

## Anchors

M0 the spike · M1 the MS-VBAL front end · M2 the runtime core · M3 procedural codegen · M4 COM
and the Excel host · M5 class modules · M6 the importer, `Declare`, and the corpus compiling and
running · M7 the memory model (D20) · M8 the alpha · M9+ what became the 1.0.0-rc1 and later rows
of ROADMAP.md. M7's steps: B strings as BSTRs · C arrays and records · D objects with a COM
identity · E compiler storage and addresses · F interop, host, and measurement. WP0 the
2026-09-09 review's defects · WP1 the corpus compiling against the registered libraries · WP2 the
memory model · WP3 the libraries' own suites running under `vbang test` · WP4 host fidelity · WP5
golden growth · WP6 the import round trip · WP7 packaging, intake, and the benchmarks.

The rules, now cited by name in CLAUDE.md: R1 real VBA is the specification · R2 no semantic
claim without a golden · R3 bug-for-bug · R4 cite the spec · R5 error numbers and text match ·
R6 no GPL-derived code · R7 dependencies minimal and permissive · R8 ARCHITECTURE.md is
authoritative · R9 build and test never require Excel · R10 generated C# deterministic and
readable · R11 stable diagnostic ids · R12 three test tiers · R13 a failing test first · R14
parser round-trip coverage · R15 Verify snapshots of generated C# · R16 golden regeneration is
deliberate · R17 COM tests on an STA thread in a throwaway Excel · R18 VBA fixtures stored CRLF ·
R19 C# latest, .NET 10, x64, warnings as errors · R20 measure before optimizing · R21 names, and
the Runtime public API as a contract · R22 tooling is dotnet-based · R23 trunk-based, small
commits · R24 work follows the current version · R25 a version is done when its criterion is
demonstrated · R26 status is ROADMAP.md's · R27 say what the work serves · R28 build, test, and
format before reporting done · R29 ask before a dependency, a Runtime API change, a regeneration,
or a design change · R30 commit at checkpoints, ask before pushing · R31 Excel may be opened for
E2E tests and regeneration · R32 prefer editing existing files · R33 keep CLAUDE.md short · R34 a
doc's scope is its contract · R35 ARCHITECTURE.md is present tense · R36 ROADMAP.md stays a table
· R37 quirks and diagnostics updated in the same commit as the behavior.

## 1.0.0-alpha1 ships when

The five criteria set 2026-09-09, as they stand. Each is shown by a test or a report.

1. **Compile.** Every corpus library builds against the registered libraries, and stdVBA's `src`
   builds at least 90% of its files. *Met*: the compile scorecard in measurements.md.
2. **Run.** VBA-JSON's, VBA-Dictionary's and VBA-UTC's spec suites and VBA-Better-Array's own
   TestRunner pass under `vbang test` in Excel. *Met 2026-09-12*; re-shown by the final gate with
   the corpus configured.
3. **Goldens.** Every miss closed or carried by a decision. *Met*: two named misses, D22.
4. **Real workbooks.** Five go through `vbang import` and `vbang build` and run in Excel with no
   source edits, every divergence filed as an issue before the tag. *Met 2026-09-14*: the fifth is
   VBA-Web's spec workbook, which reaches the network and so runs only in the local gate.
5. **Packaging.** A tester installs from the release zip by following README.md in ten minutes;
   `vbang init` starts a project, an unhandled error shows on screen, and `vbang report` writes the
   bug bundle. *Written and tested 2026-09-14*: README installs from the zip, and
   ReleasePackageTests installs it in Excel. Open: a first tester's timing.
   `vbang install` and `--out` moved to 1.0.0-alpha2, and so did ribbon `onAction` and the MSForms
   argument order, which docs/vba-quirks.md lists as unverified.

## 2026-09-14

- **The fifth real workbook: VBA-Web's spec workbook.** Imported with no edits, its eight suites run
  under VBA and under vbang and agree spec by spec, error lines included. Getting there found three
  divergences, each now a golden area recorded in Excel. A `Class_Terminate` that stores `Me` keeps
  the object's variables; VBA-TDD records every spec that way, and vba-ng released them, so every
  suite said 91. `Line Input` ends a line at CR or CRLF, never at a lone LF. A late-bound call
  passes a parameterless member's extra arguments to the default member of what it returns; the
  rule went into the runtime as `LateBound.PassOn` rather than into every generated class.
- **vba-ng's own notices never block an unattended Excel.** The not-bound notice was a modal box
  even in a hidden automation Excel, and it hung that test's first run for twenty minutes. It opens
  a dialog only in a visible Excel now. `MsgBox` and the run-time error dialog keep VBA's behavior.
- **The release zip is proven by installing it.** One script builds it for release.yml and for a
  local run. ReleasePackageTests unpacks it, loads its `vba-ng.xll`, and opens the Quickstart sample
  with no build output, so the packaged CLI has to build it. The add-in looks for the CLI beside
  itself, which is why the two share a folder. README's Quickstart starts from the zip; a
  `workflow_dispatch` run of release.yml is the dry run.

## 2026-09-12

- **CI is the half that needs no Excel, and a release is a tag.** GitHub-hosted runners have no
  Office, so ci.yml runs build, `dotnet format`, the unit and golden tests, the version guard and a
  check that docs/measurements.md still matches `golden-report.txt`. That is not a compromise: the
  rule that a build and a test never require Excel means the runner enforces the rule. The Excel
  half stays local behind tools/Invoke-Gate.ps1, which runs CI's steps in CI's order and then the
  E2E tests and, on request, the Release benchmarks. A self-hosted runner was considered and
  rejected for now: on a public repository a fork's pull request would run on the machine hosting
  it. release.yml takes the version from the tag (`v1.0.0-alpha1`), gates it, publishes the CLI and
  the packed `.xll`, and attaches the zip and its SHA256 to a GitHub release through `gh`, so no
  third-party action handles the upload. Package lock files are committed and CI restores in locked
  mode, so a dependency change arrives as a reviewable diff.
- **Measurements moved out of this file** into docs/measurements.md, where the golden block is
  generated from the report by tools/Update-Measurements.ps1 rather than typed. This file keeps why
  a number moved.
- **One job per doc.** ROADMAP.md (1,473 lines) and ARCHITECTURE.md's decision section had both
  become build logs; this file is their destination, and each doc got a scope contract and a line
  budget. The numbered `R`/`D`/`WP`/`M` labels were discontinued with 1.0.0-alpha1 — they turned
  every doc and comment into a lookup — while source comments keep theirs.
- **The import round trip (WP6), which is M4's exit criterion through the importer.** The manifest
  now names every document module and its kind, read from the workbook's own parts, so a project
  folder in git without its workbook still says which `.cls` files are document modules (D19). The
  E2E imports the Legacy fixture with `--to-xlsx`, builds, opens the copy in Excel, and reads back
  `Workbook_Open`, `Worksheet_Change`, an ActiveX click, a Form macro, and `=Gross(100)`.
- **UserForms build as a warning (D-D), which the round trip needed first.** A `.frm` joins the
  sources, the binder takes the class header from the extension, and a designer block becomes
  blank lines so every statement keeps its `#line`. Nothing in a form's body is bound and each
  procedure raises 438 naming the form, which is what makes `Show`, `Load`, and a control
  reference raise.
- **Divergences found by driving the criterion-4 workbooks, fixed.** A type library generated for
  a document and registered nowhere falls back to the library its name gives; the `documents` map
  is the whole list, so a predeclared and exposed class is no longer taken for a sheet module;
  `End` and `Exit` keywords are interchangeable and VBA checks neither which one closes a
  procedure; `ws.[Name] = value` assigns through the sheet's `Evaluate`. The compile scorecard was
  unchanged, as it should be: these only show in the copies the workbooks carry.
- **`stdVBA/fullBuild.xlsm` went 14 errors to 1, `VBA-Web - Example.xlsm` 24 to 2.** Ours were
  stdole, which every VBA project references and no manifest names, and an `Optional` default of
  `library.Enum.Member`. What is left is not ours: `oNext.children.Object` on a `Collection` (a VBE
  probe blocked Excel on the same line too) and Example's own duplicate backup modules.
- **`MsgBox` as a value emits an `int`, as its declared `Long` says.** A private project in the
  corpus hid 113 VBA0003 of ours behind its own dead procedures; with those gone, one emitter bug
  remained. Strict whole-project compilation (D-I) held: the fix went into the project, not vba-ng.
- **`vbang init` and `vbang report` (WP7).** `init [--vscode]` writes the manifest, a `.gitignore`,
  a build task whose `$msCompile` matcher reads the diagnostics vbang already prints, and an
  `EXCEL.EXE` attach configuration; it never overwrites a project. `report` zips the versions, a
  build it runs itself (diagnostics go to stdout, so there is no build log to collect), the
  manifest, the sources, `out/`'s json and log, the generated C#, and `status.json` when the
  add-in answers — never the workbook, which makes the zip safe on a public issue. Still open:
  the release zip, `vbang install`, `--out`.
- **Cross-workbook `Application.Run` deferred past the alpha.** A `Book!Proc` prefix naming another
  loaded project's workbook does not resolve, because Excel no longer has a VBA project to answer
  it; the fix is a seam to the other project's Run table, a Runtime public API change. Of 37
  `Application.Run` uses across the corpus none name a workbook, nor do any of the four `OnTime`
  uses. The `[in, out]` write-back item closed on the argument instead: the one such
  parameter that matters in the wild, an event's `Cancel`, already works through the sinks.
- **A probe must exercise what it claims to test**, now a CLAUDE.md rule, since `Application.Run`
  compiles only the procedure it runs: the `End` and `Exit` findings were re-probed accordingly.

## 2026-09-11

- **Arrays are SAFEARRAYs and the Variant is the VARIANT (M7 C4).** `VbaArray` is a view of a
  SAFEARRAY from oleaut32, elements in their native layout, column-major, with the feature flags
  the Memory golden recorded; a Variant holds the descriptor under VT_ARRAY, so nothing in it is
  managed. A For Each locks the array: an Excel probe showed `ReDim`, `ReDim Preserve`, `Erase`,
  and an array assignment all raise 10 and leave it as it was, the loop reading live.
- **Pointers answer from where storage already lies (M7 E1, D21).** `StrPtr` is the BSTR, `ObjPtr`
  the interface pointer, and `VarPtr` the address of storage already stable — a local, a
  parameter, a With alias, a standard module's unmanaged static, an array element. An array element
  passed ByRef to a `Declare` now travels as its address, closing WP0's memory-safety item, and the
  per-thread arena was dropped: the stack already gives a local a stable address the debugger sees.
- **Typed fast paths (M7 E2, D-J).** A `For` counter declared Byte through Double is a native loop
  whose step raises 6 on overflow, under `On Error` as in structured code; `+`, `-`, and `*` over
  declared numbers call typed helpers. Unit tests hold every helper to the Variant operators over
  each pair of boundary values. The Variant loop went 23.3x VBA to 13.6x in Debug.
- **Typed object variables, instance blocks, native records (M7 E3, E4, C5–C11).** An object
  variable is an `ObjectSlot<T>`, the interface pointer itself owning one reference; a class
  instance keeps its variables and `Static` locals in a native block. A record is laid out as VBA
  lays it out: fixed-length strings inline at one-byte alignment, Booleans two bytes, fixed-size
  arrays and arrays of records inline, a dynamic array member VBA's eight-byte descriptor pointer.
  Memory cases recorded from Excel pinned each step before it was built.
- **Every `Declare` shape by address (M7 E5–E14).** Records, Variants, objects, Booleans, arrays,
  array and Variant and String results, a type library's interface parameters, and callbacks
  taking or returning Strings and Variants all cross as VBA passes them; a record with String
  members travels as a copy whose strings are ANSI, read back after the call. VBE7's `VarPtr`,
  `VarPtrArray`, and `rtcCallByName` are answered by the runtime, so code declaring them works
  without VBE7 loaded. The corpus then reported no VBA0002.
- **The Financial functions compute in x87 extended precision (D-F closed).** About eight hundred
  values probed from Excel showed VBA evaluates them in the extended format, rounding to Double
  only at `pow`, the due factor, the objective `Rate` and `IRR` iterate on, and the result; `NPV`
  of six flows is the exact value rounded once, which no ordering in Double reaches. Two earlier
  sessions had chased evaluation order through sixty forms: the wrong thing. Financial went from 17
  named misses to none.
- **Typed Double arithmetic stays in Double (D22), the one deliberate exception to bug-for-bug.**
  Measured before deciding: the software extended format costs about a hundred times a hardware
  operation and would have roughly doubled the typed-Double benchmark's gap, to remove one-ulp
  differences — 109 of 569 probed chains, at most 2.2E-16 relative, 3 visible in VBA's fifteen
  printed digits. The benchmark did find a typed `/` still going through the Variant operators;
  with a Double quotient it now calls `Operators.DivideDouble`, 2.08x VBA to 1.09x.
- **`Application.Run` of the project's own procedures runs in-process (D23).** A project a host
  command loads registers no names with Excel, so VBA-TDD's `RunMatcher` could not find
  `Specs.ToBeAnEmptyArray`. The runtime now looks a macro name up in the calling project first,
  through a Run table per module, with the name forms an Excel probe showed.
- **The corpus libraries' own suites pass in Excel (WP3 closed).** VBA-JSON's, VBA-Dictionary's,
  and VBA-UTC's spec suites pass under `vbang.Test`, and VBA-Better-Array's own TestRunner gives
  VBA's report — 407 tests, 406 passed, the one VBA fails failing the same way. Compiling is not
  compatibility: running the authors' tests found about fifteen gaps no scorecard showed, each
  closed with a golden. `Format` took a pattern with digit placeholders as a number format even
  when a date code came first, so ISO 8601 came back as the pattern; a late-bound `Set` with an
  argument went out as a put, not a putref; and a For Each over a COM collection now keeps the
  element it is on, which had crashed Excel across processes.
- **Host fidelity (WP4).** `Application.Volatile`, `Caller`, and `ThisCell` answer from the call
  Excel is making while a UDF is on the thread. Measured first, and only one of the three had
  worked: `Caller` was right through COM, `ThisCell` raised 1004, and `Volatile` was accepted and
  silently ignored, so a function that asked to be volatile never evaluated again. Every UDF now
  registers as macro-type, since the XLM functions may only be called from such a function, which
  is what VBA is. An unhandled error shows VBA's dialog text and reaches the project's log, and a
  workbook that still carries its VBA project is not bound (D17). `WithEvents App As Application`
  in a document module, `Application.OnTime`, the workbook-qualified `Application.Run`, the
  two-projects-one-name policy, and `Stop` were verified rather than written; `End` left `Open`'s
  files open, caught by a failing test first.
- **A project log, and the locals window checked (M7 F3, WP4).** `ProjectLog` writes
  `out/output.log`, one line per print with its time, rolling at a megabyte, and `vbang logs
  [--follow]` reads it. samples/Storage holds every kind of storage and stops; stepped through in
  VS Code, every local read as its VBA value once the native kinds carried debugger views and each
  generated class named a view of its variables.
- **The E2E suite is serial, and the Legacy fixture opens again.** Running its classes in parallel,
  one hidden Excel each, made a CLI test find another instance and skip itself on every full run.
  The committed import fixture took Excel down within seconds of opening, with or without the
  add-in: its own macros looped — `Workbook_Open` writes A1, `Worksheet_Change` answers with B1.
- **Benchmarks moved to Release (D-K), and a dispid cache was priced and dropped (D-L).** Most of
  the Debug baseline's 26x on the Variant loop had been the Debug runtime, not generated code.

## 2026-09-10

- **Strings become BSTRs (M7 B1–B4).** A `String` variable, parameter, field, record member, or
  result is a `VbaString` slot owning its BSTR; a literal is a BSTR the module allocates once and
  every use views; the Strings module, the number and date parsers, `Like`, and the operators read
  arguments as spans over the BSTR with no .NET string on the way. The `&` operator joins bytes,
  binary comparison ranks a trailing odd byte, and the `B` functions closed the nine misses (D-E).
- **Ownership rails for arrays and records (M7 B5, D18).** A store copies what it must own, a
  release frees what the slot held, and a fresh array with elements to release or a Function's
  array or record result is a temporary of its statement. Two constructs take a value over
  explicitly: a For Each over a temporary array detaches it for the loop, and a With over a
  function's record result owns a copy while a With over a variable aliases it — the frame
  releasing that alias had wiped a module-level record's strings.
- **The handle table was dropped, and D before C.** B5's other half, a blittable VARIANT with a
  handle table standing in for objects, arrays, and records, was scaffolding C and D would replace.
  The VARIANT lands when every payload has a native form, which puts objects first: a Variant
  array's elements are VARIANTs, and a VARIANT cannot hold an object until objects have pointers.
- **A COM object a call returns is a temporary of its statement.** The benchmark E2E hung inside a
  COM call Excel never returned from, and a second Range loop in the same Excel read wrong values
  or raised 438. The same probes hung against older builds, so the fault predated M7 and was
  timing-dependent: a returned wrapper was released by the finalizer, on the finalizer thread,
  against Excel's STA objects, and Excel's state went wrong once enough Ranges had gone that way.
  Every wrapper Interop makes from a returned pointer is now a temporary of its statement, host
  calls open a frame, the add-in pins what it keeps, and the finalizer only counts. The Range loop
  then ran in a quarter of the time, 2.29x VBA from 2.6x.
- **A project resets at `End` and at unload, without `Class_Terminate` (M7 D1).** Excel probes
  showed that a workbook closing destroys the objects its project holds without running
  `Class_Terminate`, and that `End` does the same for the objects its unwound locals hold, runs
  nothing after it, and returns module-level variables to their initial values.
- **Class instances, Collections, `Err`, and the enumerator get a COM identity (M7 D2, D3).** A
  native block made with the first reference: an `IDispatch` vtable of `[UnmanagedCallersOnly]`
  thunks, D18's count as COM's own number, and a strong handle to the managed object while the
  count is above zero. A runtime object crosses to COM as its own pointer and comes back as itself
  — a Scripting.Dictionary holding a class instance returns that instance, and its `RemoveAll`
  runs `Class_Terminate` once. A project reset cuts the blocks off (RPC_E_DISCONNECTED).
- **Objects, Decimals, and records in the Variant (M7 C1–C3).** A Variant holds an object as its
  IDispatch pointer and owns it through `AddRef`/`Release`; a Decimal lies inline in the VARIANT's
  explicit 24-byte layout; a user-defined type is a generated struct copied by value. A Variant
  never holds a record, which VBA allows only for public types of class modules.

## 2026-09-09

- **The project review that set the road to the alpha.** Numbers first: 2,004 of 2,028 goldens
  (98.8%), and 64 of 131 corpus files compiling against the fixture libraries. Then seven defects
  no list carried: reading the Office type library threw an unhandled exception out of the CLI,
  and the importer dodged it by treating Office as implicit and losing the library instead; the
  document-module attribute rule also matched a PublicNotCreatable class with a predeclared
  instance, which is how stdVBA exports every class, so seven compiled as worksheets; a `Property
  Let` with a parameter emitted its dispatch case without `ref`; `Mid$(s, 1, 1) =` did not parse;
  `[A1]` foreign names bound as variables; duplicate enum member names were reported at the
  declaration where VBA reports only at an unqualified use; and the scorecard's total left out the
  cases the binder rejects. All closed the same day (WP0).
- **VBA storage is real memory (D20), the largest decision since the runtime core.** Pointers are
  a paradigm code in the wild builds on, not a corner — `CopyMemory` between variables, the BSTR
  swap, reading a VARIANT's type bytes, `VarPtrArray` into the SAFEARRAY descriptor, a weak
  reference kept in a `LongPtr`. The corpus makes 206 uses of the three functions, 139 of them
  arguments of a `Declare`. Rejected, each incomplete or silently wrong: leaving them diagnosed
  (four of seven corpus projects would not build); pinning a variable for the call, which hands a
  dangling address to any callee that keeps it; a handle for `ObjPtr` alone; a stub that raises.
- **Two hazards a compile does not show,** found reviewing a `Declare`-heavy private project: an
  array element passed ByRef to a `Declare` was copied into an eight-byte temporary, so a DLL that
  reads the elements after it (the Fortran convention of passing an array by its first element)
  read garbage and overwrote the stack; and with the Office library missing, `msoTrue` in a module
  without `Option Explicit` bound as an implicit Empty, wrong rather than failing to build.
- **The corpus compiles against the registered libraries (WP1).** Named arguments to intrinsics by
  their MS-VBAL parameter names and to late-bound targets through `GetIDsOfNames` — the recording
  found that `InStr` and `StrComp` take no named argument at all in Excel. `WithEvents` on a
  referenced library's type advises through the runtime's `IComEventSource`. The scorecard loads
  every type library model the CLI has cached, one replacing the fixture of the same library, so
  `dotnet test` without `VBANG_CORPUS` still needs nothing registered. `MSForms.Control` members
  its interface lacks bind late, as VBA does: the reader records whether an interface carries
  `TYPEFLAG_FNONEXTENSIBLE` (Excel's do, MSForms' do not).
- **Six golden areas recorded, and what they forced (WP5).** ControlFlow (73 cases), Types (41),
  Lexer (14), Scope (27), Memory (46), and eight more FileSystem. `LSet` and `RSet` were
  implemented; `For Each` reads array elements live; a typed `For` counter let-coerces its limit
  and step; `Select Case` compares by declared type and raises 94 on a Null clause; `Len` and
  `LenB` of a type count a variable-length string member as its pointer. The lexer follows the VBE
  on suffixes (`a&b`, `x^2`, `2^3` rejected) and reserves `Local`, which MS-VBAL 3.3.5.2 omits.
- **Errors and Conversion to 100%.** A Double division by zero or overflow stores its IEEE result
  and raises afterwards, where a Variant target does not; `Erl` comes from a per-procedure table of
  line numbers and code offsets, not the lines that ran (forty cases pinned the rules down); and
  `CDec` keeps a zero's scale and sign.
- **The performance baseline, in Debug (WP7),** taken before the memory model changed the runtime
  so it would have a number to answer to. The Variant loop's 26x traced to the generated code, not
  the runtime's data: every operator and every `For` step went through the Variant helpers, which
  became D-J. Decisions D-A through D-J were all taken this day.

## 2026-09-06

- **Class modules with VBA's object lifetime (M5).** Reference counting in generated code (D18)
  rather than the collector, because `Class_Terminate` is observable and code in the wild depends
  on when it runs; the Classes golden pins the order, down to a failed `Class_Initialize` never
  seeing `Class_Terminate`. `Implements` emits a C# interface beside the class's own
  implementation, so a variable of the interface type holds the implementing object itself and
  `TypeOf`, `Is`, and a cast back all work without a wrapper.
- **The importer, `Declare`, and `AddressOf` (M6 opened).** `vbang import` reads vbaProject.bin
  per MS-OVBA and writes the VBE's own export formats, with no COM and no Excel; the corpus
  compile scorecard started running. The README and its quickstart sample followed the same day.

## 2026-09-05

- **The back end (M3).** Binder, C# emitter with the dispatch loop, and sixteen diagnostics. The
  golden replay stopped using the syntax-tree interpreter and began compiling every case through
  the emitter, with the same 1,809 of 1,845 passing, which made it a test of the compiler too.
- **`vbang test` and the `'@Test` annotation (D13).** A Public Sub with no parameters whose
  leading comments include one starting with `@Test`, asserting through the runtime's `Assert`
  module, reported as text or JUnit.
- **COM, late and early (M4).** Interop marshals VARIANTs and SAFEARRAYs by hand and calls
  `IDispatch::Invoke` through the vtable, with no `dynamic` (D7); the runtime's `IDispatchObject`
  contract keeps the runtime free of COM. Registered type libraries are read into a JSON model the
  compiler consumes without referencing Interop, cached per library version; the Scripting model
  is a committed fixture, so compiler tests need no COM.
- **The Excel host (M4).** The add-in binds a workbook to the folder named after it, builds it
  through the CLI when stale, sets `Me` by CodeName, advises Worksheet, Workbook, and ActiveX
  events through a native IDispatch sink, runs `Workbook_Open` and `Auto_Open`, registers public
  Subs as commands and Functions as UDFs, hot reloads, and collates under NLS. Two facts came from
  Excel itself: its objects expose no class information, so event interfaces come from the type
  library model, and it fires events with rgvarg in declared order.
- **Host services, non-interactive by default (D8).** Real dialogs for code run inside Excel,
  defaults during a host command unless `--ui` asks otherwise, so an agent triggering a legacy
  module cannot deadlock Excel behind a modal dialog. `IMEStatus` is unsupported for good.
- **File statements and FileSystem**, `Open` through the registry functions. The golden recorder
  now reopens its results file per record, so a case's own `Close` cannot cut a recording short.
- **D14 to D17 settled** the open questions of the day. Host command responses longer than Excel's
  32,767-character string limit travel through `out/response.json`.

## 2026-09-04

- **Scaffold and the spike (M0).** A breakpoint in `Hello.bas` hit inside `EXCEL.EXE` and `vbang
  run` printed `Debug.Print` output. That settled the approach: compile out of process (D1) to C#
  with a `#line` per statement (D2) and let Roslyn and the portable PDB do the debugging, which
  every .NET debugger then reads for free.
- **The MS-VBAL front end (M1).** Lexer, recursive-descent parser, conditional compilation,
  attributes, and error recovery, written from the specification with no GPL grammar consulted
  (D3, D10). The corpus — 131 files of permissively licensed VBA, gathered outside the repository
  — parsed 85% on the first run; triage gave three parser fixes and two quirk entries, and it
  ended at 131 of 131 parsing and round-tripping byte-exact.
- **The runtime core and the golden harness (M2 opened).** 15 areas, 1,818 cases and 2,099
  recorded values from Excel 16.0.20326 under en-US, 271 of them ending in a VBA error. The
  `Variant`, let-coercion, the operators, and the library modules replayed 1,809 of 1,845 through
  a small interpreter over the syntax tree, so the library was measured before the back end existed.
- **The machine, and .NET 10.** 64-bit Excel (Microsoft 365, build 16.0.20326) with the .NET 10
  SDK only, so D9's ".NET 8 or later" was pinned to .NET 10.

## Decisions

In force unless marked; the design itself is ARCHITECTURE.md's.

- **D1** — Compilation happens out of process, so Roslyn never loads into Excel, a compiler crash
  cannot take a workbook down, and CI produces the same bytes as the desktop.
- **D2** — Codegen targets C# with a `#line` per statement, compiled by Roslyn, so the PDB maps IL
  back to `.bas` and any .NET debugger attaches to `EXCEL.EXE` with working locals.
- **D3** — Own the front end and the runtime; no VB.NET transpilation, where `Integer` is 32-bit,
  `Currency` does not exist, arrays are zero-based, and error numbers do not match.
- **D4** — One project folder per workbook: `Sales.xlsx` binds `Sales.vbang/`, files in the VBE's
  export format verbatim, `Attribute` lines included.
- **D5** — The interface is the CLI, files, and stdout. Diagnostics use MSBuild's canonical
  format, which every editor turns into clickable problems; every command takes `--json`.
- **D6** — The CLI reaches Excel only by COM automation of the running instance, through hidden
  XLL commands called with `Application.Run`. No sockets, pipes, or daemons.
- **D7** — The runtime owns its `IDispatch` invoke path, hand-written, with no `dynamic`, because
  `Currency`, `Date`, `Decimal`, `Empty`, `Null`, `Error`, `Missing`, ByRef arguments, and
  non-zero lower bounds need exact treatment. *Revised by D20:* now a passthrough.
- **D8** — CLI-driven runs are non-interactive by default; `--ui` restores dialogs, and runs
  triggered inside Excel are always interactive.
- **D9** — .NET 10 via Excel-DNA 1.9, Excel 2016 and later, 64-bit first. Originally ".NET 8 or
  later"; pinned to the LTS and the only runtime on the dev machine.
- **D10** — MIT. No GPL-derived code enters the repository.
- **D11** — VS Code is the primary editor; `vbang init --vscode` writes the `$msCompile` build
  task and a `coreclr` attach configuration for `EXCEL.EXE`.
- **D12** — UserForms are deferred; worksheet controls are fully in scope and need nothing beyond
  the event and registration machinery. What happens to a `.frm` meanwhile is D-D.
- **D13** — A test is marked by a `'@Test` comment, not a naming convention, so procedures keep
  their names and the mark survives the VBE round trip.
- **D14** — The project folder suffix stays `.vbang`: it names the tool that reads the folder.
- **D15** — Build output stays in `out/` inside the project folder, gitignored. `--out` is open.
- **D16** — `Workbook_Open` and `Auto_Open` run for workbooks already open when the add-in loads,
  once per load, never on hot reload: a workbook whose handlers never ran is a state VBA never has.
- **D17** — `vbang import` never modifies the workbook; `--to-xlsx` writes the macro-free copy,
  which is the form vba-ng runs, and a workbook that still carries a VBA project is not bound.
- **D18** — Object lifetime is reference counted in generated code, not left to the collector,
  because `Class_Terminate` is observable and code in the wild depends on when it runs. Cycles
  leak, as in VBA. Rejected: finalizers (non-deterministic, wrong order) and `IDisposable` scopes
  (no shared ownership). Later extended to COM wrappers, then to arrays and records.
- **D19** — A document module is identified by the workbook, not by attributes: CodeNames from the
  OPC parts, the manifest's `documents` map overriding, the attribute pair last — which alone
  cannot tell a sheet module from a predeclared PublicNotCreatable class.
- **D20** — VBA storage is real memory: a Variant is a VARIANT, a String a BSTR, an array a
  SAFEARRAY, a record a block, an object an interface pointer, and the Ptr functions report those
  addresses. Explicit ownership is also the ground the multithreading extension stands on; the
  2026-09-09 entry has what was rejected.
- **D21** — A frame is the stack; a stable address comes from where storage already lies. The
  per-thread arena the design had described was dropped as adding only contiguity.
- **D22** — Typed Double arithmetic computes in Double, not the x87 extended format VBA keeps its
  intermediates in — the one deliberate exception to bug-for-bug, priced at one ulp against
  roughly doubling the typed-Double gap. The Financial functions still use the extended format.
- **D23** — `Application.Run` of the calling project's own procedures runs in-process, since a
  project a host command loads registers no names with Excel, and registering them would leave
  names Excel-DNA cannot unregister.
- **D-A** — *Superseded by D19.* How a document module is identified: by the workbook.
- **D-B** — *Superseded by D20.* The pointer family is a representation change, not a diagnostic.
- **D-C** — Office is not a default reference: nothing adds a reference a project did not ask for,
  and the importer writes Office only when the workbook carries it.
- **D-D** — UserForms are post-1.0.0; a `.frm` builds with a warning and its procedures raise.
- **D-E** — *Withdrawn, resolved by D-B:* a BSTR holds any byte count.
- **D-F** — Last-bit Financial results and `Tan(1)` are fixed properly, not left as named misses.
- **D-G** — A second locale is post-1.0.0: one regeneration of Conversion, Format, DateTime, and
  DateLiterals under de-DE, stored as a second golden set the replay runs under that culture.
- **D-H** — The alpha's import criterion takes the corpus spec workbooks plus stdVBA's two.
- **D-I** — A project containing a procedure that does not compile does not build. VBA compiles
  lazily, so a dead procedure with an error can ship unnoticed; vba-ng stays strict, and the
  recommended opt-in `"deadCode": "warn"` setting was rejected.
- **D-J** — Typed fast paths went into M7's compiler storage work, not a pass of their own.
- **D-K** — Benchmarks always run the Release build, generated code as `vbang build` emits it.
- **D-L** — No dispid cache: measured first, and late-bound calls already beat VBA's own.
- **D-M** — *Became D22.* Option (c), the cases stay named, after measurement killed option (a).

## Measurements

What the suites last measured is [measurements.md](measurements.md)'s, with the command and report
file behind every number. This file records why a number moved, not what it is.
