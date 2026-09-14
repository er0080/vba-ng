# vba-ng Architecture

The design in force: structure, components, data flow. Planned work is [ROADMAP.md](ROADMAP.md);
the reasoning behind a change is [docs/build-log.md](docs/build-log.md).

## 1. Vision

**Goals.** 100% VBA language compatibility, measured against real VBA rather than asserted; full
interop with COM libraries — the Excel object model, other Office applications, any registered type
library; no IDE, since VS Code and terminal-based coding agents are the primary environments;
git-native, one project folder per workbook in the exact export format the VBE produces; real
compilation to IL, debuggable through standard .NET debuggers. The lexer, parser, binder, and
runtime are vba-ng's own, with no VB.NET transpilation.

**Non-goals.** Bridging to or hosting the real VBA engine; an MCP server or any other
agent-specific protocol, where the CLI, files, and stdout are the interface; the VBE extensibility
model (`Application.VBE`); Excel for Mac, Excel Online, and non-Excel Office hosts.

**Targets.** .NET 10 through Excel-DNA 1.9; Excel 2016 and later plus Microsoft 365; 64-bit, with
`LongPtr` 64 bits. Only one .NET runtime version can live in an Excel process — the first .NET
add-in to load picks it for every other — so the add-in targets the LTS runtime alone.

## 2. System overview

```
  editor / agent          vbang CLI                   EXCEL.EXE
  +-------------+  files  +---------------+  COM      +--------------------------+
  | VS Code     | ------> | build         | --------> | vba-ng add-in (.xll)     |
  | Claude Code |         | run / test    | App.Run   |  loader (collectible ALC)|
  | any editor  | <------ | import / init | <-------- |  runtime + interop       |
  +-------------+  diags  +------+--------+  results  |  UDF/command registration|
                                 |                    |  event sinks, watcher    |
                                 +--> Sales.vbang/out/ <-- spawns `vbang build` --+
```

| Project | Responsibility |
|---|---|
| `VbaNg.Runtime` | `Variant`, strings, arrays, records, intrinsic library, `Err`, `Collection`, project hosting |
| `VbaNg.Compiler` | Lexer, parser, full-fidelity syntax tree, binder, lowering, C# emitter, Roslyn driver |
| `VbaNg.Interop` | Type library reader, `IDispatch` invoke and marshaling, connection-point sinks |
| `VbaNg.Import` | `vbaProject.bin` reader (MS-OVBA) and the source writer |
| `VbaNg.AddIn` | The `.xll`: project discovery, watcher, build spawn, loading, registration, host commands |
| `VbaNg.Cli` | The `vbang` executable |

**Dependency direction**, enforced by project references: Runtime and Import reference nothing of
vba-ng's; Compiler and Interop reference only Runtime; the AddIn references Runtime and Interop plus
Excel-DNA; the CLI references Compiler, Import, Interop, and Runtime. Only Interop and its two hosts
touch COM, and generated code sees only Runtime's public surface.

**Compilation is out of process.** The compiler is a plain .NET library driven by `vbang`; it
produces a DLL and a portable PDB into the project's output folder, and the add-in only loads that
output, so Roslyn never enters Excel's process. Type libraries come from the registry, so binding
needs no live Excel either.

## 3. Project layout on disk

```
Sales.xlsx
Sales.vbang/
  vbang.json     manifest, optional      Sheet1.cls   document module for CodeName Sheet1
  Module1.bas    standard module         out/         gitignored build output
  Customer.cls   class module            .vscode/     optional, from `vbang init --vscode`
```

`out/` holds `Sales.dll`, `Sales.pdb`, `gen/*.cs`, `build.json`, `output.log`, and `response.json`.
The add-in derives the project folder from the workbook's full path by replacing the extension with
`.vbang`; a workbook with no such folder is ignored.

**Document modules.** The compiler reads the workbook's CodeNames out of the file next to the
project folder — the OPC parts `xl/workbook.xml` and each sheet or chart-sheet part, a zip read with
no Excel — and a `.cls` whose `VB_Name` is one of them is a document module of that object's kind.
The manifest's `documents` map overrides, and carries a project with no workbook beside it. Only
without either does the attribute pairing `VB_PredeclaredId = True` with `VB_Exposed = True` mark
one; `VB_PredeclaredId = True` alone is an ordinary class with a predeclared instance.

**Manifest.** `vbang.json` carries the project name, its `references`, and optionally a `documents`
map of module name to `Workbook`, `Worksheet`, or `Chart`:

```json
"references": [
  { "name": "Excel" },
  { "name": "Scripting", "guid": "{420B2830-E718-11CF-893D-00A0C9054228}", "version": "1.0" }
]
```

A reference resolves by guid and version, or by name alone against the libraries vba-ng knows
(Excel, Office, Scripting, stdole, VBIDE, MSForms) — and by that name when its guid is registered
nowhere, as a UserForm workbook's `MSForms` is. `VBA` and `stdole` are implicit; `Office`
is not, so its constants and types resolve only when the manifest names it, though `vbang init`
writes it alongside `Excel`. The name `vbang` is vba-ng's own library, bringing the `Assert` module
of test procedures (section 8) — it is the opt-in every vba-ng language addition needs. An invalid
manifest is diagnostic VBA0022; a missing one means no references.

**File formats.** Exactly the VBE export formats, with attribute lines parsed and honored:
`VB_Name`, `VB_PredeclaredId`, `VB_Exposed`, `VB_Creatable`, `VB_UserMemId`, and the rest. Files are
read with BOM detection and default to UTF-8; CRLF and LF are both accepted. A `.frm` compiles to
signatures only, with a warning per form: its controls are `Nothing` and the binder replaces every
procedure body with error 438 naming the gap, so the rest of such a workbook still builds and runs.

## 4. Compiler pipeline

1. **Lexing.** Line-oriented per MS-VBAL: `_` continuation, `:` statement separators, comments, date
   literals, typed literals and type suffixes, case-insensitive keywords. Identifiers keep the
   user's casing; comparison is case-insensitive.
2. **Conditional compilation.** A pass between lexer and parser evaluates `#If` and `#Const` against
   the predefined constants (`Win64`, `Win32`, `VBA6`, `VBA7` true; `Win16`, `Mac` false) and turns
   directive and excluded lines into trivia on the next kept token, so the tree still reproduces the
   file and excluded text need not be valid VBA.
3. **Parsing.** Hand-written recursive descent producing a full-fidelity tree with trivia that
   round-trips to the original text byte for byte. Every character belongs to exactly one token or
   trivia; a line terminator and a statement-separating colon are trailing trivia of a statement's
   last token, so a statement owns its line and block parsing stays line-agnostic. Error recovery is
   statement-level: absent expected tokens become zero-width missing tokens, unintelligible ones are
   kept in bad-statement nodes, and block terminators are matched against the stack of open blocks.
4. **Binding.** Symbol tables for modules, procedures, locals, UDTs, enums, and constants; the
   `Option` statements; implicit declarations when `Option Explicit` is off; COM types from the type
   library reader; resolution of default members, property `Get`/`Let`/`Set`, `Optional` and
   `ParamArray`, named arguments, and the `Let` versus `Set` distinction.
5. **Lowering.** `With` blocks, `For Each` over arrays and COM enumerators, `Select Case`, `Like`,
   the `Mid` statement, `Static` locals, `GoSub`, and the error-handling state machine.
6. **Emit.** C# source with `#line` directives naming the `.bas`/`.cls` file and line, one class per
   module, written to `out/gen/` and compiled in memory by Roslyn to `out/<Project>.dll` plus a
   portable PDB. `out/build.json` records a hash of every input, so the add-in can tell when a build
   is stale. Locals are hoisted to the top of the method with their default values and `Dim` lines
   carry no sequence point, as in VBA; a `ByRef` argument is a C# `ref` when the variable has the
   parameter's type, and a Variant variable or array element passes through a temporary copied back
   after the call. C# forbids a member named like its enclosing type, which VBA allows (`Sub Main`
   in module `Main`), so such a member is emitted with a trailing underscore and a `VbaName`
   attribute, through which the host resolves `Application.Run` and test names; test procedures
   carry a `VbaTest` attribute.

The same inputs give byte-identical output. Diagnostics have stable ids, listed with a triggering
snippet in [docs/diagnostics.md](docs/diagnostics.md).

### Error handling

A procedure that uses `On Error`, `GoTo`, `GoSub`, or `Resume` compiles into a statement dispatch
loop, its structured statements flattened to jumps first: a `switch` over a program counter inside
`while (true)`, each case ending in a `goto case`, wrapped in a `try` whose `catch` filter hands the
exception and the counter to the procedure's `ProcedureFrame` and takes back the counter to resume
at. `Resume`, `Resume Next`, `Resume label`, `Erl`, and `GoSub`/`Return` then need no special cases,
and `GoTo` into a block works. `Erl` is a compile-time table of the line each statement index
reports, since Excel resolves it from the procedure's numbered lines and code offsets rather than
from the numbered lines that ran. A procedure without those constructs compiles to ordinary
structured C# at full speed.

Runtime errors, COM errors from `IDispatch::Invoke` included, are `VbaException` instances carrying
VBA's number and message text, the project name as source, and VBA's help file and context. An
unhandled error at the top of a macro shows the run-time error dialog in interactive mode; in CLI mode
it is printed and logged.

## 5. Runtime library

### Value model

VBA storage is real memory. Every variable lives at an address stable for its lifetime, in the
layout VBA uses: a `Variant` is a COM VARIANT (24 bytes on 64-bit Excel, every VBA type in its OLE
Automation form), a `String` a BSTR, an array a SAFEARRAY over contiguous column-major elements with
the feature flags VBA sets, a user-defined type a block laid out by `RecordLayout`, an object
variable an interface pointer to a real COM object. `VarPtr`, `StrPtr`, and `ObjPtr` report those
addresses, and `CopyMemory` between variables means what it means in VBA. Nothing in a Variant is
managed, and conversion to and from ANSI happens only where VBA converts: a `Declare`'s
`ByVal String` arguments, the file statements, `StrConv`.

The address comes from where storage already lies:

| Storage | Lives in |
|---|---|
| Local, parameter, `With` alias, statement temporary | the stack frame of the procedure's own method |
| Module-level and `Static` variable of a standard module | statics outside the GC heap, from load to unload |
| Class instance's variables | the instance's own native block |
| Record member | inline in the record's block |
| Array element | inline in the SAFEARRAY's data |

The statics depend on the .NET runtime keeping statics of such types out of the GC heap, which the
Memory golden checks on every run; a `WithEvents` variable and a fixed-length string stay managed
fields of their instance. A BSTR admits any byte count, odd included, so the byte functions and a
String assigned from an odd-length Byte array are exact.

### Object lifetime

Reference counted, in generated code, because VBA runs `Class_Terminate` the moment the last
reference to an object goes away. The count is the object's COM reference count, so generated code's
`ObjectRefs` helpers and a COM caller's `AddRef`/`Release` reach the same number. The rails:

- A store into any slot adds a reference to the new value and drops one from the old, taking what it
  must own — a reference, or a copy of a string, array, or record with elements of its own — and
  releasing what the slot held. Reads are views: an element read, a `Collection` item, an
  enumerator's current element hand out the stored value, and only the next store copies.
- `New`, a Function or Property Get result, a fresh array or record with elements to release, and a
  COM object a call returns are temporaries of the statement that receives them; locals are released
  at procedure exit in declaration order, through a `finally`; arrays, records, `Collection`, and
  class fields release what they hold when erased, reassigned, or terminated. An owner's
  `Class_Terminate` runs before its fields are released; a failed `Class_Initialize` never sees
  `Class_Terminate`. Cycles leak, as in VBA.
- Two constructs take a value over past its statement: `For Each` over a temporary array detaches it
  and releases it when the loop ends, and `With` over a function's record result owns a copy, while
  `With` over a variable or element owns nothing.
- A COM wrapper is released on the thread that made it. A host call with no statement — an event
  dispatch, a UDF, a macro — opens a frame for the same purpose, and a wrapper the host keeps beyond
  it (the Application, a bound workbook, a sheet) is pinned with a reference of its own. The
  finalizer never releases a pointer, since Excel's objects live in the apartment of the thread that
  made them; one that reaches it still holding a pointer is counted as a missed rail.

Every BSTR, SAFEARRAY, and block has an owner, and the golden replay holds the live-allocation count
of each case to where it started. A project resets at `End`, and when the add-in unloads it at a
rebuild or at add-in shutdown: the reset destroys every object the project holds without running
`Class_Terminate`, and module-level variables and `Static` locals return to their initial values,
through each module's generated `__Reset`.

### Library and objects

The intrinsic library is organized as VBA's own modules — `Strings`, `Conversion`, `DateTime`,
`Math`, `Information`, `Interaction`, `FileSystem`, `Financial`, `Constants` — with `Format` its own
subproject carrying the full named and user-defined format grammar. Coercion follows MS-VBAL
let-coercion; the operators follow VBA's promotion and overflow rules, taking a flag for which
operands are declared `Variant` since MS-VBAL 5.6.9.3 lets an overflow widen only then, and calling
typed helpers instead of the Variant path where every operand's type is declared. Such a typed
Double expression computes in Double throughout, where VBA keeps its intermediates in the x87
extended format — a deliberate exception to bug-for-bug compatibility
([docs/vba-quirks.md](docs/vba-quirks.md)); `Financial` is unaffected, computing in the extended
format through `Extended`. Text comparison and case mapping go through Windows NLS, which Excel's
VBA uses, so hosts and tests run with `System.Globalization.UseNls`.

A class instance is a real COM object in a block the runtime owns: the vtables of `IUnknown` and
`IDispatch` (an enumerator also carries `IEnumVARIANT`), the reference count, and the instance's
module variables and `Static` locals. The vtable entries are `[UnmanagedCallersOnly]` thunks in the
runtime, never in the project's collectible assembly; they reach the project's methods through a
handle table that unloading neuters, so an object that outlives a hot reload raises an error instead
of calling unloaded code. Every class carries a dispid table built from its public members, honoring
`VB_UserMemId`, so late binding, `CallByName`, default members, `!`, and `For Each` reach a project
class through the same contract as a COM object, and `Implements` is resolved in generated code as
explicit implementations of the generated interface. `RaiseEvent` and `WithEvents` pass arguments as
a Variant array, which lets a handler's `ByRef` parameter write back; a source holds its sinks
weakly and a sink unsubscribes when it terminates. `Collection`, `Err`, the enumerator, and the event
sink have the same shape, so any can be handed to a COM method.

The file statements compile to calls on `FileSystem`, which numbers the current thread's open files
as VBA does and lays Binary and Random values out byte for byte; its channel table is the runtime's
rather than one project's, so `End` closes every loaded project's files. `IHostServices` covers
`Debug.Print` output, `MsgBox`, `InputBox`, `Beep`, and `SendKeys`: the add-in supplies the
interactive implementation, the CLI and test harnesses its defaults.

## 6. COM binding

- **Type library reader.** Reads `ITypeLib`/`ITypeInfo` for each manifest reference from the
  registry, with no Excel running, into a compact symbol model of coclasses, interfaces, dispids,
  members, parameter types and defaults, enums, default members, enumerators, and event sources. The
  model (`VbaNg.Runtime.TypeLibraries`) is plain records with a JSON form, kept in the Runtime so
  the Compiler can consume it without referencing Interop, and cached per library version under
  `%LOCALAPPDATA%\vbang\typelibs`.
- **Binding modes.** A variable declared with a COM type binds by dispid at compile time, so no
  `GetIDsOfNames` at run time: arguments are let-coerced to the declared parameter types, named
  arguments placed by position, omitted optionals passed as `Missing`, and a parameterless property
  returning an object takes trailing arguments on that object's default member (`Cells(1, 2)`).
  `Object` and `Variant` use late binding with a per-object name cache. A member the library does
  not list is a compile error on an interface marked `TYPEFLAG_FNONEXTENSIBLE` (Excel's
  `_Worksheet`) and a late-bound call on any other (`Range`, `Application`), as in VBA.
- **Invoke path, and the runtime stays free of COM.** The runtime defines the contract —
  `IDispatchObject`: dispid lookup, invoke, property put, enumeration, identity — which the
  late-bound helpers, `Coerce`, and `For Each` use exclusively; Interop implements it against the
  vtable of `IDispatch` with a hand-written `Invoke` wrapper and no C# `dynamic`; values are already
  VARIANTs, BSTRs, and SAFEARRAYs, so the layer is mostly a passthrough and `Currency`, `Date`,
  `Decimal`, `Empty`, `Null`, `Missing`, `ByRef` arguments, and non-zero lower bounds get exact
  treatment. `DISP_E` codes map to VBA's numbers;
  `DISP_E_EXCEPTION` becomes the object's own error.
- **Application objects, default members, creation, iteration.** The members of a library's
  application object (a coclass marked appobject) are in scope unqualified, over one instance per
  library supplied by the host's provider (`Com.Provider`). An object used where a value is needed
  stands for its default member's value, which `Coerce` resolves, so `x = Range("A1")` reads the
  cell. `New` on a coclass is `CoCreateInstance` by CLSID; `CreateObject` and `GetObject` resolve by
  ProgID and the Running Object Table; `For Each` uses `DISPID_NEWENUM` and `IEnumVARIANT`, which
  the runtime's own enumerator answers too; `TypeOf x Is Y` queries by IID and `TypeName` asks
  `IProvideClassInfo` first.
- **Events.** A sink is an `IDispatch` built in unmanaged memory — a vtable of
  `UnmanagedCallersOnly` functions and a `GCHandle` behind the object — whose `Invoke` marshals the
  arguments, calls the handler, and writes `ByRef` arguments back; it is advised through
  `IConnectionPointContainer`, with every advisory tracked so unload can unadvise. Used for
  `WithEvents`, document modules, and worksheet ActiveX controls. A `WithEvents` variable of a
  library type carries its coclass's default source interface and dispids from the type library at
  compile time and hands them to the runtime on `Set`, which advises through `IComEventSource`, an
  interface Interop implements. Excel fires events
  in declared order, not the reversed order `IDispatch` prescribes, so sources taken from a type
  library model are marked declared-order.
- **`Declare`.** Generated P/Invoke stubs following VBA marshaling rules: ANSI strings by default,
  `LongPtr` as `nint`, `Any` handled by an overload set. `AddressOf` produces delegates kept alive
  for the lifetime of the loaded project.

## 7. Excel integration (the add-in)

**Lifecycle.** The add-in advises `AppEvents` on the running Application and handles `WorkbookOpen`
and `WorkbookBeforeClose`; workbooks already open when it loads go through the same path.

1. Locate the project folder for the workbook. One whose folder does not exist is ignored, and so is
   one that still carries a VBA project (`Workbook.HasVBProject`): the add-in says why and to save
   a macro-free copy, so no macro runs twice.
2. Compare `out/build.json` with the sources and the add-in's version, and spawn `vbang build` out
   of process when either differs, finding the CLI through `VBANG_CLI`, beside the add-in, or in
   this repository's build output.
3. Load the DLL into a collectible `AssemblyLoadContext`, set the `Me` field of every document
   module to the object whose CodeName matches, advise the event procedures on that object or on the
   named ActiveX control, and register the project's names.
4. Run `Workbook_Open` and every public `Auto_Open`, since the workbook's own Open event fired before
   the sink existed — once per workbook per add-in load, never on hot reload.
5. On `WorkbookBeforeClose`, run the close handlers, unadvise, drop the binding, and stop watching
   the folder. The registered names stay: Excel-DNA offers no way to unregister them, so they
   outlive the workbook and are registered again after a rebuild. The project's assembly context
   unloads at add-in shutdown.

**Registration.** A public parameterless Sub of a standard module becomes an XLL command and a
public Function becomes an XLL UDF, visible in the function wizard under a category named after the
project, so `OnAction`, `Application.Run`, `OnTime`, and `OnKey` resolve by name. XLL registration
is global per Excel session, where VBA scopes UDFs per workbook. An ActiveX control's events are
advised onto the named procedures of the sheet's document module, bound early against MSForms.

**`Application.Run` of the calling project's own procedures runs in-process.** A call to Excel's
`Run`, on `Application` or unqualified, goes through the runtime, which looks the name up in the
calling project first: `Proc` among its standard modules, `Module.Proc` in a standard or document
module, either behind `Book!` when that is the project's own workbook. Every module carries a Run
table of its Subs and Functions, private ones included; a name found there runs in the same loaded
project, so the procedure sees the caller's module state as in VBA and an unhandled error escapes
every handler of the caller. Any other name goes to Excel, which raises 1004 unless it is another
workbook's macro or a registered command. This is also what makes a project loaded by a host command
reachable, since such a project registers no names.

**Host commands.** `vbang.Run`, `vbang.Test`, and `vbang.Status` are XLL functions registered hidden
from the wizard, which the CLI calls through `Application.Run`. Each returns one JSON string; a
response over 32,000 characters — under Excel's cap of 32,767 on an XLL function's
string result — goes to `out/response.json`, and the returned JSON only names that file. A run
triggered from inside Excel (a button, `OnTime`) is always interactive; a CLI-driven run is not unless
`--ui` is given, so an agent cannot deadlock Excel behind a modal dialog: `MsgBox` logs its text and
returns `vbOK`, `InputBox` returns its default, and vba-ng's own notices open no dialog in a hidden
Excel.

**Hot reload.** A debounced `FileSystemWatcher` per bound project folder, over the sources and
manifest at its top level, builds through the same out-of-process `vbang build` on a worker thread,
then loads and rebinds on Excel's thread between macros, without rerunning `Workbook_Open`. The old
assembly context unloads once every sink is unadvised and every COM reference released.

**UDFs.** Every UDF runs on Excel's main thread and may touch the object model, as in VBA.
Excel-DNA marshals its arguments, which the add-in turns into the Variants VBA would receive; a
reference becomes the `Range` itself, found through the reference's address, which is why a UDF
registers as macro-type — only a macro-type function may call the XLM reference functions. Errors
in events, commands, and UDFs go to the add-in's output, whose recent lines `vbang status` shows.

## 8. CLI

The interface is the CLI, files, and stdout: `init`, `import`, `build`, `run`, `test`, `status`,
`logs`, and `report`, whose options, output, exit codes, and project-folder discovery are
[docs/user-manual.md](docs/user-manual.md)'s. Diagnostics use the MSBuild canonical format,
`path(line,col): error VBA0001: message`, which every editor turns into a clickable problem;
`--json` carries the structured form where a command has one. Excel is found through the Running
Object Table or by walking its workbook window, and a call it rejects as busy is retried;
`Application.Run` blocks until the macro returns, so a long macro blocks `vbang run`.

**Test runner.** A test is a Public Sub without parameters whose leading comment lines include one
starting with `@Test`; a comment annotation rather than a naming convention keeps existing procedure
names, survives the VBE round trip, and lets a module mix tests with the code under test. Tests
assert through the `Assert` module of the `vbang` library, which the manifest must reference:
without that reference the name `Assert` means what it means in VBA, and with it a project's own
`Assert` still wins. A failed assertion ends the test at once and is not a VBA error any `On Error`
can absorb. `vbang.Test` writes the JUnit file itself, inside Excel, so the path is Excel's.

## 9. Editor and agents

`vbang init --vscode` writes `tasks.json` (a build task with the `$msCompile` problem matcher) and
`launch.json` (a `coreclr` attach configuration for `EXCEL.EXE`). With the `#line`-mapped PDB that
is the whole debugging setup: breakpoints in `.bas` files, stepping, locals, watches. `.bas` files
register no language here, so VS Code needs `debug.allowBreakpointsEverywhere` to break in one.

**Start and run under one keypress.** A launch configuration that starts Excel with the add-in and
a workbook and sets `VBANG_PROJECT` and `VBANG_RUN` gets a procedure run at startup: the add-in
polls until Excel is ready, splash screen gone and a workbook open, runs that procedure,
and mirrors `Debug.Print` to the debugger's console. Throwaway instances hold the add-in files
locked, so such a configuration closes them first.

Runtime types carry `DebuggerDisplay` and `DebuggerTypeProxy` attributes so a `Variant` reads as its
VBA value in the locals window and helper locals are hidden; where the debugger cannot read native
memory as a VBA value the type carries its own view, so an object variable shows `Nothing` or its
class and an array its elements under their VBA subscripts.

## 10. Import

`vbang import Book.xlsm` parses `vbaProject.bin` per MS-OVBA with no Excel and no "trust access to
the VBA project" setting, through its own compound file reader (MS-CFB) and MS-OVBA decompressor.
The `dir` stream gives the project's name, code page, references, and modules; the `PROJECT` stream
tells a class module from a document module; each module's stream carries its source after the
offset the `dir` stream records, and a class module's file gets the `VERSION 1.0 CLASS` header the
VBE writes but the stream does not. Sources are decoded from the system ANSI code page and written
as UTF-8. The output is export-format files, a manifest with the references, and a report of
what did not come over: a UserForm's code is written as a `.frm` while its `.frx` resources stay
behind.

The workbook is never modified, and an import never overwrites a project folder or a copy, so
rolling an import back is deleting a folder. `--to-xlsx` also saves a macro-free `Book.xlsx` copy
through Excel automation, in a hidden instance of its own with macros disabled, preserving CodeNames.
That copy is the form vba-ng runs.

## 11. Testing strategy

**Unit tests**, one project per source project: the lexer, the parser, the binder, lowering, and the
emitter, whose tests assert over the generated C# text. `tests/VbaNg.Compiler.Tests` also runs the
corpus over open-source VBA when `VBANG_CORPUS` points at it, reporting the parse and compile rates.

**Golden differential tests**, which are how any semantic claim is backed. Cases are `.cases` files
under `tests/VbaNg.Golden/Cases`, one per library area, a few lines of VBA each, with `?` lines — the
Immediate window shorthand — marking the expressions to record; a case may declare typed variables,
so typed and Variant operands are both covered, and an area needing other modules keeps them as
`<Area>.<Name>.cls` and `.bas` files in export format beside its case file.

`tests/VbaNg.Golden.Regen` records: it generates a recorder module per area, injects it into a hidden
throwaway Excel through the VBE object model, and runs every case in its own procedure under an error
handler, recording VBA's `TypeName`, a readable `Str` form, the exact bits of the floating- and
fixed-point types, and `Err` where a case ends in an error. The golden keeps the case source beside
its results, so a case file edited without regeneration fails a test that needs no Excel.

`tests/VbaNg.Golden` replays with no Excel present, compiling every case through the real back end
into a collectible context and running each in turn with `Err`, `Erl`, and the `Rnd` generator reset
between cases; every case must also leave the live native-allocation count where it found it, and
one the binder rejects is reported as unsupported, never as a pass. `tests/VbaNg.Golden/Expected`
lists, per area, the cases that do not pass, the ones tolerated at one ulp, and the ones that leak;
an unlisted failure fails the test, and so does a listed case that starts passing. The replay writes
`golden-report.txt`, the compatibility scorecard by library area; what it last said is
docs/measurements.md.

**End to end**, in `tests/VbaNg.E2E`: hidden throwaway Excel instances load the add-in from the build
output and run the sample projects under `vbang test`, plus, with a corpus configured, the corpus
spec workbooks; benchmarks run in the Release build and report against VBA.
