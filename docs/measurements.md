# Measurements

What the suites measure and what they last measured. Every number here is copied from a report file
a test writes, never typed from memory, so any of them can be re-taken by running the command beside
it. The design is [ARCHITECTURE.md](../ARCHITECTURE.md)'s and the history
[build-log.md](build-log.md)'s.

| What | Produced by | Report | Needs Excel |
|---|---|---|---|
| Golden replay | `dotnet test tests/VbaNg.Golden` | `golden-report.txt`, `golden-leaks.txt` | no |
| Corpus parse and compile | `$env:VBANG_CORPUS = "<roots>"; dotnet test` | `corpus-report.txt`, `corpus-compile-report.txt` | no |
| Benchmarks | `$env:VBANG_E2E = "1"; dotnet test tests/VbaNg.E2E -c Release` | `benchmark-report.txt` | yes |

Every report lands next to the test assembly that wrote it, under that project's
`bin/<configuration>/<framework>/`. A folder for another configuration or framework can hold an
older copy of the same file name, so read the one whose path matches the run.

## Goldens

<!-- generated from golden-report.txt by tools/Update-Measurements.ps1 -->
**2,472 cases, 2,470 pass (3 of them within one ulp of VBA), 2 fail,
99.9%**, over 25 areas.

Every case passes in: Arrays 88, ArraysBase1 12, Classes 48, Constants 129, ControlFlow 73,
Conversion 318, DateLiterals 12, DateTime 178, Declares 44, Errors 97, FileSystem 64, Financial 101,
Format 148, Information 191, Interaction 80, Lexer 14, Math 89, Memory 63, Objects 24,
Procedures 81, Scope 27, Strings 280, StringsText 26, Types 45.

Short of it: Operators, 240 cases, 99.2%: 1 within one ulp, 2 named misses.
<!-- end generated -->

Cases that do not pass are named by area under `tests/VbaNg.Golden/Expected`. Every one-ulp
tolerance and every named miss above comes from one decision rather than a defect: a typed `Double`
expression computes in `Double`, where VBA keeps its intermediates in the x87 extended format. That
is why Financial holds its 100% with two tolerated cases and Operators does not reach it.

The replay also asserts that every case returns the count of live native allocations to where it
found it, writing `golden-leaks.txt`. Five cases are exempt by name in `Expected/Leaks.txt`: two
Classes cases, one of them the reference cycle that never terminates in VBA either, and one each in
Memory, Scope and Types, which fill module-level storage.

## The corpus, last run 2026-09-13

Several roots separated by `;` join one run.

Parse: **139 of 139 files, 100%, no round-trip failure.**

Compile, against the type library models the CLI has cached — here Excel, Scripting, stdole, VBIDE,
MSForms, ADODB, Office and WinHttp:

| Project | Files | Without errors | Errors |
|---|---:|---:|---:|
| stdVBA | 48 | 43 | 10 |
| VBA-Better-Array | 20 | 20 | 0 |
| VBA-Dictionary | 2 | 1 | 5 |
| VBA-JSON | 2 | 1 | 3 |
| VBA-TDD | 8 | 8 | 0 |
| VBA-UTC | 2 | 1 | 3 |
| VBA-Web | 44 | 33 | 69 |
| xlwings (addin) | 13 | 13 | 0 |
| **Total** | **139** | **120 (86%)** | **90** |

Two causes hold 79 of the 90 errors and are the same thing: a loose specs module wants the
`SpecSuite` its workbook carries (47 VBA0012, plus 32 `New` of that undefined class). The rest are
stdVBA's test modules reaching document-module members, a duplicate module name, and VBA-Web's
specs wanting VBA-TDD and their workbooks' modules. Every folder is compiled whole, specs and
examples included; the spec workbooks themselves import and build.

## Benchmarks, last run 2026-09-11, Release

samples/Benchmarks, run by tests/VbaNg.E2E in the Release build only, since Debug measures the
Debug runtime rather than the design. Budget: no row slower than the baseline, and more than 3x
VBA on the Range loop becomes a work item. VBA's own times vary by up to 3x between runs here, so
only ratios inside one run compare — the Range loop's 2.89x below sits beside 1.94x the same day.

| Benchmark | vbang | VBA | ratio |
|---|---:|---:|---:|
| Variant loop | 194 ms | 48 ms | 4.04x |
| Typed Double arithmetic | 46 ms | 37 ms | 1.25x |
| String building | 115 ms | 50 ms | 2.31x |
| Range loop | 2,180 ms | 755 ms | 2.89x |
| UDF over 10,000 cells | 50 ms | 68 ms | 0.74x |
| Late-bound Collection (fifty thousand Adds and Counts) | 19 ms | 72 ms | 0.26x |
| Late-bound Excel (twenty thousand property reads) | 48 ms | 80 ms | 0.60x |

The UDF beats VBA because Excel-DNA calls the function directly rather than through automation. The
two late-bound rows are what priced a dispid cache out: it measured slower than the call it was
meant to save.
