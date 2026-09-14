# Measurements

Every number here is copied from the report a test writes, and the command beside it re-takes it.
The design is [ARCHITECTURE.md](../ARCHITECTURE.md)'s and the history [build-log.md](build-log.md)'s.

| What | Produced by | Report | Needs Excel |
|---|---|---|---|
| Golden replay | `dotnet test tests/VbaNg.Golden` | `golden-report.txt`, `golden-leaks.txt` | no |
| Corpus parse and compile | `$env:VBANG_CORPUS = "<roots>"; dotnet test` | `corpus-report.txt`, `corpus-compile-report.txt` | no |
| Benchmarks | `$env:VBANG_E2E = "1"; dotnet test tests/VbaNg.E2E -c Release` | `benchmark-report.txt` | yes |

Every report lands next to the test assembly that wrote it, under that project's
`bin/<configuration>/<framework>/`; another configuration's or framework's folder can hold a stale
copy.

## Goldens

<!-- generated from golden-report.txt by tools/Update-Measurements.ps1 -->
**2,488 cases, 2,486 pass (3 of them within one ulp of VBA), 2 fail,
99.9%**, over 28 areas.

Every case passes in: Arrays 88, ArraysBase1 12, Classes 48, Constants 129, ControlFlow 73,
Conversion 318, DateLiterals 12, DateTime 178, Declares 44, Errors 97, FileSystem 64, Financial 101,
Format 148, Information 191, Interaction 80, LateBinding 10, Lexer 14, Lifetime 3, LineInput 3,
Math 89, Memory 63, Objects 24, Procedures 81, Scope 27, Strings 280, StringsText 26, Types 45.

Short of it: Operators, 240 cases, 99.2%: 1 within one ulp, 2 named misses.
<!-- end generated -->

Cases that do not pass are named by area under `tests/VbaNg.Golden/Expected`.

The replay also asserts that every case returns the count of live native allocations to where it
found it, writing `golden-leaks.txt`. Six cases are exempt by name in `Expected/Leaks.txt`: two
Classes cases, one of them the reference cycle that never terminates in VBA either, and one each in
Lifetime, Memory, Scope and Types, which fill module-level storage.

## The corpus, last run 2026-09-14

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

79 of the 90 errors have one cause: a loose specs module wants the `SpecSuite` its workbook carries
(47 VBA0012, 32 `New` of that undefined class). The rest are stdVBA's test modules reaching
document-module members, a duplicate module name, and VBA-Web's specs wanting VBA-TDD and their
workbooks' modules. Every folder is compiled whole, specs and examples included; the spec workbooks
themselves import and build.

## Benchmarks, last run 2026-09-14, Release

samples/Benchmarks, best of two runs each. Budget: no row slower than the baseline, and more than 3x
VBA on the Range loop becomes a work item. VBA's own times vary up to 3x between runs, so only
ratios within one run compare.

| Benchmark | vbang | VBA | ratio |
|---|---:|---:|---:|
| Variant loop | 161 ms | 46 ms | 3.51x |
| Typed Double arithmetic | 44 ms | 43 ms | 1.03x |
| String building | 103 ms | 59 ms | 1.74x |
| Range loop | 2,336 ms | 835 ms | 2.80x |
| UDF over 10,000 cells | 44 ms | 60 ms | 0.74x |
| Late-bound Collection (fifty thousand Adds and Counts) | 19 ms | 63 ms | 0.30x |
| Late-bound Excel (twenty thousand property reads) | 57 ms | 86 ms | 0.67x |
