using Xunit.Sdk;
using Xunit.v3;

// One Excel at a time. xUnit parallelizes collections by default, which starts a hidden Excel
// per test class: the instances contend for COM, and Cli_TestAndRun_ReachTheRunningExcel
// disqualifies itself because it finds another instance, so a full run never exercised the CLI
// at all. The suite runs serially instead (CLAUDE.md R17).
[assembly: Parallelization(Mode = ParallelMode.None)]
