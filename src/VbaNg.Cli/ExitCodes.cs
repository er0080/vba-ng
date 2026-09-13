namespace VbaNg.Cli;

/// <summary>Exit codes (ARCHITECTURE.md section 8, "Conventions").</summary>
internal static class ExitCodes
{
    public const int Ok = 0;
    public const int CompileErrors = 1;
    public const int Usage = 2;
    public const int ExcelUnreachable = 3;
    public const int RuntimeError = 4;
    public const int TestsFailed = 5;
}
