namespace VbaNg.Compiler;

/// <summary>One VBA source file. <see cref="Path"/> is absolute; it ends up in <c>#line</c> directives.</summary>
public sealed record SourceFile(string Path, string Text)
{
    /// <summary>Reads a file with BOM detection, defaulting to UTF-8 (ARCHITECTURE.md section 3, "File formats").</summary>
    public static SourceFile Load(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        var fullPath = System.IO.Path.GetFullPath(path);
        return new SourceFile(fullPath, File.ReadAllText(fullPath));
    }
}
