using System.Globalization;
using System.Text;

namespace VbaNg.Compiler.Emit;

/// <summary>
/// Writes generated C# with #line directives (ARCHITECTURE.md D2). Every VBA statement is
/// preceded by a directive naming its source line, and generated scaffolding (braces, hoisted
/// declarations, dispatch plumbing) sits under #line hidden so the debugger never stops on it.
/// Output uses "\n" and four-space indents, so identical inputs give identical bytes (CLAUDE.md R10).
/// </summary>
internal sealed class CSharpWriter(string sourcePath)
{
    private readonly StringBuilder text = new();
    private bool hidden;
    private int line = -1;

    public int Indent { get; set; }

    public void Line(string content = "")
    {
        if (content.Length > 0)
        {
            text.Append(' ', Indent * 4);
        }

        text.Append(content).Append('\n');
    }

    /// <summary>Maps the following lines to a VBA source line; repeated for the same line only when something hidden intervened.</summary>
    public void MapLine(int sourceLine)
    {
        if (!hidden && line == sourceLine)
        {
            return;
        }

        hidden = false;
        line = sourceLine;
        text.Append("#line ").Append(sourceLine.ToString(CultureInfo.InvariantCulture)).Append(" \"").Append(sourcePath).Append("\"\n");
    }

    public void Hidden()
    {
        if (hidden)
        {
            return;
        }

        hidden = true;
        line = -1;
        text.Append("#line hidden\n");
    }

    /// <summary>A line that belongs to no VBA statement: braces, declarations, plumbing.</summary>
    public void HiddenLine(string content)
    {
        Hidden();
        Line(content);
    }

    /// <summary>A line that belongs to a VBA statement.</summary>
    public void MappedLine(int sourceLine, string content)
    {
        MapLine(sourceLine);
        Line(content);
    }

    public void Open()
    {
        HiddenLine("{");
        Indent++;
    }

    public void Close(string suffix = "")
    {
        Indent--;
        HiddenLine("}" + suffix);
    }

    public override string ToString() => text.ToString();
}
