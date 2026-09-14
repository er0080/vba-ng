using System.Text;

using Xunit;

using FileAccessMode = VbaNg.Runtime.Library.FileAccessMode;
using FileLockMode = VbaNg.Runtime.Library.FileLockMode;
using FileSystem = VbaNg.Runtime.Library.FileSystem;
using OpenMode = VbaNg.Runtime.Library.OpenMode;

namespace VbaNg.Runtime.Tests;

/// <summary>
/// Line Input # ends a line at CR or CRLF and nowhere else, so a lone LF stays in the line (LineInput cases). vba-ng
/// ended lines at LF too, and VBA-Web's LF-only example credentials file loaded under vbang where VBA reads it as one
/// comment line.
/// </summary>
public sealed class LineInputTests : IDisposable
{
    private readonly string path = Path.Combine(Path.GetTempPath(), "vbang-lineinput-" + Guid.NewGuid().ToString("N") + ".txt");

    public void Dispose() => File.Delete(path);

    [Fact]
    public void LineInput_LoneLf_StaysInTheLine()
    {
        Assert.Equal(["one\ntwo\n"], ReadLines("one\ntwo\n"));
    }

    [Fact]
    public void LineInput_CrAndCrLf_EachEndOneLine()
    {
        Assert.Equal(["one", "two", "three"], ReadLines("one\rtwo\r\nthree"));
    }

    [Fact]
    public void LineInput_LfBeforeCr_StaysInTheLineTheCrEnds()
    {
        Assert.Equal(["one\n", "two"], ReadLines("one\n\rtwo"));
    }

    private List<string> ReadLines(string text)
    {
        File.WriteAllBytes(path, Encoding.ASCII.GetBytes(text));
        var number = Variant.FromInt32(FileSystem.FreeFile(Variant.Missing));
        FileSystem.Open(Variant.FromString(path), OpenMode.Input, FileAccessMode.Default, FileLockMode.Default, number, Variant.Missing);
        try
        {
            var lines = new List<string>();
            while (!FileSystem.Eof(number))
            {
                lines.Add(FileSystem.LineInput(number).ToString());
            }

            return lines;
        }
        finally
        {
            FileSystem.Close([number]);
        }
    }
}
