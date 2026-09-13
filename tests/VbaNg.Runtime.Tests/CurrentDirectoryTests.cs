using System.Runtime.InteropServices;

using Xunit;

using FileSystem = VbaNg.Runtime.Library.FileSystem;

namespace VbaNg.Runtime.Tests;

/// <summary>
/// CurDir returns the current directory as the operating system holds it. .NET's
/// Environment.CurrentDirectory expands an 8.3 short name on the way out, which VBA does not, so a
/// ChDir to a short path followed by <c>CurDir = path</c> printed False where VBA prints True. A CI
/// runner found it: its TEMP is C:\Users\RUNNER~1\AppData\Local\Temp (FileSystem golden).
/// </summary>
[Collection(Name)]
public sealed partial class CurrentDirectoryTests
{
    /// <summary>The process has one current directory, so these never run beside each other.</summary>
    public const string Name = "Current directory";

    [Fact]
    public void CurDir_AfterChDirToAShortPath_KeepsTheShortForm()
    {
        var longPath = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
        var shortPath = ShortPathOf(longPath);
        Assert.SkipWhen(shortPath is null || shortPath.Equals(longPath, StringComparison.OrdinalIgnoreCase), "This volume keeps no 8.3 name for Program Files.");

        var previous = Environment.CurrentDirectory;
        try
        {
            FileSystem.ChDir(Variant.FromString(shortPath));
            Assert.Equal(shortPath, FileSystem.CurDir(Variant.Missing).ToString());
        }
        finally
        {
            Environment.CurrentDirectory = previous;
        }
    }

    private static unsafe string? ShortPathOf(string path)
    {
        var buffer = stackalloc char[260];
        var length = GetShortPathName(path, buffer, 260);
        return length is 0 or > 260 ? null : new string(buffer, 0, (int)length);
    }

    [LibraryImport("kernel32.dll", EntryPoint = "GetShortPathNameW", StringMarshalling = StringMarshalling.Utf16)]
    private static unsafe partial uint GetShortPathName(string longPath, char* shortPath, uint bufferLength);
}
