using VbaNg.Runtime.Hosting;

using Xunit;

namespace VbaNg.Runtime.Tests;

/// <summary>
/// The host command response and its transport through <c>Application.Run</c>, whose string
/// return is capped by Excel (ARCHITECTURE.md section 7): short responses travel inline, long ones
/// through <c>out/response.json</c>.
/// </summary>
public sealed class RunResponseTests : IDisposable
{
    private readonly string projectDir = Path.Combine(Path.GetTempPath(), "vbang-tests", Guid.NewGuid().ToString("N"), "Big.vbang");

    public void Dispose()
    {
        try
        {
            Directory.Delete(Path.GetDirectoryName(projectDir)!, recursive: true);
        }
        catch (IOException)
        {
        }
    }

    [Fact]
    public void ToTransportJson_ShortResponse_StaysInline()
    {
        var response = RunResponse.Success("hello");

        var json = response.ToTransportJson(projectDir);

        Assert.Equal(response.ToJson(), json);
        Assert.DoesNotContain("\"file\"", json, StringComparison.Ordinal);
        Assert.False(Directory.Exists(ProjectPaths.OutputDir(projectDir)));
        Assert.Equal(json, RunResponse.ResolveTransport(json));
    }

    [Fact]
    public void ToTransportJson_LongResponse_GoesThroughTheFile()
    {
        var output = string.Join(Environment.NewLine, Enumerable.Range(1, 600).Select(i => new string('x', 80)));
        var response = RunResponse.Failure(output, "2 tests: 1 passed, 1 failed, 0 errors. 3 ms.", null);

        var stub = response.ToTransportJson(projectDir);

        Assert.True(stub.Length <= RunResponse.MaxInlineLength, "the stub must fit the inline limit");
        var parsed = RunResponse.FromJson(stub);
        Assert.False(parsed.Ok);
        Assert.Equal(ProjectPaths.ResponsePath(projectDir), parsed.ResponseFile);
        Assert.Equal(string.Empty, parsed.Output);

        var resolved = RunResponse.FromJson(RunResponse.ResolveTransport(stub));
        Assert.Equal(response, resolved);
        Assert.Equal(output, resolved.Output);
    }

    [Fact]
    public void FromJson_RoundTripsEveryField()
    {
        var response = RunResponse.Failure("out", "err", "detail");

        Assert.Equal(response, RunResponse.FromJson(response.ToJson()));
        Assert.Null(RunResponse.FromJson(RunResponse.Success("x").ToJson()).Error);
    }
}
