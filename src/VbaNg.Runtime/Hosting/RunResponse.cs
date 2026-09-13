using System.Text.Json;
using System.Text.Json.Serialization;

namespace VbaNg.Runtime.Hosting;

/// <summary>
/// Result of a host command such as <c>vbang.Run</c>, passed from the add-in to the CLI as a JSON
/// string through <c>Application.Run</c> (ARCHITECTURE.md D6). Excel returns an XLL function's
/// string through an XLOPER12, which holds at most 32,767 characters, so a response longer than
/// <see cref="MaxInlineLength"/> is written to <c>out/response.json</c> in the project folder and
/// the returned JSON only names that file (section 8); <see cref="ResolveTransport"/> reads it back.
/// </summary>
public sealed record RunResponse(
    [property: JsonPropertyName("ok")] bool Ok,
    [property: JsonPropertyName("output")] string Output,
    [property: JsonPropertyName("error")] string? Error,
    [property: JsonPropertyName("detail")] string? Detail)
{
    /// <summary>The longest JSON returned inline, under Excel's 32,767-character string limit with room for the marshaling.</summary>
    public const int MaxInlineLength = 32_000;

    private static readonly JsonSerializerOptions Options = new()
    {
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    };

    /// <summary>The file holding the full response when it was too long to return inline; null otherwise.</summary>
    [JsonPropertyName("file")]
    public string? ResponseFile { get; init; }

    public static RunResponse Success(string output) => new(true, output, null, null);

    public static RunResponse Failure(string output, string error, string? detail) => new(false, output, error, detail);

    public string ToJson() => JsonSerializer.Serialize(this, Options);

    public static RunResponse FromJson(string json)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(json);
        return JsonSerializer.Deserialize<RunResponse>(json, Options)
            ?? throw new JsonException("Host command returned an empty response.");
    }

    /// <summary>The JSON to return from a host command: the response itself when it fits, otherwise a stub naming the file it was written to.</summary>
    public string ToTransportJson(string projectDir)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(projectDir);
        var json = ToJson();
        if (json.Length <= MaxInlineLength)
        {
            return json;
        }

        var path = ProjectPaths.ResponsePath(projectDir);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path, json);
        return new RunResponse(Ok, string.Empty, null, null) { ResponseFile = path }.ToJson();
    }

    /// <summary>The full response JSON behind what a host command returned: the text itself, or the content of the file it names.</summary>
    public static string ResolveTransport(string json)
    {
        var response = FromJson(json);
        return response.ResponseFile is null ? json : File.ReadAllText(response.ResponseFile);
    }
}
