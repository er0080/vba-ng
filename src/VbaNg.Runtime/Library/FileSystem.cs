using System.Globalization;
using System.Text;

namespace VbaNg.Runtime.Library;

/// <summary>Open mode of a file (MS-VBAL 5.4.5.1).</summary>
public enum OpenMode
{
    Input,
    Output,
    Append,
    Random,
    Binary,
}

/// <summary>The Access clause of Open.</summary>
public enum FileAccessMode
{
    Default,
    Read,
    Write,
    ReadWrite,
}

/// <summary>The lock clause of Open.</summary>
public enum FileLockMode
{
    Default,
    Shared,
    LockRead,
    LockWrite,
    LockReadWrite,
}

/// <summary>
/// The file statements and the FileSystem module of the VBA standard library (MS-VBAL 5.4.5,
/// 6.1.2.?): files are open per file number on the current thread, text goes through the
/// system ANSI code page with CRLF line ends as VBA writes it, Binary and Random modes lay
/// values out the way VBA does, and every error carries VBA's number (FileSystem golden).
/// </summary>
public static partial class FileSystem
{
    private const int MaxFileNumber = 511;
    private const int DefaultRecordLength = 128;

    [ThreadStatic]
    private static Dictionary<int, FileChannel>? channels;

    private static Dictionary<int, FileChannel> Channels => channels ??= [];

    /// <summary>The ANSI code page VBA uses for text files.</summary>
    internal static Encoding Ansi { get; } = LoadAnsi();

    /// <summary>FreeFile([rangenumber]): the lowest free number in 1 to 255, or 256 to 511 for any other range (golden: FreeFile(2) is 256); 67 when none is free.</summary>
    public static short FreeFile(in Variant rangeNumber)
    {
        var range = rangeNumber.IsMissing ? 0 : Coerce.ToInt32(rangeNumber);
        var (first, last) = range == 0 ? (1, 255) : (256, MaxFileNumber);
        for (var number = first; number <= last; number++)
        {
            if (!Channels.ContainsKey(number))
            {
                return (short)number;
            }
        }

        throw new VbaException(67);
    }

    /// <summary>Open pathname For mode [Access access] [lock] As [#]filenumber [Len = reclength] (MS-VBAL 5.4.5.1).</summary>
    public static void Open(in Variant pathName, OpenMode mode, FileAccessMode access, FileLockMode lockMode, in Variant fileNumber, in Variant recordLength)
    {
        var number = Coerce.ToInt32(fileNumber);
        if (number is < 1 or > MaxFileNumber)
        {
            throw new VbaException(52);
        }

        if (Channels.ContainsKey(number))
        {
            throw new VbaException(55);
        }

        var path = Coerce.ToString(pathName);
        var length = recordLength.IsMissing ? DefaultRecordLength : Coerce.ToInt32(recordLength);
        if (length <= 0 && mode == OpenMode.Random)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        Channels[number] = FileChannel.Open(path, mode, access, lockMode, length);
    }

    /// <summary>Close [#filenumber, ...]: the named files, or every open file when none is named.</summary>
    public static void Close(ReadOnlySpan<Variant> fileNumbers)
    {
        if (fileNumbers.Length == 0)
        {
            Reset();
            return;
        }

        foreach (var fileNumber in fileNumbers)
        {
            var number = Coerce.ToInt32(fileNumber);
            if (number is < 1 or > MaxFileNumber)
            {
                throw new VbaException(52);
            }

            if (Channels.Remove(number, out var channel))
            {
                channel.Dispose();
            }
        }
    }

    /// <summary>Reset: closes every file opened by Open.</summary>
    public static void Reset()
    {
        foreach (var channel in Channels.Values)
        {
            channel.Dispose();
        }

        Channels.Clear();
    }

    /// <summary>Print #filenumber, output list (MS-VBAL 5.4.5.6): the text Debug.Print would show, ended by CRLF unless the list ends in a separator.</summary>
    public static void Print(in Variant fileNumber, PrintList list, bool endLine)
    {
        ArgumentNullException.ThrowIfNull(list);
        Channel(fileNumber, requireOutput: true).Print(list, endLine);
    }

    /// <summary>Write #filenumber, items (MS-VBAL 5.4.5.7): strings quoted, the other types marked, commas between, CRLF at the end unless the list ends in a separator.</summary>
    public static void Write(in Variant fileNumber, ReadOnlySpan<Variant> items, bool endLine)
    {
        var text = new StringBuilder();
        for (var i = 0; i < items.Length; i++)
        {
            if (i > 0)
            {
                text.Append(',');
            }

            text.Append(WriteForm(items[i]));
        }

        Channel(fileNumber, requireOutput: true).Write(text.ToString(), endLine, items.Length > 0);
    }

    /// <summary>Input #filenumber: the next item as the variable of the given type receives it (MS-VBAL 5.4.5.8).</summary>
    public static Variant Input(in Variant fileNumber, VarType target) => Channel(fileNumber, requireInput: true).InputItem(target);

    /// <summary>Line Input #filenumber: the next line without its line end (MS-VBAL 5.4.5.4); error 62 past the end.</summary>
    public static VbaString LineInput(in Variant fileNumber) => VbaString.Temporary(Channel(fileNumber, requireInput: true).LineInput());

    /// <summary>Input(number, [#]filenumber): the next characters; error 62 when fewer remain.</summary>
    public static VbaString InputChars(in Variant count, in Variant fileNumber) => VbaString.Temporary(Channel(fileNumber, requireInput: true).InputChars(Coerce.ToInt32(count)));

    /// <summary>InputB(number, [#]filenumber): the next bytes as a byte string, a temporary of the statement; error 62 when fewer remain (FileSystem golden).</summary>
    public static VbaString InputBytes(in Variant count, in Variant fileNumber)
    {
        var channel = Channel(fileNumber);
        if (channel.Mode is not (OpenMode.Input or OpenMode.Binary))
        {
            throw new VbaException(54);
        }

        var text = VbaString.AllocBytes(channel.InputBytes(Coerce.ToInt32(count)));
        ObjectRefs.OwnedString(text.Pointer);
        return text;
    }

    public static bool Eof(in Variant fileNumber) => Channel(fileNumber).Eof;

    public static int Lof(in Variant fileNumber) => checked((int)Channel(fileNumber).Length);

    public static int Loc(in Variant fileNumber) => checked((int)Channel(fileNumber).Loc);

    public static int SeekPosition(in Variant fileNumber) => checked((int)Channel(fileNumber).SeekPosition);

    /// <summary>Seek #filenumber, position: the next byte (Binary, sequential) or record (Random), 1-based; 0 or less raises 63.</summary>
    public static void Seek(in Variant fileNumber, in Variant position) => Channel(fileNumber).Seek(Coerce.ToInt64(position));

    /// <summary>Lock or Unlock #filenumber [, start [To end]].</summary>
    public static void Lock(in Variant fileNumber, in Variant from, in Variant to, bool unlock) => Channel(fileNumber).Lock(from, to, unlock);

    /// <summary>Width #filenumber, width: the line width Print # wraps at; 0 means none.</summary>
    public static void Width(in Variant fileNumber, in Variant width) => Channel(fileNumber, requireOutput: true).Width = Coerce.ToInt32(width);

    /// <summary>Get #filenumber, [recnumber], variable for a scalar or a String: the value read; a String reads its current length (Binary) or a length-prefixed string (Random).</summary>
    public static Variant Get(in Variant fileNumber, in Variant record, in Variant current, VarType type, int fixedLength) =>
        Channel(fileNumber, requireRecord: true).Get(record, current, type, fixedLength);

    /// <summary>Get #filenumber, [recnumber], variable for a user-defined type: its fields read in place, through a box written back whatever the read raises (ROADMAP.md M7 C3).</summary>
    public static void GetRecord<T>(in Variant fileNumber, in Variant record, ref T target)
        where T : struct, IVbaRecord
    {
        var channel = Channel(fileNumber, requireRecord: true);
        object box = target;
        try
        {
            channel.GetRecord(record, box);
        }
        finally
        {
            target = (T)box;
        }
    }

    public static void GetArray(in Variant fileNumber, in Variant record, VbaArray target) =>
        Channel(fileNumber, requireRecord: true).GetArray(record, target);

    /// <summary>Put #filenumber, [recnumber], variable for a scalar or a String.</summary>
    public static void Put(in Variant fileNumber, in Variant record, in Variant value, VarType type, int fixedLength) =>
        Channel(fileNumber, requireRecord: true).Put(record, value, type, fixedLength);

    public static void PutRecord<T>(in Variant fileNumber, in Variant record, in T value)
        where T : struct, IVbaRecord =>
        Channel(fileNumber, requireRecord: true).PutRecord(record, value);

    public static void PutArray(in Variant fileNumber, in Variant record, VbaArray value) =>
        Channel(fileNumber, requireRecord: true).PutArray(record, value);

    /// <summary>Len of a user-defined type: the bytes Put writes for it, as an Integer (golden: Len of a record).</summary>
    public static short RecordLength<T>(in T record)
        where T : struct, IVbaRecord =>
        checked((short)RecordLayout.Size(record));

    /// <summary>LenB of a user-defined type: its size in memory, fields aligned and strings in UTF-16 (golden: LenB of a record).</summary>
    public static short RecordLengthB<T>(in T record)
        where T : struct, IVbaRecord =>
        checked((short)RecordLayout.MemorySize(record));

    private static FileChannel Channel(in Variant fileNumber, bool requireOutput = false, bool requireInput = false, bool requireRecord = false)
    {
        var number = Coerce.ToInt32(fileNumber);
        if (!Channels.TryGetValue(number, out var channel))
        {
            throw new VbaException(52);
        }

        if ((requireOutput && channel.Mode is not (OpenMode.Output or OpenMode.Append))
            || (requireInput && channel.Mode != OpenMode.Input)
            || (requireRecord && channel.Mode is not (OpenMode.Binary or OpenMode.Random)))
        {
            throw new VbaException(54);
        }

        return channel;
    }

    /// <summary>The text Write # produces for one item (MS-VBAL 5.4.5.7).</summary>
    internal static string WriteForm(in Variant value)
    {
        switch (value.Type)
        {
            case VarType.Empty:
                return string.Empty;
            case VarType.Null:
                return "#NULL#";
            case VarType.Boolean:
                return value.AsBoolean() ? "#TRUE#" : "#FALSE#";
            case VarType.String:
                return "\"" + value.AsString().Replace("\"", "\"\"", StringComparison.Ordinal) + "\"";
            case VarType.Error:
                return "#ERROR " + value.AsError().Number.ToString(CultureInfo.InvariantCulture) + "#";
            case VarType.Date:
                {
                    // A date with a time of day carries both parts, even on day zero (golden: TimeSerial writes #1899-12-30 06:07:08#).
                    var date = value.AsDate();
                    var hasTime = Math.Abs(date.Serial - Math.Truncate(date.Serial)) > 1e-9;
                    var time = date.ToDateTime();
                    return hasTime
                        ? time.ToString("'#'yyyy-MM-dd HH:mm:ss'#'", CultureInfo.InvariantCulture)
                        : time.ToString("'#'yyyy-MM-dd'#'", CultureInfo.InvariantCulture);
                }

            case VarType.Object:
            case VarType.Array:
            case VarType.UserDefinedType:
                throw VbaErrors.TypeMismatch();
            default:
                // Numbers take the form of Str without its leading blank (golden: 0.1 writes .1).
                return Coerce.ToString(Conversion.Str(value)).TrimStart(' ');
        }
    }

    private static Encoding LoadAnsi()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        try
        {
            return Encoding.GetEncoding(CultureInfo.CurrentCulture.TextInfo.ANSICodePage);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException)
        {
            return Encoding.Latin1;
        }
    }
}
