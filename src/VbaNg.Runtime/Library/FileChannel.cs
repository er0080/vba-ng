using System.Globalization;

namespace VbaNg.Runtime.Library;

/// <summary>
/// One open file (MS-VBAL 5.4.5): the stream, its mode, and the position bookkeeping Loc, Seek,
/// and EOF report. Sequential text is read and written through the ANSI code page byte by byte,
/// so positions are byte positions as VBA counts them (FileSystem golden).
/// </summary>
internal sealed class FileChannel : IDisposable
{
    private readonly FileStream stream;
    private readonly int recordLength;
    private int column;
    private long lastRecord;
    private bool pastEnd;

    private FileChannel(FileStream stream, OpenMode mode, int recordLength)
    {
        this.stream = stream;
        Mode = mode;
        this.recordLength = recordLength;
    }

    public OpenMode Mode { get; }

    public int Width { get; set; }

    public long Length => stream.Length;

    /// <summary>EOF: at the end of a sequential Input file; after a Get that ran short in Binary and Random modes.</summary>
    public bool Eof => Mode == OpenMode.Input ? stream.Position >= stream.Length : pastEnd;

    /// <summary>Loc: for sequential files the 128-byte block the next byte falls in, 1 right after Open and 2 from byte 129 on (FileSystem golden); the last byte read or written for Binary; the last record for Random.</summary>
    public long Loc => Mode switch
    {
        OpenMode.Binary => stream.Position,
        OpenMode.Random => lastRecord,
        _ => stream.Position / 128 + 1,
    };

    /// <summary>Seek(): the next byte (1-based), or the next record for Random.</summary>
    public long SeekPosition => Mode == OpenMode.Random ? (stream.Position / recordLength) + 1 : stream.Position + 1;

    public static FileChannel Open(string path, OpenMode mode, FileAccessMode access, FileLockMode lockMode, int recordLength)
    {
        if (path.Length == 0 || path.IndexOfAny(Path.GetInvalidPathChars()) >= 0 || Path.GetFileName(path).IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
        {
            throw new VbaException(52);
        }

        // Binary and Random create a missing file whatever the Access clause says (golden: Access Read on a missing file).
        var fileMode = mode switch
        {
            OpenMode.Input => FileMode.Open,
            OpenMode.Output => FileMode.Create,
            OpenMode.Append => FileMode.Append,
            _ => FileMode.OpenOrCreate,
        };
        var fileAccess = mode switch
        {
            OpenMode.Input => FileAccess.Read,
            OpenMode.Output or OpenMode.Append => FileAccess.Write,
            _ => access switch
            {
                FileAccessMode.Read => FileAccess.Read,
                FileAccessMode.Write => FileAccess.Write,
                _ => FileAccess.ReadWrite,
            },
        };
        var share = lockMode switch
        {
            FileLockMode.LockRead => FileShare.Write,
            FileLockMode.LockWrite => FileShare.Read,
            FileLockMode.LockReadWrite => FileShare.None,
            _ => FileShare.ReadWrite,
        };

        FileStream stream;
        try
        {
            stream = new FileStream(path, fileMode, fileAccess, share, 4096, FileOptions.None);
        }
        catch (FileNotFoundException)
        {
            throw new VbaException(53);
        }
        catch (DirectoryNotFoundException)
        {
            throw new VbaException(76);
        }
        catch (UnauthorizedAccessException)
        {
            throw new VbaException(75);
        }
        catch (IOException)
        {
            throw new VbaException(70);
        }
        catch (ArgumentException)
        {
            throw new VbaException(52);
        }
        catch (NotSupportedException)
        {
            throw new VbaException(52);
        }

        return new FileChannel(stream, mode, recordLength);
    }

    public void Dispose() => stream.Dispose();

    // Sequential output.

    /// <summary>
    /// Print #: the items of the list, then CRLF unless the list ended in a separator. Under a
    /// Width, an item that would pass the width starts a new line first; a single item longer
    /// than the width is written whole (golden: Width wraps Print # lines).
    /// </summary>
    public void Print(PrintList list, bool endLine)
    {
        foreach (var segment in list.Segments)
        {
            if (Width > 0 && column > 0 && column + segment.Length > Width)
            {
                WriteText("\r\n");
                column = 0;
            }

            WriteText(segment);
            column += segment.Length;
        }

        if (endLine)
        {
            WriteText("\r\n");
            column = 0;
        }
    }

    /// <summary>Write #: the items already formatted; a list ending in a separator ends with a comma instead of CRLF.</summary>
    public void Write(string text, bool endLine, bool hasItems)
    {
        WriteText(text);
        if (endLine)
        {
            WriteText("\r\n");
        }
        else if (hasItems)
        {
            WriteText(",");
        }
    }

    private void WriteText(string text)
    {
        var bytes = FileSystem.Ansi.GetBytes(text);
        stream.Write(bytes, 0, bytes.Length);
    }

    // Sequential input.

    /// <summary>The bytes up to a CR or a CRLF, which ends one line; a lone LF is part of the line (LineInput cases).</summary>
    public string LineInput()
    {
        if (stream.Position >= stream.Length)
        {
            throw new VbaException(62);
        }

        var bytes = new List<byte>();
        while (true)
        {
            var b = stream.ReadByte();
            if (b < 0)
            {
                break;
            }

            if (b == '\r')
            {
                if (PeekByte() == '\n')
                {
                    stream.ReadByte();
                }

                break;
            }

            bytes.Add((byte)b);
        }

        return FileSystem.Ansi.GetString(bytes.ToArray());
    }

    /// <summary>Input(n, #f): the next n bytes as characters, line ends included; error 62 when fewer remain.</summary>
    /// <summary>
    /// InputB(number, #filenumber): the next bytes as they are, for a byte string (a BSTR holds any
    /// byte count, D20). In Input mode fewer remaining raises 62; in Binary mode the bytes past the
    /// end come back as zeros and the position stops at the end (FileSystem golden).
    /// </summary>
    public byte[] InputBytes(int count)
    {
        if (count < 0)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        if (Mode != OpenMode.Binary && stream.Position + count > stream.Length)
        {
            throw new VbaException(62);
        }

        var bytes = new byte[count];
        var read = 0;
        while (read < count && stream.Read(bytes, read, count - read) is var chunk and > 0)
        {
            read += chunk;
        }

        // A Binary read that ran out sets EOF, as Get past the end does.
        pastEnd |= read < count;
        return bytes;
    }

    public string InputChars(int count)
    {
        if (count < 0)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        if (stream.Position + count > stream.Length)
        {
            throw new VbaException(62);
        }

        var bytes = new byte[count];
        var read = stream.Read(bytes, 0, count);
        return FileSystem.Ansi.GetString(bytes, 0, read);
    }

    /// <summary>
    /// One Input # item (MS-VBAL 5.4.5.8): leading blanks and line ends skipped, then quoted text,
    /// or bare text up to the next comma or line end with trailing blanks removed; #TRUE#, #FALSE#,
    /// #NULL#, #ERROR n#, and #date# are the marks Write # leaves.
    /// </summary>
    public Variant InputItem(VarType target)
    {
        if (stream.Position >= stream.Length)
        {
            throw new VbaException(62);
        }

        SkipInputSpaces();
        var bytes = new List<byte>();
        var quoted = false;
        if (PeekByte() == '"')
        {
            stream.ReadByte();
            quoted = true;
            while (true)
            {
                var b = stream.ReadByte();
                if (b < 0 || b == '"')
                {
                    break;
                }

                bytes.Add((byte)b);
            }

            SkipToDelimiter();
        }
        else
        {
            while (true)
            {
                var b = PeekByte();
                if (b < 0 || b == ',' || b == '\r' || b == '\n')
                {
                    break;
                }

                stream.ReadByte();
                bytes.Add((byte)b);
            }

            ConsumeDelimiter();
        }

        var raw = FileSystem.Ansi.GetString(bytes.ToArray());
        return Classify(quoted ? raw : raw.TrimEnd(' ', '\t'), quoted, target);
    }

    private static Variant Classify(string item, bool quoted, VarType target)
    {
        if (quoted)
        {
            return Coerce.ToType(Variant.FromString(item), target == VarType.Variant ? VarType.String : target);
        }

        if (item.Length == 0)
        {
            return target == VarType.Variant ? Variant.Empty : Coerce.ToType(Variant.Empty, target);
        }

        if (item.Length >= 2 && item[0] == '#' && item[^1] == '#')
        {
            var inner = item[1..^1];
            Variant marked;
            if (inner.Equals("TRUE", StringComparison.OrdinalIgnoreCase))
            {
                marked = Variant.True;
            }
            else if (inner.Equals("FALSE", StringComparison.OrdinalIgnoreCase))
            {
                marked = Variant.False;
            }
            else if (inner.Equals("NULL", StringComparison.OrdinalIgnoreCase))
            {
                marked = Variant.Null;
            }
            else if (inner.StartsWith("ERROR ", StringComparison.OrdinalIgnoreCase) && int.TryParse(inner[6..], NumberStyles.Integer, CultureInfo.InvariantCulture, out var code))
            {
                marked = Variant.FromError(ErrorValue.FromNumber(code));
            }
            else if (DateTime.TryParse(inner, CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsed))
            {
                marked = Variant.FromDate(VbaDate.FromDateTime(parsed));
            }
            else
            {
                marked = Variant.FromString(item);
            }

            return target == VarType.Variant ? marked : Coerce.ToType(marked, target);
        }

        if (target == VarType.String)
        {
            return Variant.FromString(item);
        }

        if (target != VarType.Variant)
        {
            // Text that is not a number leaves a numeric variable at zero (golden: "abc" into a Long).
            return double.TryParse(item, NumberStyles.Float, CultureInfo.InvariantCulture, out var number)
                ? Coerce.ToType(Variant.FromDouble(number), target)
                : Coerce.ToType(Variant.Empty, target);
        }

        return ClassifyNumber(item);
    }

    /// <summary>
    /// A bare number into a Variant: a decimal without an exponent becomes Currency (golden: "12.5"),
    /// an integer the smallest of Integer and Long that holds it, anything else Double, and text
    /// that is not a number stays a String.
    /// </summary>
    private static Variant ClassifyNumber(string item)
    {
        if (!double.TryParse(item, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
        {
            return Variant.FromString(item);
        }

        var hasExponent = item.Contains('e', StringComparison.OrdinalIgnoreCase);
        var hasPoint = item.Contains('.', StringComparison.Ordinal);
        if (!hasExponent && !hasPoint)
        {
            if (number is >= short.MinValue and <= short.MaxValue)
            {
                return Variant.FromInt16((short)number);
            }

            if (number is >= int.MinValue and <= int.MaxValue)
            {
                return Variant.FromInt32((int)number);
            }
        }
        else if (!hasExponent && decimal.TryParse(item, NumberStyles.Float, CultureInfo.InvariantCulture, out var exact)
            && exact == decimal.Round(exact, 4) && Math.Abs(exact) < 922337203685477.5807m)
        {
            return Variant.FromCurrency(Currency.FromDecimal(exact));
        }

        return Variant.FromDouble(number);
    }

    private int PeekByte()
    {
        if (stream.Position >= stream.Length)
        {
            return -1;
        }

        var b = stream.ReadByte();
        stream.Position--;
        return b;
    }

    /// <summary>Blanks before an item are skipped; a line end is not, it ends an empty item (golden: Write # of Empty reads back as Empty).</summary>
    private void SkipInputSpaces()
    {
        while (true)
        {
            var b = PeekByte();
            if (b == ' ' || b == '\t')
            {
                stream.ReadByte();
                continue;
            }

            break;
        }
    }

    private void SkipToDelimiter()
    {
        while (true)
        {
            var b = PeekByte();
            if (b < 0 || b == ',' || b == '\r' || b == '\n')
            {
                break;
            }

            stream.ReadByte();
        }

        ConsumeDelimiter();
    }

    private void ConsumeDelimiter()
    {
        var b = PeekByte();
        if (b == ',')
        {
            stream.ReadByte();
        }
        else if (b == '\r')
        {
            stream.ReadByte();
            if (PeekByte() == '\n')
            {
                stream.ReadByte();
            }
        }
        else if (b == '\n')
        {
            stream.ReadByte();
        }
    }

    // Positions.

    public void Seek(long position)
    {
        if (position < 1)
        {
            throw new VbaException(63);
        }

        stream.Position = Mode == OpenMode.Random ? (position - 1) * recordLength : position - 1;
        pastEnd = false;
    }

    public void Lock(in Variant from, in Variant to, bool unlock)
    {
        long start;
        long length;
        if (from.IsMissing)
        {
            start = 0;
            length = long.MaxValue / 2;
        }
        else
        {
            var first = Coerce.ToInt64(from);
            var last = to.IsMissing ? first : Coerce.ToInt64(to);
            if (first < 1 || last < first)
            {
                throw new VbaException(63);
            }

            var unit = Mode == OpenMode.Random ? recordLength : 1;
            start = (first - 1) * unit;
            length = (last - first + 1) * unit;
        }

        if (OperatingSystem.IsMacOS())
        {
            return;
        }

        try
        {
            if (unlock)
            {
                stream.Unlock(start, length);
            }
            else
            {
                stream.Lock(start, length);
            }
        }
        catch (IOException)
        {
            throw new VbaException(70);
        }
    }

    // Binary and Random.

    public Variant Get(in Variant record, in Variant current, VarType type, int fixedLength)
    {
        Position(record);
        return Read(type, current, fixedLength);
    }

    public void GetRecord(in Variant record, object target)
    {
        Position(record);
        var start = stream.Position;
        RecordLayout.Read(this, target);
        EndRecord(start);
    }

    public void GetArray(in Variant record, VbaArray target)
    {
        Position(record);
        var start = stream.Position;
        foreach (var index in RecordLayout.Indices(target))
        {
            target.Set(index, RecordLayout.Decode(this, target.ElementType, target.Get(index), 0, lengthPrefixed: false));
        }

        EndRecord(start);
    }

    public void Put(in Variant record, in Variant value, VarType type, int fixedLength)
    {
        Position(record);
        WriteRecord(RecordLayout.Encode(value, type, fixedLength, Mode == OpenMode.Random));
    }

    public void PutRecord(in Variant record, object value)
    {
        Position(record);
        WriteRecord(RecordLayout.EncodeRecord(value));
    }

    public void PutArray(in Variant record, VbaArray value)
    {
        Position(record);
        var bytes = new List<byte>();
        foreach (var index in RecordLayout.Indices(value))
        {
            bytes.AddRange(RecordLayout.Encode(value.Get(index), value.ElementType, 0, lengthPrefixed: false));
        }

        WriteRecord(bytes.ToArray());
    }

    /// <summary>
    /// Random writes the value at the start of its record and moves to the next record without
    /// padding the file (golden: the default record length leaves LOF at 132); a value longer than
    /// the record raises 59. Binary writes the bytes where the position is.
    /// </summary>
    private void WriteRecord(byte[] bytes)
    {
        if (Mode == OpenMode.Random && bytes.Length > recordLength)
        {
            throw new VbaException(59);
        }

        var start = stream.Position;
        stream.Write(bytes, 0, bytes.Length);
        EndRecord(start);
    }

    private void Position(in Variant record)
    {
        if (record.IsMissing)
        {
            return;
        }

        var number = Coerce.ToInt64(record);
        if (number < 1)
        {
            throw new VbaException(63);
        }

        stream.Position = Mode == OpenMode.Random ? (number - 1) * recordLength : number - 1;
        pastEnd = false;
    }

    private void EndRecord(long start)
    {
        if (Mode == OpenMode.Random)
        {
            lastRecord = (start / recordLength) + 1;
            stream.Position = start + recordLength;
        }
    }

    internal Variant Read(VarType type, in Variant current, int fixedLength)
    {
        var start = stream.Position;
        var value = RecordLayout.Decode(this, type, current, fixedLength, Mode == OpenMode.Random);
        EndRecord(start);
        return value;
    }

    /// <summary>Reads exactly <paramref name="count"/> bytes; fewer sets EOF and leaves the rest zero, as a short Get does.</summary>
    internal byte[] ReadBytes(int count)
    {
        var bytes = new byte[count];
        var total = 0;
        while (total < count)
        {
            var read = stream.Read(bytes, total, count - total);
            if (read <= 0)
            {
                pastEnd = true;
                break;
            }

            total += read;
        }

        return bytes;
    }
}
