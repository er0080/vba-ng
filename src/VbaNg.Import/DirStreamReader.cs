using System.Buffers.Binary;
using System.Text;

namespace VbaNg.Import;

/// <summary>What the <c>dir</c> stream says about the project (MS-OVBA 2.3.4.2).</summary>
public sealed record DirContents(string Name, int CodePage, IReadOnlyList<VbaReference> References, IReadOnlyList<DirModule> Modules);

/// <summary>One module's entry in the <c>dir</c> stream: where its source lives and what kind it is.</summary>
public sealed record DirModule(string Name, string StreamName, int TextOffset, bool IsDocument);

/// <summary>
/// The <c>dir</c> stream reader (MS-OVBA 2.3.4.2). Most records are an id, a size, and that many
/// bytes, but several carry a Unicode twin that sits outside the size, and REFERENCECONTROL is a
/// record of records, so every id is read with its own shape rather than by skipping the size.
/// </summary>
public static class DirStreamReader
{
    public static DirContents Read(byte[] dir)
    {
        ArgumentNullException.ThrowIfNull(dir);
        var reader = new Cursor(dir);
        var name = "VBAProject";
        var codePage = 1252;
        var references = new List<VbaReference>();
        var modules = new List<DirModule>();
        string? referenceName = null;
        string? originalLibraryId = null;
        DirModuleBuilder? module = null;

        while (reader.Remaining >= 2)
        {
            var id = reader.UInt16();
            switch (id)
            {
                case 0x0003:
                    codePage = reader.Sized() is { Length: >= 2 } bytes ? BinaryPrimitives.ReadUInt16LittleEndian(bytes) : codePage;
                    break;
                case 0x0004:
                    name = Encodings.ForCodePage(codePage).GetString(reader.Sized());
                    break;

                // A record with a Unicode twin: the reserved id, the size, and the text follow the size the record declared.
                case 0x0005 or 0x0006 or 0x000C or 0x0016:
                    {
                        var text = Encodings.ForCodePage(codePage).GetString(reader.Sized());
                        reader.UInt16();
                        reader.Sized();
                        if (id == 0x0016)
                        {
                            referenceName = text;
                        }

                        break;
                    }

                case 0x000D or 0x000E or 0x0033:
                    {
                        // REFERENCEREGISTERED, REFERENCEPROJECT, REFERENCEORIGINAL: the size covers a size-prefixed library id.
                        var libraryId = LibraryId(reader.Sized(), codePage);
                        if (id == 0x0033)
                        {
                            // The original library of a control reference: the REFERENCECONTROL that follows carries the identity to keep.
                            originalLibraryId = libraryId;
                            break;
                        }

                        references.Add(new VbaReference(referenceName ?? string.Empty, libraryId));
                        referenceName = null;
                        break;
                    }

                case 0x002F:
                    {
                        var extended = ReadControlReference(reader, codePage, ref referenceName);
                        references.Add(new VbaReference(referenceName ?? string.Empty, extended ?? originalLibraryId));
                        referenceName = null;
                        originalLibraryId = null;
                        break;
                    }
                case 0x0019:
                    module = new DirModuleBuilder { Name = Encodings.ForCodePage(codePage).GetString(reader.Sized()) };
                    break;
                case 0x0047:
                    {
                        var unicode = Encoding.Unicode.GetString(reader.Sized());
                        if (module is not null && unicode.Length > 0)
                        {
                            module.Name = unicode;
                        }

                        break;
                    }

                case 0x001A:
                    {
                        var streamName = Encodings.ForCodePage(codePage).GetString(reader.Sized());
                        reader.UInt16();
                        reader.Sized();
                        if (module is not null)
                        {
                            module.StreamName = streamName;
                        }

                        break;
                    }

                case 0x001C:
                    reader.Sized();
                    reader.UInt16();
                    reader.Sized();
                    break;
                case 0x0031:
                    {
                        var offset = reader.Sized();
                        if (module is not null && offset.Length >= 4)
                        {
                            module.TextOffset = (int)BinaryPrimitives.ReadUInt32LittleEndian(offset);
                        }

                        break;
                    }

                case 0x0021 or 0x0022:
                    reader.Skip(4);
                    if (module is not null)
                    {
                        module.IsDocument = id == 0x0022;
                    }

                    break;
                case 0x002B:
                    reader.Skip(4);
                    if (module is not null)
                    {
                        modules.Add(module.Build());
                        module = null;
                    }

                    break;
                case 0x0009:
                    // PROJECTVERSION (MS-OVBA 2.3.4.2.1.9): the four reserved bytes are followed by a major and a minor that the size does not count.
                    reader.Skip(10);
                    break;
                case 0x0010:
                    reader.Skip(4);
                    break;
                default:
                    reader.Sized();
                    break;
            }
        }

        if (module is not null)
        {
            modules.Add(module.Build());
        }

        return new DirContents(name, codePage, references, modules);
    }

    /// <summary>
    /// REFERENCECONTROL (MS-OVBA 2.3.4.2.2.3): the twiddled library, an optional name that
    /// replaces the one before the record, then the extended library, whose identity is the one to
    /// keep, followed by its original type library and a cookie inside the same size.
    /// </summary>
    private static string? ReadControlReference(Cursor reader, int codePage, ref string? referenceName)
    {
        reader.Sized();
        if (reader.PeekUInt16() == 0x0016)
        {
            reader.UInt16();
            referenceName = Encodings.ForCodePage(codePage).GetString(reader.Sized());
            reader.UInt16();
            reader.Sized();
        }

        if (reader.PeekUInt16() != 0x0030)
        {
            return null;
        }

        reader.UInt16();
        var extended = reader.Sized();
        return LibraryId(extended, codePage);
    }

    /// <summary>A reference's body starts with the size of its library id, then the id itself (MS-OVBA 2.3.4.2.2.5).</summary>
    private static string? LibraryId(byte[] body, int codePage)
    {
        if (body.Length < 4)
        {
            return null;
        }

        var length = (int)BinaryPrimitives.ReadUInt32LittleEndian(body);
        if (length <= 0 || 4 + length > body.Length)
        {
            return null;
        }

        return Encodings.ForCodePage(codePage).GetString(body, 4, length);
    }

    private sealed class DirModuleBuilder
    {
        public string Name { get; set; } = string.Empty;

        public string StreamName { get; set; } = string.Empty;

        public int TextOffset { get; set; }

        public bool IsDocument { get; set; }

        public DirModule Build() => new(Name, StreamName.Length == 0 ? Name : StreamName, TextOffset, IsDocument);
    }

    /// <summary>A forward-only reader over the stream.</summary>
    private sealed class Cursor(byte[] data)
    {
        private int position;

        public int Remaining => data.Length - position;

        public ushort UInt16()
        {
            if (Remaining < 2)
            {
                position = data.Length;
                return 0;
            }

            var value = BinaryPrimitives.ReadUInt16LittleEndian(data.AsSpan(position, 2));
            position += 2;
            return value;
        }

        public ushort PeekUInt16() => Remaining >= 2 ? BinaryPrimitives.ReadUInt16LittleEndian(data.AsSpan(position, 2)) : (ushort)0;

        /// <summary>A four-byte size followed by that many bytes.</summary>
        public byte[] Sized()
        {
            if (Remaining < 4)
            {
                position = data.Length;
                return [];
            }

            var size = (int)BinaryPrimitives.ReadUInt32LittleEndian(data.AsSpan(position, 4));
            position += 4;
            if (size < 0 || size > Remaining)
            {
                size = Remaining;
            }

            var bytes = data.AsSpan(position, size).ToArray();
            position += size;
            return bytes;
        }

        public void Skip(int count) => position = Math.Min(data.Length, position + count);
    }
}
