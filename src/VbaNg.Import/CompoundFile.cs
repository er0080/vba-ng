using System.Buffers.Binary;
using System.Text;

namespace VbaNg.Import;

/// <summary>
/// A read-only Compound File Binary reader (MS-CFB), which is what a <c>vbaProject.bin</c> and an
/// <c>.xls</c> workbook are: a file system of storages and streams inside one file. Only what the
/// importer needs is implemented, reading whole streams by path.
/// </summary>
public sealed class CompoundFile
{
    private const int HeaderSize = 512;
    private const uint EndOfChain = 0xFFFFFFFE;
    private const uint FreeSector = 0xFFFFFFFF;
    private const int DirectoryEntrySize = 128;

    private static readonly byte[] Signature = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

    private readonly byte[] data;
    private readonly int sectorSize;
    private readonly int miniSectorSize;
    private readonly int miniStreamCutoff;
    private readonly uint[] fat;
    private readonly uint[] miniFat;
    private readonly List<DirectoryEntry> directory = [];
    private readonly byte[] miniStream;

    private CompoundFile(byte[] data)
    {
        this.data = data;
        if (data.Length < HeaderSize || !data.AsSpan(0, 8).SequenceEqual(Signature))
        {
            throw new InvalidDataException("Not a compound file: the signature is missing.");
        }

        sectorSize = 1 << ReadUInt16(0x1E);
        miniSectorSize = 1 << ReadUInt16(0x20);
        miniStreamCutoff = (int)ReadUInt32(0x38);
        if (sectorSize is not (512 or 4096) || miniSectorSize != 64)
        {
            throw new InvalidDataException($"Unsupported compound file sector size {sectorSize.ToString(System.Globalization.CultureInfo.InvariantCulture)}.");
        }

        fat = ReadFat();
        miniFat = ReadChainValues(ReadUInt32(0x3C), (int)ReadUInt32(0x40));
        ReadDirectory(ReadUInt32(0x30));
        miniStream = directory.Count > 0 ? ReadSectorChain(directory[0].StartSector, directory[0].Size, mini: false) : [];
    }

    /// <summary>The names of the entries directly under a storage, in directory order; the root is the empty path.</summary>
    public IEnumerable<string> Entries(string storagePath)
    {
        var parent = Find(storagePath) ?? throw new FileNotFoundException($"Storage not found in the compound file: '{storagePath}'.");
        foreach (var entry in directory)
        {
            if (entry.Parent == parent.Id && entry.Id != parent.Id)
            {
                yield return entry.Name;
            }
        }
    }

    public static CompoundFile Read(byte[] bytes)
    {
        ArgumentNullException.ThrowIfNull(bytes);
        return new CompoundFile(bytes);
    }

    public static CompoundFile Open(string path) => Read(File.ReadAllBytes(path));

    /// <summary>True when a stream or storage exists at the path, whose parts are separated by "/" and compared case-insensitively.</summary>
    public bool Exists(string path) => Find(path) is not null;

    /// <summary>The whole content of a stream; the path's parts are separated by "/".</summary>
    public byte[] ReadStream(string path)
    {
        var entry = Find(path) ?? throw new FileNotFoundException($"Stream not found in the compound file: '{path}'.");
        if (entry.Type != 2)
        {
            throw new InvalidDataException($"'{path}' is a storage, not a stream.");
        }

        return entry.Size < miniStreamCutoff
            ? ReadMiniChain(entry.StartSector, entry.Size)
            : ReadSectorChain(entry.StartSector, entry.Size, mini: false);
    }

    private DirectoryEntry? Find(string path)
    {
        ArgumentNullException.ThrowIfNull(path);
        if (directory.Count == 0)
        {
            return null;
        }

        var current = directory[0];
        foreach (var part in path.Split('/', StringSplitOptions.RemoveEmptyEntries))
        {
            var child = directory.FirstOrDefault(e => e.Parent == current.Id && e.Name.Equals(part, StringComparison.OrdinalIgnoreCase));
            if (child is null)
            {
                return null;
            }

            current = child;
        }

        return current;
    }

    /// <summary>The FAT, read through the DIFAT: the first 109 entries live in the header, the rest in DIFAT sectors.</summary>
    private uint[] ReadFat()
    {
        var fatSectors = new List<uint>();
        for (var i = 0; i < 109; i++)
        {
            var sector = ReadUInt32(0x4C + (i * 4));
            if (sector is FreeSector or EndOfChain)
            {
                break;
            }

            fatSectors.Add(sector);
        }

        var difatSector = ReadUInt32(0x44);
        var difatCount = (int)ReadUInt32(0x48);
        var perSector = (sectorSize / 4) - 1;
        for (var i = 0; i < difatCount && difatSector is not (FreeSector or EndOfChain); i++)
        {
            var offset = SectorOffset(difatSector);
            for (var slot = 0; slot < perSector; slot++)
            {
                var sector = ReadUInt32(offset + (slot * 4));
                if (sector is FreeSector or EndOfChain)
                {
                    break;
                }

                fatSectors.Add(sector);
            }

            difatSector = ReadUInt32(offset + (perSector * 4));
        }

        var values = new uint[fatSectors.Count * (sectorSize / 4)];
        var index = 0;
        foreach (var sector in fatSectors)
        {
            var offset = SectorOffset(sector);
            for (var slot = 0; slot < sectorSize / 4; slot++)
            {
                values[index++] = ReadUInt32(offset + (slot * 4));
            }
        }

        return values;
    }

    /// <summary>Every 4-byte value of a sector chain, for the mini FAT.</summary>
    private uint[] ReadChainValues(uint start, int sectorCount)
    {
        var values = new List<uint>(sectorCount * (sectorSize / 4));
        var sector = start;
        for (var read = 0; read < sectorCount && sector is not (FreeSector or EndOfChain); read++)
        {
            var offset = SectorOffset(sector);
            for (var slot = 0; slot < sectorSize / 4; slot++)
            {
                values.Add(ReadUInt32(offset + (slot * 4)));
            }

            sector = Next(fat, sector);
        }

        return values.ToArray();
    }

    private void ReadDirectory(uint start)
    {
        var bytes = ReadSectorChain(start, long.MaxValue, mini: false);
        for (var offset = 0; offset + DirectoryEntrySize <= bytes.Length; offset += DirectoryEntrySize)
        {
            var nameLength = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(offset + 64, 2));
            var type = bytes[offset + 66];
            if (type == 0)
            {
                directory.Add(new DirectoryEntry(directory.Count, string.Empty, 0, 0xFFFFFFFF, 0, 0));
                continue;
            }

            var name = nameLength > 2 ? Encoding.Unicode.GetString(bytes, offset, nameLength - 2) : string.Empty;
            var startSector = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(offset + 116, 4));
            var size = (long)BinaryPrimitives.ReadUInt64LittleEndian(bytes.AsSpan(offset + 120, 8));
            directory.Add(new DirectoryEntry(directory.Count, name, type, startSector, size, 0xFFFFFFFF));
        }

        // The directory is a red-black tree per storage; walking it fills in each entry's parent.
        if (directory.Count > 0)
        {
            AssignParents(bytes, 0, 0);
        }
    }

    private void AssignParents(byte[] bytes, int entryId, int parentId)
    {
        var pending = new Stack<(int Entry, int Parent)>();
        pending.Push((entryId, parentId));
        var seen = new HashSet<int>();
        while (pending.Count > 0)
        {
            var (id, parent) = pending.Pop();
            if (id < 0 || id >= directory.Count || !seen.Add(id))
            {
                continue;
            }

            directory[id] = directory[id] with { Parent = (uint)parent };
            var offset = id * DirectoryEntrySize;
            var left = (int)BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(offset + 68, 4));
            var right = (int)BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(offset + 72, 4));
            var child = (int)BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(offset + 76, 4));

            // Siblings share this entry's parent; the child subtree hangs under this entry.
            pending.Push((left, parent));
            pending.Push((right, parent));
            pending.Push((child, id));
        }
    }

    private byte[] ReadSectorChain(uint start, long size, bool mini)
    {
        var unit = mini ? miniSectorSize : sectorSize;
        var chain = mini ? miniFat : fat;
        var output = new MemoryStream();
        var sector = start;
        var guard = 0;
        while (sector is not (FreeSector or EndOfChain) && output.Length < (size == long.MaxValue ? long.MaxValue : size))
        {
            var offset = mini ? (int)((long)sector * miniSectorSize) : SectorOffset(sector);
            var source = mini ? miniStream : data;
            if (offset < 0 || offset + unit > source.Length)
            {
                break;
            }

            output.Write(source, offset, unit);
            sector = Next(chain, sector);
            if (++guard > 1_000_000)
            {
                throw new InvalidDataException("The compound file has a cyclic sector chain.");
            }
        }

        var bytes = output.ToArray();
        return size != long.MaxValue && size < bytes.Length ? bytes.AsSpan(0, (int)size).ToArray() : bytes;
    }

    private byte[] ReadMiniChain(uint start, long size) => ReadSectorChain(start, size, mini: true);

    private static uint Next(uint[] chain, uint sector) => sector < chain.Length ? chain[sector] : EndOfChain;

    private int SectorOffset(uint sector) => (int)(((long)sector + 1) * sectorSize);

    private ushort ReadUInt16(int offset) => BinaryPrimitives.ReadUInt16LittleEndian(data.AsSpan(offset, 2));

    private uint ReadUInt32(int offset) => BinaryPrimitives.ReadUInt32LittleEndian(data.AsSpan(offset, 4));

    private sealed record DirectoryEntry(int Id, string Name, byte Type, uint StartSector, long Size, uint Parent);
}
