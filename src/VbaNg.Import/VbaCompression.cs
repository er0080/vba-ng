using System.Buffers.Binary;

namespace VbaNg.Import;

/// <summary>
/// The compressed container of MS-OVBA 2.4.1, which holds the <c>dir</c> stream and every module's
/// source: a signature byte followed by chunks that are either literal or run-length encoded, with
/// the split between a copy token's length and offset depending on how much has been decompressed.
/// </summary>
public static class VbaCompression
{
    private const byte ContainerSignature = 0x01;
    private const int ChunkSize = 4096;

    /// <summary>Decompresses a container that starts at <paramref name="start"/> (MS-OVBA 2.4.1.3.1).</summary>
    public static byte[] Decompress(byte[] compressed, int start = 0)
    {
        ArgumentNullException.ThrowIfNull(compressed);
        if (start >= compressed.Length || compressed[start] != ContainerSignature)
        {
            throw new InvalidDataException("The compressed container does not start with the 0x01 signature.");
        }

        var output = new MemoryStream();
        var position = start + 1;
        while (position + 1 < compressed.Length)
        {
            var header = BinaryPrimitives.ReadUInt16LittleEndian(compressed.AsSpan(position, 2));
            position += 2;
            var size = (header & 0x0FFF) + 3;
            var compressedChunk = (header & 0x8000) != 0;
            if ((header & 0x7000) != 0x3000)
            {
                throw new InvalidDataException("A chunk header of the compressed container has the wrong signature.");
            }

            var end = Math.Min(position + size - 2, compressed.Length);
            if (!compressedChunk)
            {
                output.Write(compressed, position, Math.Min(ChunkSize, end - position));
                position = end;
                continue;
            }

            var chunkStart = (int)output.Length;
            while (position < end)
            {
                var flags = compressed[position++];
                for (var bit = 0; bit < 8 && position < end; bit++)
                {
                    if ((flags & (1 << bit)) == 0)
                    {
                        output.WriteByte(compressed[position++]);
                        continue;
                    }

                    if (position + 1 >= end + 1)
                    {
                        break;
                    }

                    var token = BinaryPrimitives.ReadUInt16LittleEndian(compressed.AsSpan(position, 2));
                    position += 2;
                    var (lengthBits, lengthMask, offsetShift) = TokenShape((int)output.Length - chunkStart);
                    var length = (token & lengthMask) + 3;
                    var offset = (token >> offsetShift) + 1;
                    _ = lengthBits;
                    var copyFrom = (int)output.Length - offset;
                    if (copyFrom < 0)
                    {
                        throw new InvalidDataException("A copy token of the compressed container points before the chunk.");
                    }

                    var buffer = output.GetBuffer();
                    for (var i = 0; i < length; i++)
                    {
                        output.WriteByte(buffer[copyFrom + i]);
                        buffer = output.GetBuffer();
                    }
                }
            }

            position = end;
        }

        return output.ToArray();
    }

    /// <summary>
    /// How a copy token splits into length and offset (MS-OVBA 2.4.1.3.19.1): the offset takes as
    /// many bits as the current position in the chunk needs, and the length takes the rest.
    /// </summary>
    private static (int LengthBits, int LengthMask, int OffsetShift) TokenShape(int decompressedInChunk)
    {
        // BitCount is Max(4, Ceiling(Log2(difference))): it grows only once the chunk is longer than a power of two.
        var bits = 4;
        for (var limit = 16; limit < decompressedInChunk && bits < 12; limit <<= 1)
        {
            bits++;
        }

        var lengthBits = 16 - bits;
        return (lengthBits, (1 << lengthBits) - 1, lengthBits);
    }
}
