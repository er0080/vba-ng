namespace VbaNg.Golden.Replay;

/// <summary>
/// Whether this machine's C runtime rounds Sin and Cos as the one that recorded the goldens did.
/// vba-ng takes their last bit from the Windows C runtime, so a golden recorded on one Windows can
/// differ in that bit on another. Probed on Cos(1).
/// </summary>
public static class HostMath
{
    // Cos(1) as VBA printed it on the recording host: one ulp below the correctly rounded value,
    // 0x3FE14A280FB5068C, which a newer C runtime returns instead.
    private const long RecordedCosOfOne = 0x3FE14A280FB5068B;

    public static bool RoundsAsRecorded { get; } = BitConverter.DoubleToInt64Bits(Math.Cos(1.0)) == RecordedCosOfOne;
}
