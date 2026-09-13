namespace VbaNg.Runtime;

/// <summary>
/// A VBA Error value (MS-VBAL 2.1): an SCODE carried in a Variant, as <c>CVErr</c> produces and
/// as omitted optional arguments (Missing) are represented.
/// </summary>
public readonly struct ErrorValue : IEquatable<ErrorValue>
{
    /// <summary>The facility VBA stamps on <c>CVErr</c> codes: 0x800A0000 plus the number.</summary>
    private const int VbaFacility = unchecked((int)0x800A0000);

    /// <summary>DISP_E_PARAMNOTFOUND, the value of an omitted optional Variant argument.</summary>
    public static readonly ErrorValue Missing = new(unchecked((int)0x80020004));

    public ErrorValue(int scode) => Scode = scode;

    /// <summary>The raw SCODE.</summary>
    public int Scode { get; }

    /// <summary>
    /// The error number the value shows to VBA code: the low 16 bits for codes in VBA's facility
    /// (<c>CVErr(5)</c> is 0x800A0005 and reads back as 5), otherwise the SCODE itself.
    /// </summary>
    public int Number => (Scode & unchecked((int)0xFFFF0000)) == VbaFacility ? Scode & 0xFFFF : Scode;

    public bool IsMissing => Scode == Missing.Scode;

    /// <summary><c>CVErr(n)</c> for 0 to 65535 (MS-VBAL 6.1.2.3 Conversion, CVErr); other numbers raise error 5.</summary>
    public static ErrorValue FromNumber(long number)
    {
        if (number is < 0 or > 65535)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        return new ErrorValue(VbaFacility | (int)number);
    }

    public bool Equals(ErrorValue other) => Scode == other.Scode;

    public override bool Equals(object? obj) => obj is ErrorValue other && Equals(other);

    public override int GetHashCode() => Scode;

    public override string ToString() => "Error " + Number.ToString(System.Globalization.CultureInfo.InvariantCulture);

    public static bool operator ==(ErrorValue left, ErrorValue right) => left.Equals(right);

    public static bool operator !=(ErrorValue left, ErrorValue right) => !left.Equals(right);
}
