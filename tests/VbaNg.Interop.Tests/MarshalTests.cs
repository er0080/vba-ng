using System.Runtime.InteropServices;

using VbaNg.Runtime;

using Xunit;

namespace VbaNg.Interop.Tests;

/// <summary>
/// The invoke path hands COM a Variant as the VARIANT it already is (ARCHITECTURE.md section 5;
/// ROADMAP.md M7 F): a String crosses byte for byte, an odd trailing byte and a null BSTR
/// included, and an array crosses as oleaut32's copy of its SAFEARRAY, which the callee owns.
/// </summary>
public sealed unsafe class MarshalTests
{
    [Fact]
    public void String_CrossesByteForByte_AndANullStaysNull()
    {
        var mark = ObjectRefs.Mark();
        try
        {
            var odd = Variant.FromVbaString(VbaString.AllocBytes([0x41, 0x00, 0x42]));
            var native = default(VARIANT);
            VariantMarshal.ToNative(odd, &native);
            Assert.Equal(Vt.Bstr, native.vt);
            Assert.Equal(3u, NativeMethods.SysStringByteLen(native.pointer));
            Assert.NotEqual(odd.AsVbaString().Pointer, native.pointer);

            var back = VariantMarshal.FromNative(&native);
            Assert.Equal(new byte[] { 0x41, 0x00, 0x42 }, back.AsVbaString().Bytes.ToArray());
            VariantMarshal.Clear(&native);

            VariantMarshal.ToNative(Variant.EmptyString, &native);
            Assert.Equal(0, native.pointer);
            var none = VariantMarshal.FromNative(&native);
            Assert.True(none.IsString);
            Assert.True(none.AsVbaString().IsNull);
        }
        finally
        {
            ObjectRefs.ReleaseTo(mark);
        }
    }

    [Fact]
    public void FixedArray_CrossesAsOleautsCopy_WithoutTheFixedSizeFlags()
    {
        var mark = ObjectRefs.Mark();
        var array = VbaArray.Create(VarType.String, [(1, 2)], fixedSize: true);
        try
        {
            array.Set([1], Variant.FromString("a"));
            array.Set([2], Variant.FromString("bc"));
            var copy = SafeArrayMarshal.ToNative(array);
            try
            {
                Assert.NotEqual(array.Descriptor, copy);
                Assert.Equal(0, Marshal.ReadInt16(copy, 2) & 0x12);
                var back = SafeArrayMarshal.FromNative(copy, Vt.Bstr);
                Assert.Equal((1, 2), (back.LBound(1), back.UBound(1)));
                Assert.Equal("bc", back.Get([2]).AsString());
                ObjectRefs.Release(ref back);
            }
            finally
            {
                _ = NativeMethods.SafeArrayDestroy(copy);
            }
        }
        finally
        {
            ObjectRefs.Release(ref array);
            ObjectRefs.ReleaseTo(mark);
        }
    }
}
