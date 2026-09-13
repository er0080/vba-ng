using VbaNg.Runtime.TypeLibraries;

using Xunit;

namespace VbaNg.Interop.Tests;

/// <summary>
/// The type library reader against the Scripting runtime's library (scrrun.dll), registered on
/// every Windows: coclasses, dual interfaces read from their dispatch view, dispids, parameters,
/// enums, and the JSON cache round trip.
/// </summary>
public sealed class TypeLibraryReaderTests
{
    private static readonly Guid ScriptingGuid = new("420B2830-E718-11CF-893D-00A0C9054228");

    private static ComLibrary Scripting() => TypeLibraryReader.Read(ScriptingGuid, 1, 0);

    /// <summary>
    /// The Office library (mso.dll), a default reference of every Excel project, carries parameter
    /// defaults with no text form (a null IDispatch); reading it used to throw out of ReadConstant
    /// and take the whole build with it (ROADMAP.md, review of 2026-09-09). Skipped where Office is
    /// not registered, since the library is Office's own, not Windows's.
    /// </summary>
    [Fact]
    public void Read_OfficeLibrary_SurvivesObjectValuedDefaults()
    {
        var officeGuid = new Guid("2DF8D04C-5BFA-101B-BDE5-00AA0044DE52");
        Assert.SkipWhen(TypeLibraryCache.Resolve("Office", officeGuid.ToString("B"), "2.8") is null && !Registered(officeGuid, "2.8"), "The Office type library is not registered on this machine.");

        var library = TypeLibraryReader.Read(officeGuid, 2, 8);

        Assert.Equal("Office", library.Name);
        Assert.NotNull(library.FindType("IRibbonControl"));
        Assert.NotNull(library.FindType("MsoTriState"));
    }

    private static bool Registered(Guid libraryId, string version) =>
        Microsoft.Win32.Registry.ClassesRoot.OpenSubKey($@"TypeLib\{libraryId:B}\{version}") is not null;

    [Fact]
    public void Read_Scripting_FindsTheDictionaryCoClassAndItsInterface()
    {
        var library = Scripting();

        Assert.Equal("Scripting", library.Name);
        Assert.Equal(ScriptingGuid, library.Guid);
        Assert.Equal((1, 0), (library.MajorVersion, library.MinorVersion));

        var dictionary = library.FindType("Dictionary");
        Assert.NotNull(dictionary);
        Assert.Equal(ComTypeKind.CoClass, dictionary.Kind);
        Assert.Equal("Scripting.IDictionary", dictionary.DefaultInterface);

        var iDictionary = library.FindType("IDictionary");
        Assert.NotNull(iDictionary);
        Assert.True(iDictionary.Kind is ComTypeKind.Dispatch or ComTypeKind.Interface);
        Assert.True(iDictionary.IsDual);

        var add = Assert.Single(iDictionary.FindMembers("Add"));
        Assert.Equal(ComMemberKind.Method, add.Kind);
        Assert.Equal(["Key", "Item"], add.Parameters.Select(p => p.Name));
        Assert.All(add.Parameters, p => Assert.Equal((int)System.Runtime.InteropServices.VarEnum.VT_VARIANT, p.Type.VarType));
        Assert.All(add.Parameters, p => Assert.True(p.IsByRef));
        Assert.Null(add.Type);

        var count = Assert.Single(iDictionary.FindMembers("Count"));
        Assert.Equal(ComMemberKind.PropertyGet, count.Kind);
        Assert.Equal((int)System.Runtime.InteropServices.VarEnum.VT_I4, count.Type!.VarType);

        var item = iDictionary.FindMembers("Item").ToList();
        Assert.Contains(item, m => m.Kind == ComMemberKind.PropertyGet && m.IsDefault);
        Assert.Contains(item, m => m.Kind == ComMemberKind.PropertyPut);
        Assert.Contains(item, m => m.Kind == ComMemberKind.PropertyPutRef);

        var newEnum = Assert.Single(iDictionary.Members, m => m.IsNewEnum);
        Assert.True(newEnum.IsHidden || newEnum.IsRestricted);

        var keys = Assert.Single(iDictionary.FindMembers("Keys"));
        Assert.Equal((int)System.Runtime.InteropServices.VarEnum.VT_VARIANT, keys.Type!.VarType);
    }

    [Fact]
    public void Read_Scripting_ReadsEnumConstantsAndObjectReturns()
    {
        var library = Scripting();

        var ioMode = library.FindType("IOMode");
        Assert.NotNull(ioMode);
        Assert.Equal(ComTypeKind.Enum, ioMode.Kind);
        Assert.Equal("1", ioMode.FindMember("ForReading", ComMemberKind.Constant)!.Value!.Text);
        Assert.Equal("8", ioMode.FindMember("ForAppending", ComMemberKind.Constant)!.Value!.Text);

        var fso = library.FindType("IFileSystem3") ?? library.FindType("IFileSystem");
        Assert.NotNull(fso);
        var openTextFile = Assert.Single(fso.FindMembers("OpenTextFile"));
        Assert.Equal("Scripting.ITextStream", openTextFile.Type!.TypeName);
        Assert.True(openTextFile.Type.IsPointer);
        var mode = openTextFile.Parameters.Single(p => p.Name == "IOMode");
        Assert.True(mode.IsOptional);
        Assert.Equal("Scripting.IOMode", mode.Type.TypeName);
        Assert.Equal("1", mode.Default!.Text);
        var fileName = openTextFile.Parameters[0];
        Assert.Equal((int)System.Runtime.InteropServices.VarEnum.VT_BSTR, fileName.Type.VarType);
        Assert.False(fileName.IsByRef);
    }

    [Fact]
    public void Model_RoundTripsThroughJson()
    {
        var library = Scripting();

        var copy = ComLibrary.FromJson(library.ToJson());

        Assert.Equal(library.Types.Count, copy.Types.Count);
        Assert.Equal(library.ToJson(), copy.ToJson());
        Assert.EndsWith("420b2830-e718-11cf-893d-00a0c9054228-1.0.json", ComLibrary.CachePath(library.Guid, 1, 0), StringComparison.Ordinal);
    }
}
