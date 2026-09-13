using System.Globalization;

using Microsoft.Win32;

namespace VbaNg.Runtime.Library;

/// <summary>The directory and attribute functions of the FileSystem module, and the registry settings of Interaction, with VBA's error numbers (FileSystem golden).</summary>
public static partial class FileSystem
{
    private const int AttrNormal = 0;
    private const int AttrReadOnly = 1;
    private const int AttrHidden = 2;
    private const int AttrSystem = 4;
    private const int AttrDirectory = 16;
    private const int AttrArchive = 32;

    [ThreadStatic]
    private static IEnumerator<string>? dirWalk;

    /// <summary>CurDir([drive]): the current directory, of the given drive when one is named.</summary>
    public static VbaString CurDir(in Variant drive) => VbaString.Temporary(CurDirText(drive));

    private static string CurDirText(in Variant drive)
    {
        if (drive.IsMissing)
        {
            return Environment.CurrentDirectory;
        }

        var letter = Coerce.ToString(drive);
        if (letter.Length == 0)
        {
            return Environment.CurrentDirectory;
        }

        var current = Environment.CurrentDirectory;
        if (current.Length >= 2 && char.ToUpperInvariant(current[0]) == char.ToUpperInvariant(letter[0]))
        {
            return current;
        }

        try
        {
            return Path.GetFullPath(letter[0] + ":");
        }
        catch (ArgumentException)
        {
            throw new VbaException(68);
        }
        catch (NotSupportedException)
        {
            throw new VbaException(68);
        }
    }

    public static void ChDir(in Variant path)
    {
        var target = Coerce.ToString(path);
        if (target.Length == 0)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        if (!Directory.Exists(target))
        {
            throw new VbaException(76);
        }

        Environment.CurrentDirectory = target;
    }

    /// <summary>ChDrive drive: an empty string does nothing; a drive that does not exist raises 68.</summary>
    public static void ChDrive(in Variant drive)
    {
        var letter = Coerce.ToString(drive);
        if (letter.Length == 0)
        {
            return;
        }

        var root = char.ToUpperInvariant(letter[0]) + ":\\";
        if (!char.IsAsciiLetter(letter[0]) || !Directory.Exists(root))
        {
            throw new VbaException(68);
        }

        Environment.CurrentDirectory = Path.GetFullPath(root[..2]);
    }

    public static void MkDir(in Variant path)
    {
        var target = Coerce.ToString(path);
        if (target.Length == 0)
        {
            throw new VbaException(76);
        }

        if (Directory.Exists(target) || File.Exists(target))
        {
            throw new VbaException(75);
        }

        try
        {
            Directory.CreateDirectory(target);
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
            throw new VbaException(75);
        }
    }

    public static void RmDir(in Variant path)
    {
        var target = Coerce.ToString(path);
        if (!Directory.Exists(target))
        {
            throw new VbaException(76);
        }

        try
        {
            Directory.Delete(target, recursive: false);
        }
        catch (IOException)
        {
            throw new VbaException(75);
        }
        catch (UnauthorizedAccessException)
        {
            throw new VbaException(75);
        }
    }

    /// <summary>Kill pathname: deletes the files matching the name or wildcard pattern; none raises 53.</summary>
    public static void Kill(in Variant pathName)
    {
        var pattern = Coerce.ToString(pathName);
        var matches = Expand(pattern, includeDirectories: false).ToList();
        if (matches.Count == 0)
        {
            throw new VbaException(53);
        }

        foreach (var file in matches)
        {
            try
            {
                File.Delete(file);
            }
            catch (UnauthorizedAccessException)
            {
                throw new VbaException(70);
            }
            catch (IOException)
            {
                throw new VbaException(70);
            }
        }
    }

    public static void FileCopy(in Variant source, in Variant destination)
    {
        var from = Coerce.ToString(source);
        var to = Coerce.ToString(destination);
        if (!File.Exists(from))
        {
            throw new VbaException(53);
        }

        try
        {
            File.Copy(from, to, overwrite: true);
        }
        catch (DirectoryNotFoundException)
        {
            throw new VbaException(76);
        }
        catch (UnauthorizedAccessException)
        {
            throw new VbaException(70);
        }
        catch (IOException)
        {
            throw new VbaException(70);
        }
    }

    /// <summary>Name oldpathname As newpathname: renames or moves a file; an existing target raises 58, a missing source 53.</summary>
    public static void Rename(in Variant oldName, in Variant newName)
    {
        var from = Coerce.ToString(oldName);
        var to = Coerce.ToString(newName);
        if (File.Exists(to) || Directory.Exists(to))
        {
            throw new VbaException(58);
        }

        if (!File.Exists(from) && !Directory.Exists(from))
        {
            throw new VbaException(53);
        }

        try
        {
            if (Directory.Exists(from))
            {
                Directory.Move(from, to);
            }
            else
            {
                File.Move(from, to);
            }
        }
        catch (DirectoryNotFoundException)
        {
            throw new VbaException(76);
        }
        catch (UnauthorizedAccessException)
        {
            throw new VbaException(70);
        }
        catch (IOException)
        {
            throw new VbaException(70);
        }
    }

    public static int FileLen(in Variant pathName)
    {
        var path = Coerce.ToString(pathName);
        if (!File.Exists(path))
        {
            throw new VbaException(53);
        }

        return checked((int)new FileInfo(path).Length);
    }

    public static VbaDate FileDateTime(in Variant pathName)
    {
        var path = Coerce.ToString(pathName);
        if (File.Exists(path))
        {
            return VbaDate.FromDateTime(File.GetLastWriteTime(path));
        }

        if (Directory.Exists(path))
        {
            return VbaDate.FromDateTime(Directory.GetLastWriteTime(path));
        }

        throw new VbaException(53);
    }

    /// <summary>GetAttr(pathname): the VbFileAttribute bits.</summary>
    public static int GetAttr(in Variant pathName)
    {
        var path = Coerce.ToString(pathName);
        FileAttributes attributes;
        try
        {
            attributes = File.GetAttributes(path);
        }
        catch (FileNotFoundException)
        {
            throw new VbaException(53);
        }
        catch (DirectoryNotFoundException)
        {
            throw new VbaException(53);
        }

        return ToVba(attributes);
    }

    public static void SetAttr(in Variant pathName, in Variant attributes)
    {
        var path = Coerce.ToString(pathName);
        var bits = Coerce.ToInt32(attributes);
        if (!File.Exists(path) && !Directory.Exists(path))
        {
            throw new VbaException(53);
        }

        if ((bits & AttrDirectory) != 0)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        var current = File.GetAttributes(path) & ~(FileAttributes.ReadOnly | FileAttributes.Hidden | FileAttributes.System | FileAttributes.Archive);
        if ((bits & AttrReadOnly) != 0)
        {
            current |= FileAttributes.ReadOnly;
        }

        if ((bits & AttrHidden) != 0)
        {
            current |= FileAttributes.Hidden;
        }

        if ((bits & AttrSystem) != 0)
        {
            current |= FileAttributes.System;
        }

        if ((bits & AttrArchive) != 0)
        {
            current |= FileAttributes.Archive;
        }

        File.SetAttributes(path, current == 0 ? FileAttributes.Normal : current);
    }

    /// <summary>
    /// Dir([pathname], [attributes]): the first name matching the pattern, then, called without a
    /// pattern, the next one, and an empty string at the end; vbDirectory includes folders. Calling
    /// it again after the end, or before any pattern, raises 5.
    /// </summary>
    public static VbaString Dir(in Variant pathName, in Variant attributes) => VbaString.Temporary(DirText(pathName, attributes));

    private static string DirText(in Variant pathName, in Variant attributes)
    {
        if (pathName.IsMissing)
        {
            if (dirWalk is null)
            {
                throw VbaErrors.InvalidProcedureCall();
            }

            if (dirWalk.MoveNext())
            {
                return dirWalk.Current;
            }

            dirWalk = null;
            return string.Empty;
        }

        var pattern = Coerce.ToString(pathName);
        var bits = attributes.IsMissing ? AttrNormal : Coerce.ToInt32(attributes);
        dirWalk = Walk(pattern, bits).GetEnumerator();
        if (dirWalk.MoveNext())
        {
            return dirWalk.Current;
        }

        dirWalk = null;
        return string.Empty;
    }

    private static IEnumerable<string> Walk(string pattern, int attributes)
    {
        var includeDirectories = (attributes & AttrDirectory) != 0;
        var includeHidden = (attributes & AttrHidden) != 0;
        var includeSystem = (attributes & AttrSystem) != 0;
        if (pattern.Length == 0)
        {
            pattern = "*";
        }

        // A folder named without vbDirectory matches nothing; with it, the folder itself, as any other entry (golden: Dir walks a wildcard pattern).
        if (pattern.EndsWith(Path.DirectorySeparatorChar) || pattern.EndsWith(Path.AltDirectorySeparatorChar))
        {
            pattern += "*";
        }

        var directory = Path.GetDirectoryName(pattern);
        if (string.IsNullOrEmpty(directory))
        {
            directory = Environment.CurrentDirectory;
        }

        var filePattern = Path.GetFileName(pattern);
        if (!Directory.Exists(directory))
        {
            yield break;
        }

        var entries = new List<string>();
        if (includeDirectories && (filePattern == "*" || filePattern == "*.*"))
        {
            entries.Add(".");
            entries.Add("..");
        }

        var options = new EnumerationOptions { MatchType = MatchType.Win32, AttributesToSkip = 0, IgnoreInaccessible = true, MatchCasing = MatchCasing.CaseInsensitive };
        foreach (var entry in Directory.EnumerateFileSystemEntries(directory, filePattern, options).OrderBy(e => Path.GetFileName(e), StringComparer.OrdinalIgnoreCase))
        {
            var info = File.GetAttributes(entry);
            if ((info & FileAttributes.Directory) != 0 && !includeDirectories)
            {
                continue;
            }

            if ((info & FileAttributes.Hidden) != 0 && !includeHidden)
            {
                continue;
            }

            if ((info & FileAttributes.System) != 0 && !includeSystem)
            {
                continue;
            }

            entries.Add(Path.GetFileName(entry));
        }

        foreach (var entry in entries)
        {
            yield return entry;
        }
    }

    private static IEnumerable<string> Expand(string pattern, bool includeDirectories)
    {
        if (pattern.Length == 0)
        {
            yield break;
        }

        if (pattern.IndexOfAny(['*', '?']) < 0)
        {
            if (File.Exists(pattern))
            {
                yield return pattern;
            }

            yield break;
        }

        var directory = Path.GetDirectoryName(pattern);
        if (string.IsNullOrEmpty(directory))
        {
            directory = Environment.CurrentDirectory;
        }

        if (!Directory.Exists(directory))
        {
            yield break;
        }

        var options = new EnumerationOptions { MatchType = MatchType.Win32, AttributesToSkip = 0, IgnoreInaccessible = true, MatchCasing = MatchCasing.CaseInsensitive };
        foreach (var entry in Directory.EnumerateFiles(directory, Path.GetFileName(pattern), options))
        {
            yield return entry;
        }

        if (includeDirectories)
        {
            foreach (var entry in Directory.EnumerateDirectories(directory, Path.GetFileName(pattern), options))
            {
                yield return entry;
            }
        }
    }

    private static int ToVba(FileAttributes attributes)
    {
        var bits = 0;
        if ((attributes & FileAttributes.ReadOnly) != 0)
        {
            bits |= AttrReadOnly;
        }

        if ((attributes & FileAttributes.Hidden) != 0)
        {
            bits |= AttrHidden;
        }

        if ((attributes & FileAttributes.System) != 0)
        {
            bits |= AttrSystem;
        }

        if ((attributes & FileAttributes.Directory) != 0)
        {
            bits |= AttrDirectory;
        }

        if ((attributes & FileAttributes.Archive) != 0)
        {
            bits |= AttrArchive;
        }

        return bits;
    }

    // Registry settings (MS-VBAL 6.1.2.6 SaveSetting, GetSetting, GetAllSettings, DeleteSetting).

    private const string SettingsRoot = @"Software\VB and VBA Program Settings";

    public static void SaveSetting(in Variant appName, in Variant section, in Variant key, in Variant setting)
    {
        if (!OperatingSystem.IsWindows())
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        var path = SettingsPath(appName, section);
        using var registryKey = Registry.CurrentUser.CreateSubKey(path) ?? throw VbaErrors.InvalidProcedureCall();
        registryKey.SetValue(Coerce.ToString(key), Coerce.ToString(setting));
    }

    public static VbaString GetSetting(in Variant appName, in Variant section, in Variant key, in Variant defaultValue) => VbaString.Temporary(GetSettingText(appName, section, key, defaultValue));

    private static string GetSettingText(in Variant appName, in Variant section, in Variant key, in Variant defaultValue)
    {
        if (!OperatingSystem.IsWindows())
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        var fallback = defaultValue.IsMissing ? string.Empty : Coerce.ToString(defaultValue);
        using var registryKey = Registry.CurrentUser.OpenSubKey(SettingsPath(appName, section));
        return registryKey?.GetValue(Coerce.ToString(key)) as string ?? fallback;
    }

    /// <summary>GetAllSettings(appname, section): a two-dimensional array of keys and values, or Empty when the section does not exist.</summary>
    public static Variant GetAllSettings(in Variant appName, in Variant section)
    {
        if (!OperatingSystem.IsWindows())
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        using var registryKey = Registry.CurrentUser.OpenSubKey(SettingsPath(appName, section));
        if (registryKey is null)
        {
            return Variant.Empty;
        }

        var names = registryKey.GetValueNames();
        if (names.Length == 0)
        {
            return Variant.Empty;
        }

        var array = VbaArray.Create(VarType.Variant, [(0, names.Length - 1), (0, 1)]);
        for (var i = 0; i < names.Length; i++)
        {
            array.Set([i, 0], Variant.FromString(names[i]));
            array.Set([i, 1], Variant.FromString(registryKey.GetValue(names[i]) as string ?? string.Empty));
        }

        return Variant.FromArray(ObjectRefs.Owned(array));
    }

    /// <summary>DeleteSetting appname, [section], [key]: removes the key, the section, or the application's whole tree.</summary>
    public static void DeleteSetting(in Variant appName, in Variant section, in Variant key)
    {
        if (!OperatingSystem.IsWindows())
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        var application = Coerce.ToString(appName);
        if (section.IsMissing)
        {
            using var root = Registry.CurrentUser.OpenSubKey(SettingsRoot, writable: true);
            if (root?.OpenSubKey(application) is null)
            {
                throw VbaErrors.InvalidProcedureCall();
            }

            root.DeleteSubKeyTree(application, throwOnMissingSubKey: false);
            return;
        }

        var path = SettingsPath(appName, section);
        if (key.IsMissing)
        {
            using var applicationKey = Registry.CurrentUser.OpenSubKey(SettingsRoot + "\\" + application, writable: true);
            if (applicationKey?.OpenSubKey(Coerce.ToString(section)) is null)
            {
                throw VbaErrors.InvalidProcedureCall();
            }

            applicationKey.DeleteSubKeyTree(Coerce.ToString(section), throwOnMissingSubKey: false);
            return;
        }

        using var sectionKey = Registry.CurrentUser.OpenSubKey(path, writable: true);
        if (sectionKey is null || sectionKey.GetValue(Coerce.ToString(key)) is null)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        sectionKey.DeleteValue(Coerce.ToString(key), throwOnMissingValue: false);
    }

    private static string SettingsPath(in Variant appName, in Variant section) =>
        string.Create(CultureInfo.InvariantCulture, $"{SettingsRoot}\\{Coerce.ToString(appName)}\\{Coerce.ToString(section)}");
}
