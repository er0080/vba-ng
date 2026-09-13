using System.Globalization;

using VbaNg.Compiler.Binding;

namespace VbaNg.Compiler.Emit;

/// <summary>The file statements (MS-VBAL 5.4.5) as calls into the runtime's FileSystem module.</summary>
public sealed partial class CSharpEmitter
{
    private const string F = L + "FileSystem.";

    private string FileStatement(BoundStatement statement)
    {
        switch (statement)
        {
            case BoundOpen open:
                {
                    var length = open.RecordLength is null ? VariantType + ".Missing" : Variant(open.RecordLength);
                    return $"{F}Open({Variant(open.PathName)}, {L}OpenMode.{open.Mode}, {L}FileAccessMode.{open.Access}, {L}FileLockMode.{open.LockMode}, {Variant(open.FileNumber)}, {length});";
                }

            case BoundFileStatement { Kind: FileStatementKind.Close } close:
                return $"{F}Close([{string.Join(", ", close.Operands.Select(o => Variant(o!)))}]);";
            case BoundFileStatement { Kind: FileStatementKind.Reset }:
                return F + "Reset();";
            case BoundFileStatement { Kind: FileStatementKind.Seek } seek:
                return $"{F}Seek({Variant(seek.FileNumber!)}, {Variant(seek.Operands[0]!)});";
            case BoundFileStatement { Kind: FileStatementKind.Lock or FileStatementKind.Unlock } lockStatement:
                return $"{F}Lock({Variant(lockStatement.FileNumber!)}, {Variant(lockStatement.Operands[0]!)}, {Variant(lockStatement.Operands[1]!)}, {(lockStatement.Kind == FileStatementKind.Unlock ? "true" : "false")});";
            case BoundFileStatement { Kind: FileStatementKind.Width } width:
                return $"{F}Width({Variant(width.FileNumber!)}, {Variant(width.Operands[0]!)});";
            case BoundFileStatement { Kind: FileStatementKind.Name } name:
                return $"{F}Rename({Variant(name.Operands[0]!)}, {Variant(name.Operands[1]!)});";
            case BoundFileOutput { IsWrite: false } print:
                return $"{F}Print({Variant(print.FileNumber)}, {PrintListCode(print.Items)}, {Bool(EndsLine(print.Items))});";
            case BoundFileOutput write:
                {
                    var items = write.Items.Where(i => i.Kind == PrintItemKind.Value && i.Value is not null).Select(i => Variant(i.Value!));
                    return $"{F}Write({Variant(write.FileNumber)}, [{string.Join(", ", items)}], {Bool(EndsLine(write.Items))});";
                }

            case BoundFileInput input:
                {
                    var stores = input.Targets.Select(target => input.IsLineInput
                        ? Store(target, new Emitted($"{F}LineInput({Variant(input.FileNumber)})", VbaType.String))
                        : Store(target, new Emitted($"{F}Input({Variant(input.FileNumber)}, {VarTypeName(target.Type)})", VbaType.Variant)));
                    return string.Join("; ", stores) + ";";
                }

            case BoundFileRecord record:
                return FileRecord(record);
            default:
                Report(statement.Syntax, $"{statement.GetType().Name} is not a file statement; this is a compiler bug.");
                return string.Empty;
        }
    }

    /// <summary>Get and Put: a record or array is read and written in place; a scalar or string goes through its VarType, with the fixed length of a String * n.</summary>
    private string FileRecord(BoundFileRecord record)
    {
        var fileNumber = Variant(record.FileNumber);
        var recordNumber = Variant(record.RecordNumber!);
        var variable = record.Variable;
        if (variable.Type.IsRecord)
        {
            var target = EmitExpression(variable).Code;
            return record.IsPut ? $"{F}PutRecord({fileNumber}, {recordNumber}, {target});" : $"{F}GetRecord({fileNumber}, {recordNumber}, ref {target});";
        }

        if (variable.Type.IsArray)
        {
            return $"{F}{(record.IsPut ? "PutArray" : "GetArray")}({fileNumber}, {recordNumber}, {EmitExpression(variable).Code});";
        }

        var fixedLength = variable.Type.Kind == TypeKind.FixedString ? variable.Type.FixedLength : 0;
        var layout = $"{VarTypeName(variable.Type)}, {fixedLength.ToString(CultureInfo.InvariantCulture)}";
        if (record.IsPut)
        {
            return $"{F}Put({fileNumber}, {recordNumber}, {Variant(variable)}, {layout});";
        }

        return Store(variable, new Emitted($"{F}Get({fileNumber}, {recordNumber}, {Variant(variable)}, {layout})", VbaType.Variant)) + ";";
    }

    /// <summary>The PrintList expression of an output list (MS-VBAL 5.4.5.6 output-list).</summary>
    private string PrintListCode(IReadOnlyList<BoundPrintItem> items)
    {
        var list = $"new {R}PrintList()";
        foreach (var item in items)
        {
            switch (item.Kind)
            {
                case PrintItemKind.Value when item.Value is not null:
                    list += $".Item({Variant(item.Value)})";
                    break;
                case PrintItemKind.Spc:
                    list += $".Spc({Convert(EmitExpression(item.Value!), VbaType.Long)})";
                    break;
                case PrintItemKind.Tab:
                    list += item.Value is null ? ".Tab()" : $".Tab({Convert(EmitExpression(item.Value), VbaType.Long)})";
                    break;
            }

            if (item.Separator == ',')
            {
                list += ".Zone()";
            }
        }

        return list;
    }

    /// <summary>The line terminator is written only when the list does not end in a separator (MS-VBAL 5.4.5.6).</summary>
    private static bool EndsLine(IReadOnlyList<BoundPrintItem> items) => items.Count == 0 || items[^1].Separator is null;

    private string Variant(BoundExpression expression) => Convert(EmitExpression(expression), VbaType.Variant);

    private static string Bool(bool value) => value ? "true" : "false";
}
