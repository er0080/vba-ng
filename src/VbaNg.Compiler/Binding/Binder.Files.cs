using VbaNg.Compiler.Syntax;
using VbaNg.Runtime;

namespace VbaNg.Compiler.Binding;

/// <summary>The file statements (MS-VBAL 5.4.5), bound to the FileSystem module of the runtime.</summary>
public sealed partial class Binder
{
    private BoundStatement BindFileStatement(StatementSyntax statement)
    {
        switch (statement)
        {
            case OpenStatementSyntax open:
                return BindOpen(open);
            case CloseStatementSyntax close:
                return new BoundFileStatement(close, FileStatementKind.Close, null, close.FileNumbers.Select(n => (BoundExpression?)BindFileNumber(n)).ToList());
            case SeekStatementSyntax seek:
                return new BoundFileStatement(seek, FileStatementKind.Seek, BindFileNumber(seek.FileNumber), [Convert(BindExpression(seek.Position), VbaType.Variant)]);
            case LockStatementSyntax lockStatement:
                return new BoundFileStatement(
                    lockStatement,
                    lockStatement.IsUnlock ? FileStatementKind.Unlock : FileStatementKind.Lock,
                    BindFileNumber(lockStatement.FileNumber),
                    [Optional(lockStatement.Range?.FromRecord, lockStatement), Optional(lockStatement.Range?.ToRecord, lockStatement)]);
            case WidthStatementSyntax width:
                return new BoundFileStatement(width, FileStatementKind.Width, BindFileNumber(width.FileNumber), [Convert(BindExpression(width.Width), VbaType.Variant)]);
            case NameStatementSyntax name:
                return new BoundFileStatement(name, FileStatementKind.Name, null, [Convert(BindExpression(name.OldName), VbaType.Variant), Convert(BindExpression(name.NewName), VbaType.Variant)]);
            case FileOutputStatementSyntax output:
                return new BoundFileOutput(output, output.Kind == SyntaxKind.WriteStatement, BindFileNumber(output.FileNumber), BindOutputItems(output.Items));
            case LineInputStatementSyntax lineInput:
                return new BoundFileInput(lineInput, isLineInput: true, BindFileNumber(lineInput.FileNumber), [BindTarget(lineInput.Variable, isSet: false)]);
            case InputStatementSyntax input:
                return new BoundFileInput(input, isLineInput: false, BindFileNumber(input.FileNumber), input.Variables.Select(v => BindTarget(v, isSet: false)).ToList());
            case PutGetStatementSyntax putGet:
                return BindPutGet(putGet);
            default:
                // Reset (MS-VBAL 5.4.5.1.2).
                return new BoundFileStatement(statement, FileStatementKind.Reset, null, []);
        }
    }

    /// <summary>Open pathname For mode [Access access] [lock] As #filenumber [Len = reclength] (MS-VBAL 5.4.5.1); no mode means Random.</summary>
    private BoundOpen BindOpen(OpenStatementSyntax open)
    {
        var mode = open.Mode is null ? "Random" : Capitalized(open.Mode.Text);
        var access = string.Concat(open.AccessModes.Select(t => Capitalized(t.Text)));
        var lockMode = string.Concat(open.LockModes.Select(t => Capitalized(t.Text)));
        return new BoundOpen(
            open,
            Convert(BindExpression(open.PathName), VbaType.Variant),
            mode,
            access.Length == 0 ? "Default" : access,
            lockMode.Length == 0 ? "Default" : lockMode,
            BindFileNumber(open.FileNumber),
            open.RecordLength is null ? null : Convert(BindExpression(open.RecordLength), VbaType.Variant));
    }

    /// <summary>Put #file, [record], variable or Get #file, [record], variable (MS-VBAL 5.4.5.9): Get needs an assignable variable.</summary>
    private BoundFileRecord BindPutGet(PutGetStatementSyntax putGet)
    {
        var isPut = putGet.Kind == SyntaxKind.PutStatement;
        var variable = isPut ? BindExpression(putGet.Variable) : BindTarget(putGet.Variable, isSet: false);
        return new BoundFileRecord(
            putGet,
            isPut,
            BindFileNumber(putGet.FileNumber),
            Optional(putGet.RecordNumber, putGet),
            variable);
    }

    private BoundExpression BindFileNumber(FileNumberSyntax fileNumber) => Convert(BindExpression(fileNumber.Number), VbaType.Variant);

    /// <summary>An operand the statement may leave out; Missing stands in for it.</summary>
    private BoundExpression Optional(ExpressionSyntax? expression, SyntaxNode at) =>
        expression is null ? new BoundLiteral(at, Variant.Missing, VbaType.Variant) : Convert(BindExpression(expression), VbaType.Variant);

    /// <summary>The output list of Print #, Write #, and Debug.Print (MS-VBAL 5.4.5.6 output-list).</summary>
    private List<BoundPrintItem> BindOutputItems(SyntaxList<OutputItemSyntax> items)
    {
        var bound = new List<BoundPrintItem>();
        foreach (var item in items)
        {
            var separator = item.Separator?.Kind switch
            {
                SyntaxKind.SemicolonToken => ';',
                SyntaxKind.CommaToken => ',',
                _ => (char?)null,
            };
            switch (item.Clause)
            {
                case null:
                    bound.Add(new BoundPrintItem(PrintItemKind.Value, null, separator));
                    break;
                case SpcClauseSyntax spc:
                    bound.Add(new BoundPrintItem(PrintItemKind.Spc, Convert(BindExpression(spc.Count), VbaType.Long), separator));
                    break;
                case TabClauseSyntax tab:
                    bound.Add(new BoundPrintItem(PrintItemKind.Tab, tab.Column is null ? null : Convert(BindExpression(tab.Column), VbaType.Long), separator));
                    break;
                case ExpressionSyntax expression:
                    bound.Add(new BoundPrintItem(PrintItemKind.Value, Convert(BindExpression(expression), VbaType.Variant), separator));
                    break;
            }
        }

        return bound;
    }

    private static string Capitalized(string word) => word.Length == 0 ? word : char.ToUpperInvariant(word[0]) + word[1..].ToLowerInvariant();
}
