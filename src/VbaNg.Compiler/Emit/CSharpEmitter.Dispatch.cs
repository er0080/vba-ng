using System.Globalization;

using VbaNg.Compiler.Binding;
using VbaNg.Compiler.Syntax;

namespace VbaNg.Compiler.Emit;

/// <summary>
/// The dispatch loop (ARCHITECTURE.md section 4, "Error handling"): a procedure that uses
/// On Error, Resume, GoTo, GoSub, or labels is flattened into a list of simple statements and
/// jumps, then emitted as a switch over a program counter inside while (true) and try/catch.
/// Resume, Resume Next, Resume label, Erl, and GoSub/Return then need no special control flow.
/// </summary>
public sealed partial class CSharpEmitter
{
    private abstract record Flat(SyntaxNode Syntax);

    /// <summary>A simple statement; <see cref="ResumeNextId"/> names where Resume Next continues when the statement fails, if not the next one (the Select Case expression, Procedures golden).</summary>
    private sealed record FlatSimple(SyntaxNode Syntax, BoundStatement Statement, int? ResumeNextId = null) : Flat(Syntax);

    private sealed record FlatLabel(SyntaxNode Syntax, int Id) : Flat(Syntax);

    private sealed record FlatJump(SyntaxNode Syntax, int Id) : Flat(Syntax);

    private sealed record FlatJumpIf(SyntaxNode Syntax, BoundExpression Condition, bool WhenTrue, int Id) : Flat(Syntax);

    private sealed record FlatForTest(SyntaxNode Syntax, BoundFor Loop, int EndId) : Flat(Syntax);

    private sealed record FlatForNext(SyntaxNode Syntax, BoundFor Loop, int TestId) : Flat(Syntax);

    private sealed record FlatForEachInit(SyntaxNode Syntax, BoundForEach Loop) : Flat(Syntax);

    private sealed record FlatForEachNext(SyntaxNode Syntax, BoundForEach Loop, int DoneId) : Flat(Syntax);

    private sealed record FlatForEachDone(SyntaxNode Syntax, BoundForEach Loop) : Flat(Syntax);
    private sealed record FlatForEachEnd(SyntaxNode Syntax, BoundForEach Loop) : Flat(Syntax);
    private sealed record FlatWith(SyntaxNode Syntax, BoundWith With) : Flat(Syntax);

    private sealed record FlatCaseTest(SyntaxNode Syntax, BoundSelect Select, BoundCaseBlock Block, int NextId) : Flat(Syntax);

    /// <summary>Flattens structured statements into simple statements, labels, and jumps (ARCHITECTURE.md section 4, step 4).</summary>
    private sealed class Flattener
    {
        private readonly List<Flat> flats = [];
        private readonly Stack<(ExitKind Kind, int EndId)> loops = new();
        private int nextId;

        public IReadOnlyList<Flat> Flats => flats;

        public void Block(BoundBlock block)
        {
            foreach (var statement in block.Statements)
            {
                Statement(statement);
            }
        }

        private int Label() => nextId++;

        private void Statement(BoundStatement statement)
        {
            switch (statement)
            {
                case BoundNop or BoundLocalDeclaration:
                    return;
                case BoundBlock block:
                    Block(block);
                    return;
                case BoundIf ifStatement:
                    {
                        var end = Label();
                        foreach (var branch in ifStatement.Branches)
                        {
                            var next = Label();
                            flats.Add(new FlatJumpIf(branch.Body.Syntax, branch.Condition, WhenTrue: false, next));
                            Block(branch.Body);
                            flats.Add(new FlatJump(branch.Body.Syntax, end));
                            flats.Add(new FlatLabel(branch.Body.Syntax, next));
                        }

                        if (ifStatement.ElseBody is not null)
                        {
                            Block(ifStatement.ElseBody);
                        }

                        flats.Add(new FlatLabel(ifStatement.Syntax, end));
                        return;
                    }

                case BoundSelect select:
                    {
                        var end = Label();
                        flats.Add(new FlatSimple(select.Syntax, new BoundAssignment(select.Syntax, new BoundVariable(select.Syntax, select.Selector), select.Value, isSet: false), end));
                        foreach (var block in select.Cases)
                        {
                            var next = Label();
                            flats.Add(new FlatCaseTest(block.Body.Syntax, select, block, next));
                            Block(block.Body);
                            flats.Add(new FlatJump(block.Body.Syntax, end));
                            flats.Add(new FlatLabel(block.Body.Syntax, next));
                        }

                        if (select.ElseBody is not null)
                        {
                            Block(select.ElseBody);
                        }

                        flats.Add(new FlatLabel(select.Syntax, end));
                        return;
                    }

                case BoundFor forLoop:
                    {
                        var test = Label();
                        var end = Label();
                        var limit = new BoundVariable(forLoop.Syntax, forLoop.Limit);
                        var increment = new BoundVariable(forLoop.Syntax, forLoop.Increment);
                        var step = forLoop.Step ?? new BoundLiteral(forLoop.Syntax, Runtime.Variant.FromInt16(1), VbaType.Integer);
                        flats.Add(new FlatSimple(forLoop.Syntax, new BoundAssignment(forLoop.Syntax, limit, forLoop.To, isSet: false)));
                        flats.Add(new FlatSimple(forLoop.Syntax, new BoundAssignment(forLoop.Syntax, increment, step, isSet: false)));
                        flats.Add(new FlatSimple(forLoop.Syntax, new BoundAssignment(forLoop.Syntax, forLoop.Counter, forLoop.From, isSet: false)));
                        flats.Add(new FlatLabel(forLoop.Syntax, test));
                        flats.Add(new FlatForTest(forLoop.Syntax, forLoop, end));
                        loops.Push((ExitKind.For, end));
                        Block(forLoop.Body);
                        loops.Pop();
                        flats.Add(new FlatForNext(forLoop.NextSyntax, forLoop, test));
                        flats.Add(new FlatLabel(forLoop.Syntax, end));
                        return;
                    }

                case BoundForEach forEach:
                    {
                        var test = Label();
                        var done = Label();
                        var end = Label();
                        flats.Add(new FlatForEachInit(forEach.Syntax, forEach));
                        flats.Add(new FlatLabel(forEach.Syntax, test));
                        flats.Add(new FlatForEachNext(forEach.Syntax, forEach, done));
                        loops.Push((ExitKind.For, end));
                        Block(forEach.Body);
                        loops.Pop();
                        flats.Add(new FlatJump(forEach.NextSyntax, test));
                        flats.Add(new FlatLabel(forEach.Syntax, done));
                        flats.Add(new FlatForEachDone(forEach.NextSyntax, forEach));
                        flats.Add(new FlatLabel(forEach.Syntax, end));
                        flats.Add(new FlatForEachEnd(forEach.NextSyntax, forEach));
                        return;
                    }

                case BoundDoLoop doLoop:
                    {
                        var top = Label();
                        var end = Label();
                        flats.Add(new FlatLabel(doLoop.Syntax, top));
                        if (doLoop.TopCondition is not null)
                        {
                            flats.Add(new FlatJumpIf(doLoop.Syntax, doLoop.TopCondition, WhenTrue: doLoop.TopIsUntil, end));
                        }

                        loops.Push((ExitKind.Do, end));
                        Block(doLoop.Body);
                        loops.Pop();
                        if (doLoop.BottomCondition is not null)
                        {
                            flats.Add(new FlatJumpIf(doLoop.LoopSyntax, doLoop.BottomCondition, WhenTrue: !doLoop.BottomIsUntil, top));
                        }
                        else
                        {
                            flats.Add(new FlatJump(doLoop.LoopSyntax, top));
                        }

                        flats.Add(new FlatLabel(doLoop.Syntax, end));
                        return;
                    }

                case BoundWith with:
                    flats.Add(new FlatWith(with.Syntax, with));
                    Block(with.Body);
                    return;
                case BoundExit { Kind: not ExitKind.Procedure } exit:
                    flats.Add(new FlatJump(exit.Syntax, loops.First(l => l.Kind == exit.Kind).EndId));
                    return;
                default:
                    flats.Add(new FlatSimple(statement.Syntax, statement));
                    return;
            }
        }
    }

    private void EmitDispatch(BoundProcedure bound)
    {
        var flattener = new Flattener();
        flattener.Block(bound.Body);

        // Resolve labels to statement indexes; a label takes the index of the statement that follows it.
        var statements = new List<Flat>();
        var targets = new Dictionary<int, int>();
        var numbered = new List<(int Offset, int Number)>();
        foreach (var flat in flattener.Flats)
        {
            switch (flat)
            {
                case FlatLabel label:
                    targets[label.Id] = statements.Count;
                    break;
                case FlatSimple { Statement: BoundLabel { Label: var symbol } }:
                    symbol.Target = statements.Count;
                    if (symbol.IsLineNumber)
                    {
                        numbered.Add((statements.Count, symbol.Number));
                    }

                    break;
                default:
                    statements.Add(flat);
                    break;
            }
        }

        var end = statements.Count;
        erlLines = ErlLines(numbered, end);
        writer.HiddenLine($"var __frame = new {R}ProcedureFrame();");
        if (erlLines is not null)
        {
            writer.HiddenLine($"static int __erl(int pc) => pc switch {{ {ErlArms(erlLines)} }};");
        }

        writer.HiddenLine("int __pc = 0;");
        writer.HiddenLine("while (true)");
        writer.Open();
        writer.HiddenLine("try");
        writer.Open();
        writer.HiddenLine("switch (__pc)");
        writer.Open();
        for (var index = 0; index < statements.Count; index++)
        {
            writer.HiddenLine($"case {Pc(index)}:");
            writer.Open();
            writer.HiddenLine($"__pc = {Pc(index)};");
            if (EmitFlat(statements[index], index, targets))
            {
                writer.HiddenLine($"goto case {Pc(index + 1)};");
            }

            writer.Close();
        }

        writer.HiddenLine($"case {Pc(end)}:");
        writer.Open();
        writer.MappedLine(Line(bound.Symbol.Syntax.EndStatement), "__frame.Exit(); " + ReturnCode());
        writer.Close();
        writer.HiddenLine("default:");
        writer.Open();
        writer.HiddenLine($"throw new global::System.InvalidOperationException(\"Unreachable statement index \" + __pc);");
        writer.Close();
        writer.Close();
        writer.Close();
        writer.HiddenLine($"catch (global::System.Exception __ex) when (__frame.TryHandle(__ex, __pc, __pc + 1, {(erlLines is null ? "0" : "__erl(__pc)")}, out __pc))");
        writer.Open();
        writer.Close();
        writer.Close();
    }

    private static string Pc(int index) => index.ToString(CultureInfo.InvariantCulture);

    /// <summary>
    /// The line Erl reports for an error at each statement index, resolved the way Excel resolves
    /// it (Errors golden, forty cases; docs/vba-quirks.md, Erl). The procedure's numbered lines
    /// form a table of (code offset, line number) in source order, a declaration taking the offset
    /// of the code that follows it. An error at offset p reports the last entry when its offset
    /// is at or before p; else the first entry when its offset is exactly p; else the last entry
    /// at or before p, or 0. The middle rule is the quirk: with nothing but declarations between
    /// the first numbered line and the failing statement, and a numbered line after it, Erl is
    /// the first line number rather than the failing one. MS-VBAL 5.4.4 describes the failing
    /// line only. The table is static, so a statement reached again by a jump after a numbered
    /// line ran still reports its own entry.
    /// </summary>
    private static int[]? ErlLines(List<(int Offset, int Number)> numbered, int count)
    {
        if (numbered.Count == 0)
        {
            return null;
        }

        var lines = new int[count];
        var (firstOffset, firstNumber) = numbered[0];
        var (lastOffset, lastNumber) = numbered[^1];
        for (var pc = 0; pc < count; pc++)
        {
            if (lastOffset <= pc)
            {
                lines[pc] = lastNumber;
                continue;
            }

            if (firstOffset == pc)
            {
                lines[pc] = firstNumber;
                continue;
            }

            foreach (var (offset, number) in numbered)
            {
                if (offset > pc)
                {
                    break;
                }

                lines[pc] = number;
            }
        }

        return lines;
    }

    /// <summary>The arms of the Erl switch: one per run of statements that report the same line, zero left to the default.</summary>
    private static string ErlArms(int[] lines)
    {
        var arms = new List<string>();
        for (var start = 0; start < lines.Length;)
        {
            var end = start;
            while (end + 1 < lines.Length && lines[end + 1] == lines[start])
            {
                end++;
            }

            if (lines[start] != 0)
            {
                var range = start == end ? Pc(start) : $">= {Pc(start)} and <= {Pc(end)}";
                arms.Add($"{range} => {Pc(lines[start])}");
            }

            start = end + 1;
        }

        arms.Add("_ => 0");
        return string.Join(", ", arms);
    }

    private string Erl(int index) => Pc(erlLines is null ? 0 : erlLines[index]);

    /// <summary>Emits one flat statement; returns true when control falls through to the next one.</summary>
    private bool EmitFlat(Flat flat, int index, Dictionary<int, int> targets)
    {
        string Target(int id) => Pc(targets[id]);
        string LabelTarget(LabelSymbol label) => Pc(label.Target < 0 ? index + 1 : label.Target);

        switch (flat)
        {
            case FlatJump jump:
                writer.MappedLine(Line(jump.Syntax), $"goto case {Target(jump.Id)};");
                return false;
            case FlatJumpIf jump:
                EmitSimple(Line(jump.Syntax), $"if ({(jump.WhenTrue ? string.Empty : "!")}{Condition(jump.Condition)}) goto case {Target(jump.Id)};");
                return true;
            case FlatForTest test:
                {
                    var counter = test.Loop.Limit.Type.IsVariant ? Convert(EmitExpression(test.Loop.Counter), VbaType.Variant) : EmitExpression(test.Loop.Counter).Code;
                    EmitSimple(Line(test.Syntax), $"if (!{R}ForLoop.Continues({counter}, {test.Loop.Limit.EmitName}, {test.Loop.Increment.EmitName})) goto case {Target(test.EndId)};");
                    return true;
                }

            case FlatForNext next:
                {
                    if (!next.Loop.Limit.Type.IsVariant)
                    {
                        // A counter the runtime steps natively (ROADMAP.md D-J, M7 E).
                        var native = EmitExpression(next.Loop.Counter).Code;
                        EmitSimple(Line(next.Syntax), Store(next.Loop.Counter, new Emitted($"{R}ForLoop.Next({native}, {next.Loop.Increment.EmitName})", next.Loop.Counter.Type)) + ";");
                    }
                    else
                    {
                        var counter = Convert(EmitExpression(next.Loop.Counter), VbaType.Variant);
                        var declared = DeclaredFlags(next.Loop.Counter.Type.IsVariant, next.Loop.Step?.Type.IsVariant ?? false);
                        EmitSimple(Line(next.Syntax), Store(next.Loop.Counter, new Emitted($"{R}Operators.Add({counter}, {next.Loop.Increment.EmitName}, {declared})", VbaType.Variant)) + ";");
                    }
                    writer.MappedLine(Line(next.Syntax), $"goto case {Target(next.TestId)};");
                    return false;
                }

            case FlatForEachInit init:
                EmitSimple(Line(init.Syntax), $"{init.Loop.Enumerator.EmitName} = {R}ForEachEnumerator.Create({Convert(EmitExpression(init.Loop.Collection), VbaType.Variant)}, __scope);");
                return true;
            case FlatForEachNext next:
                EmitSimple(Line(next.Syntax), $"if (!{next.Loop.Enumerator.EmitName}.MoveNext()) goto case {Target(next.DoneId)};");
                EmitSimple(Line(next.Syntax), Store(next.Loop.Variable, new Emitted($"{next.Loop.Enumerator.EmitName}.Current", VbaType.Variant)) + ";");
                return true;
            case FlatForEachDone done:
                if (done.Loop.Variable.Type.IsVariant)
                {
                    EmitSimple(Line(done.Syntax), $"if ({done.Loop.Enumerator.EmitName}.IsArray) {Store(done.Loop.Variable, new Emitted($"{VariantType}.Empty", VbaType.Variant))};");
                }

                return true;
            case FlatWith with:
                EmitSimple(Line(with.Syntax), WithStore(with.With));
                return true;
            case FlatForEachEnd end:
                // What the enumerator held for the loop goes with it, an Exit For included (D18).
                writer.HiddenLine($"{R}ObjectRefs.Release(ref {end.Loop.Enumerator.EmitName});");
                return true;
            case FlatCaseTest test:
                {
                    var condition = string.Join(" || ", test.Block.Clauses.Select(clause => CaseTest(test.Select, test.Select.Selector, clause)));
                    EmitSimple(Line(test.Syntax), $"if (!({condition})) goto case {Target(test.NextId)};");
                    return true;
                }

            case FlatSimple { ResumeNextId: { } resumeId, Statement: var guarded }:
                {
                    // An error here resumes past the whole block: a nested handler names the target.
                    writer.HiddenLine("try");
                    writer.Open();
                    var fallsThrough = EmitDispatchStatement(guarded, index, LabelTarget);
                    writer.Close();
                    writer.HiddenLine($"catch (global::System.Exception __inner) when (__frame.TryHandle(__inner, {Pc(index)}, {Target(resumeId)}, {Erl(index)}, out __pc))");
                    writer.Open();
                    writer.HiddenLine("continue;");
                    writer.Close();
                    return fallsThrough;
                }

            case FlatSimple { Statement: var statement }:
                return EmitDispatchStatement(statement, index, LabelTarget);
            default:
                Report(flat.Syntax, "Internal error: unknown flat statement.");
                return true;
        }
    }

    private bool EmitDispatchStatement(BoundStatement statement, int index, Func<LabelSymbol, string> target)
    {
        var line = Line(statement.Syntax);
        switch (statement)
        {
            case BoundOnError onError:
                writer.MappedLine(line, onError.Mode switch
                {
                    OnErrorMode.ResumeNext => "__frame.OnErrorResumeNext();",
                    OnErrorMode.Disable => "__frame.OnErrorGoToZero();",
                    OnErrorMode.Dismiss => "__frame.OnErrorGoToMinusOne();",
                    _ => $"__frame.OnErrorGoTo({target(onError.Label!)});",
                });
                return true;
            case BoundResume resume:
                writer.MappedLine(line, resume.Kind switch
                {
                    ResumeKind.Retry => "__pc = __frame.Resume(); continue;",
                    ResumeKind.Next => "__pc = __frame.ResumeNext(); continue;",
                    _ => $"__pc = __frame.ResumeAt({target(resume.Label!)}); continue;",
                });
                return false;
            case BoundGoTo goTo:
                writer.MappedLine(line, $"goto case {target(goTo.Label)};");
                return false;
            case BoundGoSub goSub:
                writer.MappedLine(line, $"__frame.GoSub({Pc(index + 1)}); goto case {target(goSub.Label)};");
                return false;
            case BoundReturn:
                writer.MappedLine(line, "__pc = __frame.Return(); continue;");
                return false;
            case BoundOnGoTo onGoTo:
                {
                    var selector = Convert(EmitExpression(onGoTo.Selector), VbaType.Variant);
                    var labels = "[" + string.Join(", ", onGoTo.Labels.Select(target)) + "]";
                    var next = Pc(index + 1);
                    if (onGoTo.IsGoSub)
                    {
                        EmitSimple(line, $"{{ var __jump = {R}ComputedJump.Target({selector}, {next}, {labels}); if (__jump != {next}) __frame.GoSub({next}); __pc = __jump; continue; }}");
                    }
                    else
                    {
                        EmitSimple(line, $"__pc = {R}ComputedJump.Target({selector}, {next}, {labels}); continue;");
                    }

                    return false;
                }

            case BoundExit { Kind: ExitKind.Procedure }:
                writer.MappedLine(line, "__frame.Exit(); " + ReturnCode());
                return false;
            case BoundEnd:
                writer.MappedLine(line, EndCode());
                return false;
            default:
                {
                    var code = SimpleStatementCode(statement);
                    if (code is not null)
                    {
                        EmitSimple(statement, code);
                    }

                    return true;
                }
        }
    }
}
