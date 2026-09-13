using System.Globalization;

namespace VbaNg.Compiler.Syntax;

/// <summary>
/// Conditional compilation (MS-VBAL 3.4): evaluates #Const, #If, #ElseIf, #Else, and #End If
/// lines and removes them and every excluded logical line from the token stream. Removed lines
/// survive as <see cref="SyntaxKind.ConditionalDirectiveTrivia"/> and
/// <see cref="SyntaxKind.DisabledTextTrivia"/> on the next kept token, so the tree still
/// round-trips. Lexical diagnostics inside excluded lines are dropped, as VBA ignores that text.
/// </summary>
public static class Preprocessor
{
    public static IReadOnlyList<SyntaxToken> Process(
        IReadOnlyList<SyntaxToken> tokens,
        IReadOnlyList<Diagnostic> lexerDiagnostics,
        SourceText source,
        string filePath,
        ParseOptions options,
        List<Diagnostic> diagnostics)
    {
        ArgumentNullException.ThrowIfNull(tokens);
        ArgumentNullException.ThrowIfNull(lexerDiagnostics);
        ArgumentNullException.ThrowIfNull(source);
        ArgumentNullException.ThrowIfNull(filePath);
        ArgumentNullException.ThrowIfNull(options);
        ArgumentNullException.ThrowIfNull(diagnostics);

        var state = new State(tokens, source, filePath, options, diagnostics);
        state.Run();
        foreach (var diagnostic in lexerDiagnostics)
        {
            if (!state.IsExcluded(diagnostic.Line))
            {
                diagnostics.Add(diagnostic);
            }
        }

        return state.Output;
    }

    private sealed class State(IReadOnlyList<SyntaxToken> tokens, SourceText source, string filePath, ParseOptions options, List<Diagnostic> diagnostics)
    {
        private readonly Dictionary<string, object?> constants = new(options.ConditionalConstants, StringComparer.OrdinalIgnoreCase);
        private readonly Stack<Region> regions = new();
        private readonly List<SyntaxTrivia> pending = [];
        private readonly List<(int First, int Last)> excludedLines = [];

        public List<SyntaxToken> Output { get; } = [];

        private bool Included => regions.Count == 0 || regions.Peek().Included;

        public bool IsExcluded(int line) => excludedLines.Any(r => line >= r.First && line <= r.Last);

        public void Run()
        {
            var i = 0;
            while (i < tokens.Count)
            {
                var first = tokens[i];
                if (first.Kind == SyntaxKind.EndOfFileToken)
                {
                    Output.Add(pending.Count == 0 ? first : first.WithLeadingTrivia([.. pending, .. first.LeadingTrivia]));
                    pending.Clear();
                    return;
                }

                var last = i;
                while (!tokens[last].EndsLine && tokens[last + 1].Kind != SyntaxKind.EndOfFileToken)
                {
                    last++;
                }

                if (first.Kind == SyntaxKind.HashToken && IsDirective(i, last))
                {
                    RemoveLine(i, last, SyntaxKind.ConditionalDirectiveTrivia);
                }
                else if (!Included)
                {
                    RemoveLine(i, last, SyntaxKind.DisabledTextTrivia);
                }
                else
                {
                    Keep(i, last);
                }

                i = last + 1;
            }
        }

        private void Keep(int first, int last)
        {
            var token = tokens[first];
            Output.Add(pending.Count == 0 ? token : token.WithLeadingTrivia([.. pending, .. token.LeadingTrivia]));
            pending.Clear();
            for (var i = first + 1; i <= last; i++)
            {
                Output.Add(tokens[i]);
            }
        }

        private void RemoveLine(int first, int last, SyntaxKind kind)
        {
            var start = tokens[first].Start;
            var end = tokens[last].FullEnd;
            pending.AddRange(tokens[first].LeadingTrivia);
            if (kind == SyntaxKind.DisabledTextTrivia && pending.Count > 0 && pending[^1].Kind == SyntaxKind.DisabledTextTrivia)
            {
                // Merge adjacent excluded lines into one trivia.
                var previous = pending[^1];
                pending[^1] = new SyntaxTrivia(kind, previous.Text + source.Text[start..end]);
            }
            else
            {
                pending.Add(new SyntaxTrivia(kind, source.Text[start..end]));
            }

            if (kind == SyntaxKind.DisabledTextTrivia)
            {
                var (firstLine, _) = source.GetLinePosition(start);
                var (lastLine, _) = source.GetLinePosition(Math.Max(start, end - 1));
                excludedLines.Add((firstLine, lastLine));
            }
        }

        // cc-if / cc-elseif / cc-else / cc-endif / cc-const (MS-VBAL 3.4.1, 3.4.2).
        private bool IsDirective(int first, int last)
        {
            if (first + 1 > last)
            {
                return false;
            }

            var word = tokens[first + 1];
            switch (word.Kind)
            {
                case SyntaxKind.IfKeyword:
                    regions.Push(Included ? Region.Start(EvaluateCondition(first + 2, last)) : Region.Skipped);
                    return true;

                case SyntaxKind.ElseIfKeyword:
                    if (regions.Count == 0)
                    {
                        Report(word, "#ElseIf without #If.");
                        return true;
                    }

                    var region = regions.Pop();
                    regions.Push(region.ElseIf(region.Live && !region.AnyTaken && EvaluateCondition(first + 2, last)));
                    return true;

                case SyntaxKind.ElseKeyword:
                    if (regions.Count == 0)
                    {
                        Report(word, "#Else without #If.");
                        return true;
                    }

                    regions.Push(regions.Pop().Else());
                    return true;

                case SyntaxKind.EndKeyword or SyntaxKind.EndIfKeyword:
                    if (regions.Count == 0)
                    {
                        Report(word, "#End If without #If.");
                        return true;
                    }

                    regions.Pop();
                    return true;

                case SyntaxKind.ConstKeyword:
                    if (Included)
                    {
                        DefineConstant(first + 2, last);
                    }

                    return true;

                default:
                    return false;
            }
        }

        private bool EvaluateCondition(int first, int last)
        {
            // Drop the trailing "Then".
            var end = last;
            if (end >= first && tokens[end].Kind == SyntaxKind.ThenKeyword)
            {
                end--;
            }
            else
            {
                Report(tokens[last], "Expected: Then");
            }

            var value = Evaluate(first, end);
            return ConditionalCompilationEvaluator.ToBoolean(value);
        }

        private void DefineConstant(int first, int last)
        {
            if (first > last || tokens[first].Kind != SyntaxKind.IdentifierToken || first + 1 > last || tokens[first + 1].Kind != SyntaxKind.EqualsToken)
            {
                Report(tokens[Math.Min(first, last)], "Expected: #Const name = expression");
                return;
            }

            constants[tokens[first].NameValue] = Evaluate(first + 2, last);
        }

        private object? Evaluate(int first, int last)
        {
            if (first > last)
            {
                Report(tokens[Math.Min(first, tokens.Count - 1)], "Expected: expression");
                return null;
            }

            var slice = new List<SyntaxToken>();
            for (var i = first; i <= last; i++)
            {
                slice.Add(tokens[i]);
            }

            var end = tokens[last].End;
            slice.Add(new SyntaxToken(SyntaxKind.EndOfFileToken, string.Empty, end));
            var parser = new Parser(slice, source, filePath, diagnostics);
            var expression = parser.ParseStandaloneExpression();
            try
            {
                return ConditionalCompilationEvaluator.Evaluate(expression, constants);
            }
            catch (ConditionalCompilationException ex)
            {
                var (line, column) = source.GetLinePosition(tokens[first].Start);
                diagnostics.Add(new Diagnostic(DiagnosticIds.ConditionalCompilationError, DiagnosticSeverity.Error, ex.Message, filePath, line, column));
                return null;
            }
        }

        private void Report(SyntaxToken at, string message)
        {
            var (line, column) = source.GetLinePosition(at.Start);
            diagnostics.Add(new Diagnostic(DiagnosticIds.SyntaxError, DiagnosticSeverity.Error, message, filePath, line, column));
        }

        /// <summary>One #If block: whether its enclosing text is included (Live), whether a branch has been taken, and the current branch.</summary>
        private readonly record struct Region(bool Live, bool AnyTaken, bool Included)
        {
            public static Region Skipped => new(false, true, false);

            public static Region Start(bool condition) => new(true, condition, condition);

            public Region ElseIf(bool condition) => new(Live, AnyTaken || condition, condition);

            public Region Else() => new(Live, true, Live && !AnyTaken);
        }
    }
}

/// <summary>Raised when a conditional compilation expression uses something MS-VBAL 3.4.1 does not allow or cannot be evaluated.</summary>
public sealed class ConditionalCompilationException : Exception
{
    public ConditionalCompilationException()
    {
    }

    public ConditionalCompilationException(string message)
        : base(message)
    {
    }

    public ConditionalCompilationException(string message, Exception innerException)
        : base(message, innerException)
    {
    }
}

/// <summary>
/// Evaluates conditional compilation expressions (MS-VBAL 3.4.1) over literals, constants, the
/// listed operators, and the listed intrinsic functions. Values are null (Empty), bool, long,
/// double, or string. Undefined constants are Empty. The arithmetic here is deliberately small;
/// the runtime library's Variant operators replace it once they exist with goldens (CLAUDE.md R2).
/// </summary>
public static class ConditionalCompilationEvaluator
{
    public static object? Evaluate(ExpressionSyntax expression, IReadOnlyDictionary<string, object?> constants)
    {
        ArgumentNullException.ThrowIfNull(expression);
        ArgumentNullException.ThrowIfNull(constants);
        return expression switch
        {
            LiteralExpressionSyntax literal => Literal(literal.Token),
            IdentifierNameSyntax name => constants.TryGetValue(name.Name, out var value) ? value : null,
            ParenthesizedExpressionSyntax parenthesized => Evaluate(parenthesized.Expression, constants),
            UnaryExpressionSyntax unary => Unary(unary.OperatorToken.Kind, Evaluate(unary.Operand, constants)),
            BinaryExpressionSyntax binary => Binary(binary.OperatorToken.Kind, Evaluate(binary.Left, constants), Evaluate(binary.Right, constants)),
            IndexExpressionSyntax { Expression: IdentifierNameSyntax function } call => Intrinsic(function.Identifier, call.Arguments, constants),
            _ => throw new ConditionalCompilationException("Expression is not allowed in conditional compilation."),
        };
    }

    /// <summary>Let-coercion to Boolean (MS-VBAL 5.6.9.1): numbers are True when nonzero, Empty is False.</summary>
    public static bool ToBoolean(object? value) => value switch
    {
        null => false,
        bool b => b,
        long l => l != 0,
        double d => d != 0,
        string s when bool.TryParse(s, out var b) => b,
        string s when double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var d) => d != 0,
        _ => throw new ConditionalCompilationException("Type mismatch in conditional compilation expression."),
    };

    private static object? Literal(SyntaxToken token) => token.Kind switch
    {
        SyntaxKind.IntegerLiteralToken => LiteralValues.ParseInteger(token.Text),
        SyntaxKind.FloatLiteralToken => LiteralValues.ParseFloat(token.Text),
        SyntaxKind.StringLiteralToken => token.Value as string ?? string.Empty,
        SyntaxKind.TrueKeyword => true,
        SyntaxKind.FalseKeyword => false,
        SyntaxKind.EmptyKeyword or SyntaxKind.NothingKeyword or SyntaxKind.NullKeyword => null,
        _ => throw new ConditionalCompilationException("Literal is not allowed in conditional compilation."),
    };

    private static object? Unary(SyntaxKind op, object? operand) => op switch
    {
        SyntaxKind.MinusToken => operand switch
        {
            null => 0L,
            bool b => b ? 1L : 0L,
            long l => -l,
            double d => -d,
            string s => -ToDouble(s),
            _ => throw Mismatch(),
        },
        SyntaxKind.PlusToken => operand,
        SyntaxKind.NotKeyword => operand switch
        {
            null => -1L,
            bool b => !b,
            long l => ~l,
            double d => ~(long)Math.Round(d, MidpointRounding.ToEven),
            _ => throw Mismatch(),
        },
        _ => throw new ConditionalCompilationException("Operator is not allowed in conditional compilation."),
    };

    private static object? Binary(SyntaxKind op, object? left, object? right)
    {
        switch (op)
        {
            case SyntaxKind.AmpersandToken:
                return ToText(left) + ToText(right);
            case SyntaxKind.EqualsToken:
                return Compare(left, right) == 0;
            case SyntaxKind.LessThanGreaterThanToken:
                return Compare(left, right) != 0;
            case SyntaxKind.LessThanToken:
                return Compare(left, right) < 0;
            case SyntaxKind.GreaterThanToken:
                return Compare(left, right) > 0;
            case SyntaxKind.LessThanEqualsToken:
                return Compare(left, right) <= 0;
            case SyntaxKind.GreaterThanEqualsToken:
                return Compare(left, right) >= 0;
            case SyntaxKind.AndKeyword:
                return Logical(left, right, (a, b) => a & b, (a, b) => a && b);
            case SyntaxKind.OrKeyword:
                return Logical(left, right, (a, b) => a | b, (a, b) => a || b);
            case SyntaxKind.XorKeyword:
                return Logical(left, right, (a, b) => a ^ b, (a, b) => a ^ b);
            case SyntaxKind.EqvKeyword:
                return Logical(left, right, (a, b) => ~(a ^ b), (a, b) => a == b);
            case SyntaxKind.ImpKeyword:
                return Logical(left, right, (a, b) => ~a | b, (a, b) => !a || b);
            case SyntaxKind.IsKeyword:
                return left is null && right is null;
            case SyntaxKind.LikeKeyword:
                throw new ConditionalCompilationException("Like is not supported in conditional compilation expressions yet.");
        }

        if (left is string or null && right is string or null && op == SyntaxKind.PlusToken && (left is string || right is string))
        {
            return ToText(left) + ToText(right);
        }

        if (left is long a && right is long b)
        {
            return op switch
            {
                SyntaxKind.PlusToken => checked(a + b),
                SyntaxKind.MinusToken => checked(a - b),
                SyntaxKind.AsteriskToken => checked(a * b),
                SyntaxKind.SlashToken => b == 0 ? throw new ConditionalCompilationException("Division by zero.") : (double)a / b,
                SyntaxKind.BackslashToken => b == 0 ? throw new ConditionalCompilationException("Division by zero.") : a / b,
                SyntaxKind.ModKeyword => b == 0 ? throw new ConditionalCompilationException("Division by zero.") : a % b,
                SyntaxKind.CaretToken => Math.Pow(a, b),
                _ => throw new ConditionalCompilationException("Operator is not allowed in conditional compilation."),
            };
        }

        var x = ToDouble(left);
        var y = ToDouble(right);
        return op switch
        {
            SyntaxKind.PlusToken => x + y,
            SyntaxKind.MinusToken => x - y,
            SyntaxKind.AsteriskToken => x * y,
            SyntaxKind.SlashToken => y == 0 ? throw new ConditionalCompilationException("Division by zero.") : x / y,
            SyntaxKind.BackslashToken => (long)y == 0 ? throw new ConditionalCompilationException("Division by zero.") : (long)x / (long)y,
            SyntaxKind.ModKeyword => (long)y == 0 ? throw new ConditionalCompilationException("Division by zero.") : (long)x % (long)y,
            SyntaxKind.CaretToken => Math.Pow(x, y),
            _ => throw new ConditionalCompilationException("Operator is not allowed in conditional compilation."),
        };
    }

    private static object Logical(object? left, object? right, Func<long, long, long> bitwise, Func<bool, bool, bool> boolean)
    {
        if (left is bool a && right is bool b)
        {
            return boolean(a, b);
        }

        return bitwise(ToLong(left), ToLong(right));
    }

    private static int Compare(object? left, object? right)
    {
        if (left is string s && right is string t)
        {
            return string.Compare(s, t, StringComparison.Ordinal);
        }

        if (left is string || right is string)
        {
            // A string compared with a number is greater in VBA's Variant comparison.
            return left is string ? 1 : -1;
        }

        return ToDouble(left).CompareTo(ToDouble(right));
    }

    private static object? Intrinsic(SyntaxToken function, ArgumentListSyntax arguments, IReadOnlyDictionary<string, object?> constants)
    {
        var name = function.NameValue;
        if (function.Kind == SyntaxKind.IdentifierToken)
        {
            throw new ConditionalCompilationException($"'{name}' is not allowed in a conditional compilation expression.");
        }

        if (arguments.Arguments.Count != 1 || arguments.Arguments[0].Expression is null)
        {
            throw new ConditionalCompilationException($"'{name}' takes one argument in a conditional compilation expression.");
        }

        var value = Evaluate(arguments.Arguments[0].Expression!, constants);
        return function.Kind switch
        {
            SyntaxKind.IntKeyword => Math.Floor(ToDouble(value)),
            SyntaxKind.FixKeyword => Math.Truncate(ToDouble(value)),
            SyntaxKind.AbsKeyword => value is long l ? Math.Abs(l) : Math.Abs(ToDouble(value)),
            SyntaxKind.SgnKeyword => (long)Math.Sign(ToDouble(value)),
            SyntaxKind.LenKeyword => (long)ToText(value).Length,
            SyntaxKind.LenBKeyword => (long)ToText(value).Length * 2,
            SyntaxKind.CBoolKeyword => ToBoolean(value),
            SyntaxKind.CByteKeyword or SyntaxKind.CIntKeyword or SyntaxKind.CLngKeyword or SyntaxKind.CLngLngKeyword or SyntaxKind.CLngPtrKeyword
                => (long)Math.Round(ToDouble(value), MidpointRounding.ToEven),
            SyntaxKind.CSngKeyword or SyntaxKind.CDblKeyword or SyntaxKind.CCurKeyword => ToDouble(value),
            SyntaxKind.CStrKeyword => ToText(value),
            SyntaxKind.CVarKeyword => value,
            SyntaxKind.CDateKeyword => throw new ConditionalCompilationException("CDate is not supported in conditional compilation expressions yet."),
            _ => throw new ConditionalCompilationException($"'{name}' is not allowed in a conditional compilation expression."),
        };
    }

    private static string ToText(object? value) => value switch
    {
        null => string.Empty,
        bool b => b ? "True" : "False",
        long l => l.ToString(CultureInfo.InvariantCulture),
        double d => d.ToString(CultureInfo.InvariantCulture),
        string s => s,
        _ => throw Mismatch(),
    };

    private static double ToDouble(object? value) => value switch
    {
        null => 0,
        bool b => b ? -1 : 0,
        long l => l,
        double d => d,
        string s when double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var d) => d,
        _ => throw Mismatch(),
    };

    private static long ToLong(object? value) => value switch
    {
        null => 0,
        bool b => b ? -1 : 0,
        long l => l,
        double d => (long)Math.Round(d, MidpointRounding.ToEven),
        string s when double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var d) => (long)Math.Round(d, MidpointRounding.ToEven),
        _ => throw Mismatch(),
    };

    private static ConditionalCompilationException Mismatch() => new("Type mismatch in conditional compilation expression.");
}

/// <summary>Numeric values of literal tokens (MS-VBAL 3.3.2), as far as conditional compilation needs them.</summary>
public static class LiteralValues
{
    /// <summary>The unsigned magnitude of an integer literal as a long; suffix and sign wrapping are applied by the runtime later.</summary>
    public static object ParseInteger(string text)
    {
        ArgumentNullException.ThrowIfNull(text);
        var digits = text.TrimEnd('%', '&', '^');
        if (digits.StartsWith('&'))
        {
            var radixChar = digits.Length > 1 ? char.ToUpperInvariant(digits[1]) : '\0';
            if (radixChar == 'H')
            {
                return Convert.ToInt64(digits[2..], 16);
            }

            return Convert.ToInt64(radixChar == 'O' ? digits[2..] : digits[1..], 8);
        }

        // Box explicitly: a long/double conditional would promote the long to double.
        return long.TryParse(digits, NumberStyles.None, CultureInfo.InvariantCulture, out var value)
            ? (object)value
            : double.Parse(digits, CultureInfo.InvariantCulture);
    }

    public static double ParseFloat(string text)
    {
        ArgumentNullException.ThrowIfNull(text);
        var digits = text.TrimEnd('!', '#', '@').Replace('D', 'E').Replace('d', 'E');
        return double.Parse(digits, NumberStyles.Float, CultureInfo.InvariantCulture);
    }
}
