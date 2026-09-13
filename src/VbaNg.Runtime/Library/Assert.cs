namespace VbaNg.Runtime.Library;

/// <summary>
/// The <c>Assert</c> module test procedures use (ARCHITECTURE.md section 8): a vba-ng addition
/// with no VBA counterpart, resolved by the binder only when the project declares nothing named
/// Assert. Values compare with VBA's <c>=</c> under the calling module's Option Compare, so
/// <c>Assert.AreEqual 3, "3"</c> passes as <c>3 = "3"</c> does. A failure throws
/// <see cref="AssertFailedException"/>, which <c>On Error</c> cannot handle.
/// </summary>
public static class Assert
{
    /// <summary>Passes when <c>expected = actual</c>, both are Null, or both are the same object.</summary>
    public static void AreEqual(in Variant expected, in Variant actual, in Variant message, CompareMode mode = CompareMode.Binary)
    {
        if (!Same(expected, actual, mode))
        {
            throw Failure("AreEqual", $"Expected:<{Describe(expected)}>. Actual:<{Describe(actual)}>.", message);
        }
    }

    public static void AreNotEqual(in Variant notExpected, in Variant actual, in Variant message, CompareMode mode = CompareMode.Binary)
    {
        if (Same(notExpected, actual, mode))
        {
            throw Failure("AreNotEqual", $"Expected any value except:<{Describe(notExpected)}>. Actual:<{Describe(actual)}>.", message);
        }
    }

    /// <summary>Passes when the condition is True as <c>If</c> would take it; Null counts as False.</summary>
    public static void IsTrue(in Variant condition, in Variant message)
    {
        if (!Coerce.ToCondition(condition))
        {
            throw Failure("IsTrue", $"Condition:<{Describe(condition)}>.", message);
        }
    }

    public static void IsFalse(in Variant condition, in Variant message)
    {
        if (Coerce.ToCondition(condition))
        {
            throw Failure("IsFalse", $"Condition:<{Describe(condition)}>.", message);
        }
    }

    public static void IsNothing(in Variant value, in Variant message)
    {
        if (!value.IsNothing)
        {
            throw Failure("IsNothing", $"Actual:<{Describe(value)}>.", message);
        }
    }

    /// <summary>Passes for an object reference that is set; a value that is not an object fails.</summary>
    public static void IsNotNothing(in Variant value, in Variant message)
    {
        if (!value.IsObject || value.IsNothing)
        {
            throw Failure("IsNotNothing", $"Actual:<{Describe(value)}>.", message);
        }
    }

    public static void Fail(in Variant message) => throw Failure("Fail", null, message);

    private static bool Same(in Variant expected, in Variant actual, CompareMode mode)
    {
        if (expected.IsNull || actual.IsNull)
        {
            return expected.IsNull && actual.IsNull;
        }

        if (expected.IsObject || actual.IsObject)
        {
            return expected.IsObject && actual.IsObject && Coerce.ToCondition(Operators.Is(expected, actual));
        }

        return Coerce.ToCondition(Operators.Equal(expected, actual, DeclaredTypes.None, mode));
    }

    private static AssertFailedException Failure(string assertion, string? detail, in Variant message)
    {
        var text = "Assert." + assertion + " failed.";
        if (detail is not null)
        {
            text += " " + detail;
        }

        if (!message.IsMissing && !message.IsEmpty && !message.IsNull)
        {
            var userText = Coerce.ToString(message);
            if (userText.Length > 0)
            {
                text += " " + userText;
            }
        }

        return new AssertFailedException(text);
    }

    /// <summary>The readable form of a value for a failure message: the text VBA would show, or the type when it has no text.</summary>
    private static string Describe(in Variant value)
    {
        if (value.IsNull)
        {
            return "Null";
        }

        if (value.IsEmpty)
        {
            return "Empty";
        }

        if (value.IsNothing)
        {
            return "Nothing";
        }

        if (value.IsObject || value.IsArray || value.IsError)
        {
            return Information.TypeNameText(value);
        }

        return Coerce.ToString(value);
    }
}
