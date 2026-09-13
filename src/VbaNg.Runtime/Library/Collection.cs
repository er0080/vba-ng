using System.Collections;

namespace VbaNg.Runtime.Library;

/// <summary>
/// The VBA Collection class (MS-VBAL 6.1.3.2? the Collection class of the standard library):
/// an ordered list of Variants with optional case-insensitive string keys. Error numbers and
/// the argument checks are as Excel raises them (Interaction goldens).
/// </summary>
public sealed class Collection : RuntimeObject, IVbaObject, IEnumerable<Variant>
{
    private readonly List<(string? Key, Variant Item)> entries = [];

    public string TypeName => "Collection";

    public int Count => entries.Count;

    /// <summary>A COM caller reaches the collection's members at VBA's dispids: Item, the default, then Add, Count, and Remove (ROADMAP.md M7 F).</summary>
    internal override IDispatchObject ComMembers => new LateBoundMembers(this, TypeName, [(0, "Item"), (1, "Add"), (2, "Count"), (3, "Remove"), (-4, "_NewEnum")]);

    /// <summary>The last reference to the collection goes: the objects it holds lose theirs, its strings and arrays are freed (ARCHITECTURE.md D18).</summary>
    protected override void OnLastRelease()
    {
        var held = entries.ToArray();
        entries.Clear();
        foreach (var entry in held)
        {
            ReleaseItem(entry.Item);
        }
    }

    /// <summary>[_NewEnum]: the enumerator a class hands to For Each through its own NewEnum (dispid -4).</summary>
    public VbaEnumerator NewEnum() => new(GetEnumerator(), GetEnumerator);

    public void Add(Variant item) => Add(item, Variant.Missing, Variant.Missing, Variant.Missing);

    /// <summary>Add(Item, [Key], [Before], [After]): omitted arguments are Missing; a duplicate key raises 457, both Before and After raise 5, a non-string key raises 13.</summary>
    public void Add(Variant item, Variant key, Variant before, Variant after)
    {
        string? keyText = null;
        if (!IsOmitted(key))
        {
            if (!key.IsString)
            {
                throw VbaErrors.TypeMismatch();
            }

            keyText = key.AsString();
            if (IndexOfKey(keyText) >= 0)
            {
                throw new VbaException(VbaErrors.DuplicateCollectionKey);
            }
        }

        if (!IsOmitted(before) && !IsOmitted(after))
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        var position = entries.Count;
        if (!IsOmitted(before))
        {
            position = Resolve(before);
        }
        else if (!IsOmitted(after))
        {
            position = Resolve(after) + 1;
        }

        // Item is ByVal: the collection holds a reference to an object item and its own copy of a string or an array item (MS-VBAL 5.5.1.2.5; ARCHITECTURE.md D18, D20).
        entries.Insert(position, (keyText, ObjectRefs.Own(item)));
    }

    /// <summary>Item(Index): a 1-based position (rounded) or a key; an unknown key or a position out of range raises 5 or 9.</summary>
    public Variant Item(Variant index) => entries[Resolve(index)].Item;

    /// <summary>Remove(Index): a position out of range raises 5 here, unlike Item, which raises 9 (Interaction golden).</summary>
    public void Remove(Variant index)
    {
        int position;
        try
        {
            position = Resolve(index);
        }
        catch (VbaException ex) when (ex.Number == VbaErrors.SubscriptOutOfRangeNumber)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        var removed = entries[position].Item;
        entries.RemoveAt(position);
        ReleaseItem(removed);
    }

    private static void ReleaseItem(in Variant item)
    {
        var held = item;
        ObjectRefs.Release(ref held);
    }

    public IEnumerator<Variant> GetEnumerator()
    {
        foreach (var entry in entries.ToArray())
        {
            yield return entry.Item;
        }
    }

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    private static bool IsOmitted(in Variant value) => value.IsMissing;

    private int IndexOfKey(string key) =>
        entries.FindIndex(e => e.Key is not null && string.Equals(e.Key, key, StringComparison.OrdinalIgnoreCase));

    private int Resolve(in Variant index)
    {
        switch (index.Type)
        {
            case VarType.String:
                {
                    var position = IndexOfKey(index.AsString());
                    if (position < 0)
                    {
                        throw VbaErrors.InvalidProcedureCall();
                    }

                    return position;
                }

            case VarType.Null or VarType.Empty or VarType.Object or VarType.Array or VarType.Error:
                throw VbaErrors.TypeMismatch();
            default:
                {
                    var position = Coerce.ToInt32(index);
                    if (position < 1 || position > entries.Count)
                    {
                        throw VbaErrors.SubscriptOutOfRange();
                    }

                    return position - 1;
                }
        }
    }
}
