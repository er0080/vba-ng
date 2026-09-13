# Diagnostics

Every diagnostic has a stable id that is never renumbered or reused. All of them come from
`src/VbaNg.Compiler`; the CLI prints them in the MSBuild canonical format:

```
C:\work\Sales.vbang\Module1.bas(12,9): error VBA0001: Expected: Then
```

Two of them are warnings: VBA0002 on a `.frm`, and VBA0024. Everything else is an error, and one
error fails the build. Where a message takes arguments, the shape is given rather than one
instance of it.

| Id | Severity | Meaning | Triggering snippet |
|----|----------|---------|--------------------|
| VBA0001 | error | Syntax error: the source could not be tokenized or parsed. The message names what was expected, in the VBE's style (`Expected: Then`, `Unexpected character '?'.`), and a statement outside a procedure body gives `Only declarations are allowed outside procedures.` | `If x` (no `Then`); `x = ?` |
| VBA0002 | error; warning for a `.frm` | Valid VBA that this version of the compiler does not carry: a `Declare` whose calling convention, parameter, or return type needs its own marshaling; an array element passed ByRef to a `Declare` from an array that does not keep its elements in native storage; a procedure `AddressOf` cannot hand over as a callback; an `As Any` parameter with no fixed-size argument; a `VarPtr` whose argument has no address of its own to give; `Me` in a standard module; an unimplemented intrinsic; a `MidB` assignment; or a statement or expression kind not carried yet. A `.frm` is the one warning: the project still builds, but only the form's signatures compile and every procedure of it raises error 438. | `Declare PtrSafe Sub F Lib "k" (ParamArray a() As Variant)`; any `.frm` for the warning |
| VBA0003 | error | Internal compiler error: the generated C# failed to compile. The message is `Internal compiler error <Roslyn id>: <Roslyn message>`, at the mapped VBA location. Always a compiler bug; report it with the module that caused it. | none |
| VBA0004 | error | A conditional compilation expression could not be evaluated. MS-VBAL 3.4.1 allows only literals, conditional constants, the listed operators, and the listed intrinsic functions; the message names what was refused (`'Foo' is not allowed in a conditional compilation expression.`, `Division by zero.`). | `#If Foo() Then` |
| VBA0005 | error | A name is not defined: an undeclared name under `Option Explicit` (MS-VBAL 5.2.1.2), a member an early-bound COM or class type does not have (`Method or data member not found: '<name>'.`), a missing enum member, or a class name used where a variable belongs. | `Option Explicit` then `x = 1`; `Dim d As Dictionary` then `d.Nope` |
| VBA0006 | error | Ambiguous name detected: two module-level members, two modules, two `Type` fields, two `Enum` members, or two parameters share a name. Two declarations in one procedure give `Duplicate declaration in current scope: '<name>'.` | `Public x As Long` above `Sub x()` |
| VBA0007 | error | Sub or Function not defined: a call names a procedure that no module of the project, the VBA library, or `Assert` has; also an `Event` or a `WithEvents` handler the source class does not declare. A missing member of a COM or class type is VBA0005 instead. | `Foo 1` with no `Sub Foo` |
| VBA0008 | error | Label not defined: `GoTo`, `GoSub`, `On Error GoTo`, or `Resume` names a label the procedure does not have. | `GoTo Done` with no `Done:` |
| VBA0009 | error | Wrong number of arguments or invalid property assignment: too many arguments, the wrong number of array dimensions, a named argument given twice, a positional argument after a named one, or a read of a property with no `Property Get`. | `Left("abc", 1, 2)` |
| VBA0010 | error | Argument not optional: a required parameter has no argument. | `Sub Foo(a)` called as `Foo` |
| VBA0011 | error | Expected array: an index, `ReDim`, `Erase`, `UBound`, or `LBound` applied to something that is not an array. | `Dim x As Long` then `x(1) = 2` |
| VBA0012 | error | User-defined type not defined: an `As` clause names a type that is neither a builtin, a `Type`, `Enum`, or class of the project, nor a type of a referenced library; `Implements` names something other than a class of the project; `As Any` appears outside a `Declare`. | `Dim r As Range` with no Excel reference in `vbang.json` |
| VBA0013 | error | Named argument not found: `name:=` names no parameter of the callee, the callee takes no named arguments (`InStr` and `StrComp`, as in VBA — docs/vba-quirks.md), or the target is late-bound and has no parameter names. | `Left(Text:="hello", Length:=2)` |
| VBA0014 | error | ByRef argument type mismatch: a variable of one declared type passed to a `ByRef` parameter of another (MS-VBAL 5.6.13.1), or a scalar where an array or user-defined type is expected. | `Dim x As Long` passed to `Sub Foo(s As String)` |
| VBA0015 | error | Constant expression required: a `Const` value, an array bound in `Dim`, or an `Optional` default is not made of literals, constants, and operators, or a `Const` refers to itself. | `Const X = Len("a")` |
| VBA0016 | error | A statement is not allowed where it appears: `Exit For` outside `For`, `Exit Do` outside `Do`, `Event` or `Implements` outside a class module, `WithEvents` inside a procedure or in a standard module, or a leading `.` or `!` with no `With` block. Which of `Exit Sub`, `Exit Function`, and `Exit Property` leaves a procedure is not checked, because VBA does not check it either (docs/vba-quirks.md). | `Exit For` in a plain `Sub` |
| VBA0017 | error | Object required: `Set` with a non-object variable, `Is` on values, `New` on a non-class, `With` on a non-object, or `!` applied to a value. | `Dim x As Long` then `Set x = Nothing` |
| VBA0018 | error | Duplicate label within the procedure. | `Done:` twice in one procedure |
| VBA0019 | error | The left side of an assignment is not a variable, array element, record field, or property. | `Len(s) = 3` |
| VBA0020 | error | Type mismatch the binder can see: an array assigned to a scalar, a user-defined type coerced to a Variant or assigned to another type, a Sub used as a value, an `As Decimal` declaration, a type or module name used as an expression, a non-numeric `For` counter, a `For Each` variable that is not Variant or Object, an `LSet` target that is neither a String nor a record, an array or record parameter declared `ByVal`, or `WithEvents` on a type that sources no events. | `Dim a() As Long` and `Dim x As Long`, then `x = a` |
| VBA0021 | error | A `'@Test` annotation on something other than a Public Sub without parameters. | `'@Test` above `Private Sub T()` |
| VBA0022 | error | The project manifest `vbang.json` is not valid JSON, or one of its references has no name. The message is `Invalid manifest: <reason>`, located at the offending line when the JSON reader gives one. | `vbang.json` holding `{ "references": [ { "name": "vbang" }` |
| VBA0023 | error | Can't find project or library: a manifest reference whose guid is registered nowhere and whose name is not one of the libraries vba-ng knows (Excel, Office, Scripting, stdole, VBIDE, MSForms) or `vbang`. | `{ "name": "Nope", "guid": "{00000000-0000-0000-0000-000000000001}", "version": "1.0" }` in `vbang.json` |
| VBA0024 | warning | A `.cls` carries a document module's attributes (`VB_PredeclaredId` and `VB_Exposed` both True) but the workbook beside the project folder has no object with that CodeName, so it compiles as a class and its event procedures never run; list it in the manifest's `documents` map if it is one. | `Sheet1.cls` with those attributes next to a workbook whose sheets carry no CodeNames |
