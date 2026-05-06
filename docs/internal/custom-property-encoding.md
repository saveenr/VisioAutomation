# Custom Property and User-Defined Cell encoding behavior

Reference document capturing what Visio actually does when `CustomPropertyCells` and `UserDefinedCellCells` fields are written through `CustomPropertyHelper.Set` / `UserDefinedCellHelper.Set`. Locked from characterization tests run on 2026-05-06.

This document exists so the matrix doesn't perish: the "what does Visio do with input X" knowledge is otherwise scattered across the issue tracker, code comments, and engineers' heads. When changing the encoding behavior, treat the tests in `VTest/Core/Shapes/` as ground truth and update this document in the same change.

## Background

`CustomPropertyCells.Value`, `CustomPropertyCells.Label`, `CustomPropertyCells.Format`, `CustomPropertyCells.Prompt`, `UserDefinedCellCells.Value`, and `UserDefinedCellCells.Prompt` are all stored as Visio **formulas**, not literal values. The cell write path is `cell.FormulaU = string`, where Visio parses the string as a ShapeSheet expression at write time.

Consequences:

- A bare identifier like `testVal` is parsed as a name reference.
- A bare numeric like `42` is parsed as a numeric literal.
- A quoted string like `"testVal"` is parsed as a string literal.
- A function call like `DATETIME("...")` is parsed as a function call.
- Invalid syntax raises `COMException` with message `#NAME?` (or other Visio error markers depending on the parse failure).

The library exposes `CustomPropertyCells.EncodeValues()` and `UserDefinedCellCells.EncodeValues()` to do the right thing for the common case (quote raw strings, leave already-formatted values alone). `Core.CellValue.EncodeValue` is the underlying helper:

```
- null or empty or starts with `=`        → returned unchanged
- starts with `"` and ends with `"`       → returned unchanged (already quoted)
- if quote=true                           → all `"` doubled, then surrounded by `"`
- if quote=false                          → returned unchanged
```

`EncodeValues()` calls `EncodeValue` with `quote=true` for the value field of string-typed `CustomPropertyCells`, and `quote=false` for non-string types (numeric, boolean, date — the values for those are already valid formulas).

## Encoding-aware paths in the codebase

Three callers go through `CustomPropertyHelper.Set` and pre-encode:

- `VisioScripting/Loaders/DirectedGraphDocumentLoader.cs:175` — calls `cp_cells.EncodeValues()` for each property parsed from `<directedgraph>` XML.
- `VisioScripting/Commands/CustomPropertyCommands.cs:106` — `Set-VisioCustomProperty` / `SetCustomProperty` calls `customprop.EncodeValues()`.

One internal caller does **not** pre-encode and assumes the caller already did:

- `VisioAutomation.Models/DOM/ShapeList.cs:96` — `SetCustomProperties` passes `kv.Value` through without encoding.

External callers using `CustomPropertyHelper.Set` directly are on their own.

## Behavior matrix

All cases below were captured against `CustomPropertyHelper.Set` on a fresh shape on a fresh page, reading back via `GetDictionary` in both `Formula` and `Result` modes. `[empty]` means an empty string. `[space]` means a single space character.

### `CustomPropertyCells`, Type=String (Type=0)

| Input (`cp.Value =`)       | Outcome                                                              |
|----------------------------|----------------------------------------------------------------------|
| `"testVal"` plain id       | **THROWS** `COMException` `#NAME?`                                   |
| `"42"` numeric-looking     | succeeds — formula=`42`, result=`42.0000` (numeric Result despite Type=String) |
| `"hello world"` spaces     | **THROWS** `COMException` `#NAME?`                                   |
| `""` empty unquoted        | succeeds — formula=`[empty]`, result=`0.0000`                        |
| `"\"\""` empty quoted      | round-trips — formula=`""`, result=`[empty]`                         |
| `null`                     | `HasValue=false`, value cell not written; default formula=`0`, result=`0.0000` |
| `"\"testVal\""` pre-quoted | round-trips — formula=`"testVal"`, result=`testVal`                  |
| `" "` single space unquoted| succeeds — formula=`[empty]`, result=`0.0000`                        |
| `"\" \""` single space quoted | round-trips — formula=`" "`, result=`[space]`                     |

Unencoded `Label`, `Format`, or `Prompt` fields with a plain-identifier value also throw `COMException` `#NAME?`, regardless of the `Type` setting.

The string-typed constructors `new CustomPropertyCells(string)` and `new CustomPropertyCells(string, CustomPropertyType.String)` propagate the unencoded value to `.Value`, so they hit the same throw on any non-numeric input.

### `CustomPropertyCells`, Type=Number (Type=2)

| Input                              | Outcome                                                              |
|------------------------------------|----------------------------------------------------------------------|
| `"42"`                             | succeeds — formula=`42`, result=`42.0000`                            |
| `"3.14"`                           | succeeds — formula=`3.14`, result=`3.1400`                           |
| `"testVal"` plain id               | **THROWS** `COMException` `#NAME?`                                   |
| `"\"42\""` quoted numeric          | succeeds — formula=`"42"`, result=`42`                               |
| `""` empty unquoted                | succeeds — formula=`[empty]`, result=`0.0000`                        |
| `null`                             | `HasValue=false`, cell unwritten; default formula=`0`, result=`0.0000` |

### `CustomPropertyCells`, Type=Boolean (Type=3)

| Input                              | Outcome                                                              |
|------------------------------------|----------------------------------------------------------------------|
| `true` literal bool                | succeeds — formula=`TRUE`, result=`TRUE`                             |
| `false` literal bool               | succeeds — formula=`FALSE`, result=`FALSE`                           |
| `"TRUE"` upper                     | succeeds — formula=`TRUE`, result=`TRUE`                             |
| `"FALSE"` upper                    | succeeds — formula=`FALSE`, result=`FALSE`                           |
| `"true"` lower                     | succeeds — Visio normalises to formula=`TRUE`, result=`TRUE`         |
| `"1"` numeric one                  | succeeds — formula=`1`, result=`1.0000` (NUMERIC Result despite Type=Boolean) |
| `"0"` numeric zero                 | succeeds — formula=`0`, result=`0.0000` (numeric Result despite Type=Boolean) |
| `"BAR"` plain id                   | **THROWS** `COMException` `#NAME?`                                   |
| `""` empty unquoted                | succeeds — formula=`[empty]`, result=`0.0000`                        |
| `null`                             | `HasValue=false`, cell unwritten; default formula=`0`, result=`0.0000` |

### `CustomPropertyCells`, Type=Date (Type=5)

| Input                                            | Outcome                                                              |
|--------------------------------------------------|----------------------------------------------------------------------|
| `DATETIME("03/31/2017 14:05:06")`                | succeeds, round-trips — result=`3/31/2017 2:05:06 PM` (locale-formatted) |
| `"testVal"` plain id                             | **THROWS** `COMException` `#NAME?`                                   |
| `"\"2017-03-31\""` pre-quoted ISO date           | succeeds AS A LITERAL STRING — formula=`"2017-03-31"`, result=`2017-03-31` (NOT parsed as a date) |
| `""` empty unquoted                              | succeeds — formula=`[empty]`, result=`0.0000`                        |
| `null`                                           | `HasValue=false`, cell unwritten; default formula=`0`, result=`0.0000` |

### `UserDefinedCellCells` (no `Type` concept; same shape as Type=String)

| Input (`udc.Value =`)              | Outcome                                                              |
|------------------------------------|----------------------------------------------------------------------|
| `"BAR"` plain id                   | **THROWS** `COMException` `#NAME?`                                   |
| `"42"`                             | succeeds — formula=`42`, result=`42.0000`                            |
| `"hello world"`                    | **THROWS** `COMException` `#NAME?`                                   |
| `""` empty unquoted                | succeeds — formula=`[empty]`, result=`0.0000`                        |
| `"\"\""` empty quoted              | round-trips — formula=`""`, result=`[empty]`                         |
| `null`                             | `HasValue=false`, cell unwritten; default formula=`0`, result=`0.0000` |
| `" "` single space unquoted        | succeeds — formula=`[empty]`, result=`0.0000`                        |
| `"\" \""` single space quoted      | round-trips — formula=`" "`, result=`[space]`                        |

`udc.Prompt` with an unencoded plain identifier likewise throws `COMException` `#NAME?`.

## Notable findings

1. **The headline trap is loud, not silent.** A common-case unencoded value like `cp.Value = "testVal"` raises `COMException: #NAME?` — users hit an exception, not silently-wrong data. The original framing of issue #144 ("silent foot-gun, substitutes 0") was incomplete; the silent paths are the edge cases below.

2. **Type metadata is advisory, not enforced.** Visio stores the Type cell as metadata but doesn't enforce that the Value matches. Examples:
    - Type=String + `"42"` → Result is the numeric `42.0000`, not the string `"42"`.
    - Type=Boolean + `"1"` → Result is the numeric `1.0000`, not `TRUE`.
    - Type=Date + `"\"2017-03-31\""` → Result is the literal string, not a parsed date.

   The type only constrains the editor UI, not the stored formula.

3. **Empty / null silently default to `0`.** The four "default-to-zero" failure modes (`null`, `""`, `" "`, missing-write) all produce a property whose Result mode reads `0.0000`. User reports of "my property reads as 0" almost certainly come from one of these paths, not from the plain-identifier path.

4. **Boolean lowercase is normalised.** `cp.Value = "true"` (lowercase) is rewritten by Visio to `TRUE` at write-time. Unique to Boolean — strings aren't normalised, dates aren't normalised.

5. **Quoted strings short-circuit Visio's formula parser.** Any input starting with `"` and ending with `"` round-trips as a literal, regardless of declared Type. This is what the library's `EncodeValue` relies on (its idempotence comes from this short-circuit).

6. **Result format is locale-dependent for dates.** `DATETIME(...)` Result mode produces `3/31/2017 2:05:06 PM` on a US-locale machine. Tests asserting this exact format will fail under other locales — flag for any future CI work.

## See also

- Characterization tests:
    - [`VTest/Core/Shapes/CustomPropertiesTest.cs`](../../VisioAutomation_2010/VTest/Core/Shapes/CustomPropertiesTest.cs) — `CustomProps_UnencodedValueCharacterization`, `CustomProps_NumberTypeCharacterization`, `CustomProps_BooleanTypeCharacterization`, `CustomProps_DateTypeCharacterization`.
    - [`VTest/Core/Shapes/UserDefinedCellsTests.cs`](../../VisioAutomation_2010/VTest/Core/Shapes/UserDefinedCellsTests.cs) — `UserDefinedCells_UnencodedValueCharacterization`.
- Encoding-aware code paths (search for callers of `EncodeValues()`):
    - [`VisioAutomation/Shapes/CustomPropertyCells.cs`](../../VisioAutomation_2010/VisioAutomation/Shapes/CustomPropertyCells.cs) — definition.
    - [`VisioAutomation/Shapes/UserDefinedCellCells.cs`](../../VisioAutomation_2010/VisioAutomation/Shapes/UserDefinedCellCells.cs) — definition.
    - [`VisioScripting/Loaders/DirectedGraphDocumentLoader.cs`](../../VisioAutomation_2010/VisioScripting/Loaders/DirectedGraphDocumentLoader.cs) — pre-encodes.
    - [`VisioScripting/Commands/CustomPropertyCommands.cs`](../../VisioAutomation_2010/VisioScripting/Commands/CustomPropertyCommands.cs) — pre-encodes.
- [Issue #144](https://github.com/saveenr/VisioAutomation/issues/144) — drives any future API ergonomics change.
- [Issue #117](https://github.com/saveenr/VisioAutomation/issues/117) — original user report that surfaced the encoding trap.
