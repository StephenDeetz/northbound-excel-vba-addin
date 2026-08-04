# VBA Code Style Guide

Version 1.

---

## Code Abbreviations

| Abbreviation | Meaning |
|---|---|
| Abbr | Abbreviation |
| Addr | Address |
| Arg | Argument |
| Char | Character |
| Cnt | Counting variable |
| Col | Column |
| Coll | Collection |
| Dst | Destination |
| Frm | An Excel formula (not VBA) |
| Func | Function |
| Hdr | Header |
| Idx | Index |
| Lbl | Label |
| Ref | Reference |
| Rng | Range |
| Sel | Selection |
| Sht | Worksheet |
| Src | Source |
| Sub | Subroutine |
| Tgt | Target |
| Var | Variant |
| Wbk | Workbook |

---

## Naming Conventions

### Prefixes

| Use | Prefix |
|---|---|
| Local variable | `Tmp` |
| Argument | `A` |
| Constant | `k` |

### Postfixes

| Type | Postfix |
|---|---|
| String variable or function returning a string | `Str` |
| Collection | `Coll` |
| Variant | `Var` |
| Loop count variable | `Cnt` |

### Capitalization

- Capitalized Camel Case for variables and functions.
- No global variable prefix.

---

## VBA Coding Preferences

- Functions/Subs: declare visibility explicitly. Default: `Private`.
- Arguments: declare `ByVal`/`ByRef` explicitly. Default: `ByVal`.
- Arguments: list ByRef args last.
- Use `Workbook.Worksheets` rather than `Workbook.Sheets`.
- Default workbook: `ActiveWorkbook`, not `ThisWorkbook`.
- If a function can return an exception or error, use the `AErrStr` pattern (see Error Handling).
- Avoid VBA Forms.
- Minimize extra variables.
- Don't use `Call` to invoke subs. Use `SubName Arg1, Arg2`.
- Append `$` on `Mid`, `Left`, and `Right` (i.e. `Mid$`, `Left$`, `Right$`).
- When an `On Error` handler is in effect, use `Err.Clear` after handling.
- Smaller functions topmost. Order pattern: Small Functions, Cell, Sel, Sht, Wbk.

---

## Function Rules

### Naming Pattern

`Obj-Verb-Noun` — where `Obj` is the Excel object, `Verb` is the action, and `Noun` is the optional descriptor.

Example: `WbkHideRows`.

### Default Scope

Private.

### Function Body Order

1. Local variable declarations
2. Default function result
3. Exit conditions
4. `On Error` (if used)
5. Code
6. `On Error` label (if used) — always named `OnError:`
7. Cleanup (if needed)

### Indentation

Indent by 2 spaces. No tab characters.

### Variable Declaration Spacing

Single space in declarations. Use:

```vba
Dim TmpCnt As Long
```

Not:

```vba
Dim TmpCnt               As Long
```

### Type Choices

- Use `Long` rather than `Integer` for counts and any numeric work. VBA's `Integer` is a 16-bit type and offers no benefit on modern systems.

### Counting Variable Patterns

- Single count in a function: `TmpCnt`.
- Multiple counts: `Tmp` + variable type + `Cnt`. Example: `TmpColCnt`, `TmpRowCnt`.

### Loop Skip Labels

Use `GoTo Next<IterVar>` with a label immediately before the `Next` statement to simulate `continue`.

- Label name = `Next` + iterator variable name. E.g., `TmpCnt` -> `NextTmpCnt`; `TmpRowCnt` -> `NextTmpRowCnt`.
- Label placed inside the loop, on the line directly above `Next <IterVar>`.
- Use to skip remaining body when a guard condition fails (empty cell, invalid value, etc.).
- `GoTo` permitted only for `OnError:` handlers and these loop-skip labels — no other use.
- Nested loops: each loop gets its own `Next<IterVar>` label; jump only to the label of the loop being continued.

```vba
Dim TmpRowCnt As Long
For TmpRowCnt = 1 To 100
  If Cells(TmpRowCnt, 1).Value = "" Then GoTo NextTmpRowCnt

  ' ...processing...

NextTmpRowCnt:
Next TmpRowCnt
```

### Function Calls Without Result

Don't use parentheses when calling a function whose result you discard.

---

## Argument Style

- Keep the first argument on the same line as the function name.
- Put each additional argument on its own line.
- Use the `_` line-continuation character.
- Left-align all arguments to the first argument.
- Put the return type on the same line as the last argument.
- Vertically align argument types where natural.

Example:

```vba
Public Sub DoSomething(ByVal AFirstArg  As Long, _
                       ByVal ASecondArg As String, _
                       ByRef AErrStr    As String)
```

### First Argument by Function Type

| Function type | First argument |
|---|---|
| Cell functions | `ByVal ACell As Range` |
| Range functions | `ByVal ARng As Range` |
| Sheet functions | `ByVal ASht As Worksheet` |
| Workbook functions | `ByVal AWbk As Workbook` |

---

## Error Handling

### ErrStr Pattern

- Last argument to the function: `ByRef AErrStr As String`.
- `AErrStr` is used to pass an error message back from a trapped error.
- `AResultStr` is used to pass back general information.
- Used in conjunction with a return-value pattern. Example: a Boolean return — if `False`, the caller does `MsgBox AErrStr`.
- The trapping (outer) function defines `TmpErrStr`. Inner functions receive it as `AErrStr`.

### Error Handler Label

- Always name the error handler label `OnError:`.

---

## AllClear Pattern

`AllClear` is the standard first-exit guard for any function operating on `ActiveWorkbook`, `ActiveSheet`, or `Selection`.

```vba
Public Function AllClear(ByVal AAnyProtectedShtWarning      As Boolean, _
                         ByVal AProtectedActiveShtWarning   As Boolean) As Boolean
```

- Use `AllClear` as the first exit condition for `ActiveWbk`, `ActiveSht`, and `ActiveSel` functions.
- Sub-functions (SFs) do not call `AllClear`.

---

## Excel Object Naming

When referring to public Excel objects inside function names or local references:

| Excel object | Partial name |
|---|---|
| `ActiveCell` | `ActiveCell` |
| `ActiveSheet` | `ActiveSht` |
| `ActiveWorkbook` | `ActiveWbk` |
| `Selection` | `ActiveSel` |

---

## Test Functions

- For each sub-function `SF`, write a test sub named `Test` + the SF name (e.g. `TestPrettyPrintRange`).
- Test sub signature: `Private Sub TestSF()` — no arguments.
- Test functions use `Debug.Print`. Reserve `MsgBox` for output involving `ActiveWorkbook`, `ActiveSheet`, or `Selection`.
- Test functions do not reference `ThisWorkbook`. Use `ActiveWorkbook`.

---

## Variable Naming Patterns

### Simple

Single variable of the needed type. E.g. a single integer counting loop: `TmpCnt`.

### Complex

Multiple variables of one type. E.g. counting successes and errors: `TmpSuccCnt`, `TmpErrCnt`.

### Variant Iteration

When itemizing a `Collection`, use `TmpItem` typed as `Variant`.

### Local-to-Argument Mapping

Local variable `TmpVar` is passed into a called function as `AVar`.

Per the Code Style, `Tmp` becomes `A` when crossing a function boundary.

---

## Comments

Every `Public Sub` and `Public Function` has a one-line comment immediately above its declaration describing its purpose:

```vba
' Purpose: Toggle gridlines for the entire active workbook.
Public Sub WbkGridlinesToggle()
```

This is checkable mechanically: any `Public Sub`/`Public Function` not preceded by a comment line is a violation.

---

## ChatGPT / AI Preferences

- When pasting code, please put each function in a separate code window.

---

## Patterns Summary

- **WSRC** — Workbook, Sheet, Range, Cell

These describe the typical drill-down structure for functions that operate at multiple Excel scopes.

E.g. WSRC behaviors have a beginning and ending level. Workbooks loop thru Sheets, Sheets thru Ranges, and Ranges sometimes loop thru Individual cells.

E.g. Gridlines for whole workbook has a beginning level of Workbook and ending level Sheet. The loop is thru `Workbook.Worksheets`.

E.g. Pretty Print Workbook - has beginning level of Workbook and ending level of Cell as each cell must be corrected.