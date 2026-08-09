Attribute VB_Name = "modAddOperatorWhiteSpace"
' Copyright (c) 2026 Northbound Group
' Contact: stephendeetz@northboundgroup.com
' SPDX-License-Identifier: MIT
Option Explicit

Private Function TrimEnd(ByVal AStr As String) As String

  Dim TmpCnt As Long

  For TmpCnt = Len(AStr) To 1 Step -1
    If Mid$(AStr, TmpCnt, 1) <> " " Then
      Exit For
    End If
  Next TmpCnt

  TrimEnd = Left$(AStr, TmpCnt)
End Function


' True if '-' follows an operator, '(', ',', '{', or start of formula - i.e. a
' unary sign (e.g. "-1", "(-A1)", "{-1,2}"), not a binary subtraction
' operator (e.g. "A1-B1").
Private Function IsUnaryMinusContext(ByVal AStr As String) As Boolean
  Dim TmpCnt As Long
  Dim TmpCh  As String

  For TmpCnt = Len(AStr) To 1 Step -1
    TmpCh = Mid$(AStr, TmpCnt, 1)
    If TmpCh <> " " And TmpCh <> vbLf Then
      Select Case TmpCh
        Case ",", "(", "{", "+", "-", "*", "/", "^", "&", "=", "<", ">"
          IsUnaryMinusContext = True
        Case Else
          IsUnaryMinusContext = False
      End Select
      Exit Function
    End If
  Next TmpCnt

  IsUnaryMinusContext = True

End Function


Private Function IsOperator(ByVal AStr As String, _
                            ByVal AOps As Variant) As Boolean

  Dim TmpOp As Variant

  For Each TmpOp In AOps
    If AStr = TmpOp Then
      IsOperator = True
      Exit Function
    End If
  Next TmpOp

  IsOperator = False
End Function

Private Sub TestIsOperatorHelper(ByVal AStr As String, ByVal AOps As Variant, ByVal AExpectedBool As Boolean)
  Dim TmpAnswerBool As Boolean
  Dim TmpPassFailStr As String

  TmpAnswerBool = IsOperator(AStr, AOps)
  TmpPassFailStr = IIf(TmpAnswerBool = AExpectedBool, "PASS", "FAIL")

  TestLogLine AStr & " | " & TmpAnswerBool & " | " & AExpectedBool & " | " & TmpPassFailStr
End Sub

Private Sub TestIsMultiOperator()
  Dim TmpMultiOps As Variant
  TmpMultiOps = Array("<=", ">=", "<>")

  If IsStandaloneTestRun() Then ClearImmediateWindow

  TestIsOperatorHelper "<=", TmpMultiOps, True
  TestIsOperatorHelper ">=", TmpMultiOps, True
  TestIsOperatorHelper "<>", TmpMultiOps, True
  TestIsOperatorHelper "=", TmpMultiOps, False
  TestIsOperatorHelper "<X>", TmpMultiOps, False
  TestIsOperatorHelper ">>", TmpMultiOps, False
  TestIsOperatorHelper "<= ", TmpMultiOps, False 'Trailing space -- exact match required.
End Sub


Private Sub TestIsSingleOperator()
  Dim TmpSingleOps As Variant
  TmpSingleOps = Array("+", "-", "*", "/", "^", "&", "=", "<", ">")

  If IsStandaloneTestRun() Then ClearImmediateWindow

  TestIsOperatorHelper "+", TmpSingleOps, True
  TestIsOperatorHelper "-", TmpSingleOps, True
  TestIsOperatorHelper "*", TmpSingleOps, True
  TestIsOperatorHelper "/", TmpSingleOps, True
  TestIsOperatorHelper "^", TmpSingleOps, True
  TestIsOperatorHelper "&", TmpSingleOps, True
  TestIsOperatorHelper "=", TmpSingleOps, True
  TestIsOperatorHelper "<", TmpSingleOps, True
  TestIsOperatorHelper ">", TmpSingleOps, True
  TestIsOperatorHelper "@", TmpSingleOps, False 'Not a recognized operator character.
End Sub


' Purpose: Add spacing around binary operators in a formula string.
Public Function AddOperatorWhiteSpace(ByVal AStr As String) As String
  Dim TmpCnt               As Long
  Dim TmpStr               As String
  Dim TmpChar              As String
  Dim TmpNextChar          As String
  Dim TmpNextTwo           As String
  Dim TmpNextThree         As String

  Dim TmpInQuotes          As Boolean
  Dim TmpInSingleQuotes    As Boolean
  Dim TmpInBrackets        As Boolean
  Dim TmpBracketCnt        As Long

  Dim TmpMultiOps          As Variant
  Dim TmpSingleOps         As Variant

  ' Keep the leading "=" if present
  If Len(AStr) > 0 Then
    If Mid$(AStr, 1, 1) = "=" Then
      TmpStr = "="
      TmpCnt = 2  ' Start parsing from second character
    Else
      TmpCnt = 1
    End If
  End If

  ' Define operator lists
  TmpMultiOps = Array("<=", ">=", "<>", "--") '".:", ":.") ' ".:."
  TmpSingleOps = Array("+", "-", "*", "/", "^", "&", "=", "<", ">")

  ' Loop through formula
  Do While TmpCnt <= Len(AStr)
    TmpChar = Mid$(AStr, TmpCnt, 1)

    ' Look ahead
    If TmpCnt < Len(AStr) Then
      TmpNextChar = Mid$(AStr, TmpCnt + 1, 1)
      TmpNextTwo = Mid$(AStr, TmpCnt, 2)
      TmpNextThree = Mid$(AStr, TmpCnt, 3)
    Else
      TmpNextChar = ""
      TmpNextTwo = ""
      TmpNextThree = ""
    End If

    ' Inside quotes or brackets ? Add character as-is
    If TmpInQuotes Or TmpInSingleQuotes Or TmpInBrackets Then
      TmpStr = TmpStr & TmpChar

    Else
      ' Check for multi-character operators first
      If IsOperator(TmpNextTwo, TmpMultiOps) Then
        If TmpNextTwo = "--" Then
          If IsUnaryMinusContext(TmpStr) Then
            TmpStr = TmpStr & TmpNextTwo
          Else
            TmpStr = TrimEnd(TmpStr) & TmpNextTwo
          End If
        Else
          TmpStr = TrimEnd(TmpStr) & " " & TmpNextTwo
          If Right$(TmpNextThree, 1) <> " " Then TmpStr = TmpStr & " "
        End If

        TmpCnt = TmpCnt + 1 ' Move ahead to avoid skipping
        GoTo SkipIncrement
      End If

      ' Single-character operators
      If IsOperator(TmpChar, TmpSingleOps) Then
        ' **Only add spaces if NOT inside brackets**
        If Not TmpInBrackets Then
          If TmpChar = "-" And IsUnaryMinusContext(TmpStr) Then
            TmpStr = TmpStr & TmpChar
          Else
            TmpStr = TrimEnd(TmpStr) & " " & TmpChar
            If TmpNextChar <> " " Then TmpStr = TmpStr & " "
          End If
        Else
          TmpStr = TmpStr & TmpChar
        End If

      Else
        TmpStr = TmpStr & TmpChar
      End If
    End If

    ' Toggle quotes & bracket tracking
    If TmpChar = """" Then
      TmpInQuotes = Not TmpInQuotes
    End If

    If Not TmpInQuotes Then
      If TmpChar = "'" Then
        If TmpNextChar = "!" Then
          TmpStr = TmpStr & "!"
          TmpCnt = TmpCnt + 1
          TmpInSingleQuotes = False
        Else
          TmpInSingleQuotes = Not TmpInSingleQuotes
        End If
      ElseIf TmpChar = "[" Then
        TmpBracketCnt = TmpBracketCnt + 1
        TmpInBrackets = True
      ElseIf TmpChar = "]" Then
        TmpBracketCnt = TmpBracketCnt - 1
        TmpInBrackets = (TmpBracketCnt > 0)
      End If
    End If

SkipIncrement:
    TmpCnt = TmpCnt + 1
  Loop

  AddOperatorWhiteSpace = TmpStr
End Function

Private Sub TestAddOperatorWhiteSpaceExactHelper(ByVal AFormula As String, ByVal AExpectedStr As String)
  Dim TmpAnswerStr As String
  Dim TmpPassFailStr As String

  TmpAnswerStr = AddOperatorWhiteSpace(AFormula)
  TmpPassFailStr = IIf(TmpAnswerStr = AExpectedStr, "PASS", "FAIL")

  TestLogLine AFormula & " | " & TmpAnswerStr & " | " & AExpectedStr & " | " & TmpPassFailStr
End Sub

Private Sub TestAddOperatorWhiteSpaceExact()
  If IsStandaloneTestRun() Then ClearImmediateWindow

  ' Array literals only ever hold constants in Excel -- no operators except
  ' a leading "-" on a negative number literal. "{" wasn't recognized as a
  ' valid unary-minus context (unlike "(" and ","), so a negative number
  ' right after "{" used to get split into "{ - 1" instead of staying "{-1".
  TestAddOperatorWhiteSpaceExactHelper "={-1,2,-3}", "={-1,2,-3}"
  TestAddOperatorWhiteSpaceExactHelper "={-1}", "={-1}"
  TestAddOperatorWhiteSpaceExactHelper "={1,2,3}", "={1,2,3}"

  ' Comma-space formatting (what PrettyPrint's own {} fix now produces) must
  ' not confuse the unary check -- IsUnaryMinusContext skips back past the
  ' space to find the comma, same as it always did outside of braces.
  TestAddOperatorWhiteSpaceExactHelper "={-1, 2, -3}", "={-1, 2, -3}"

  ' A pre-existing space right after "{" (before this fix, the least likely
  ' real-world case, but the purest test of the whitespace-skip itself).
  ' The unary branch never trims -- unlike the binary-spacing branch, which
  ' calls TrimEnd -- so this space is preserved, not stripped.
  TestAddOperatorWhiteSpaceExactHelper "={ -1,2,-3}", "={ -1,2,-3}"
End Sub

Private Sub TestAddOperatorWhiteSpace()
  If IsStandaloneTestRun() Then ClearImmediateWindow

  ' Simple formula with single-character operators
  TestAddOperatorWhiteSpaceExactHelper "=A1+B1*C1/D1^E1", "=A1 + B1 * C1 / D1 ^ E1"

  ' Multi-character operators
  TestAddOperatorWhiteSpaceExactHelper "=A1<=B1>=C1<>D1", "=A1 <= B1 >= C1 <> D1"

  ' Brackets (Table reference)
  TestAddOperatorWhiteSpaceExactHelper "=SUM([Sales])", "=SUM([Sales])"

  ' Nested Brackets (Table references inside functions)
  TestAddOperatorWhiteSpaceExactHelper "=INDEX(Table1[Column1], MATCH(A1, Table1[Column2], 0))", "=INDEX(Table1[Column1], MATCH(A1, Table1[Column2], 0))"

  ' Single quotes for sheet names with spaces
  TestAddOperatorWhiteSpaceExactHelper "='Sheet Name'!A1+'Another Sheet'!B1", "='Sheet Name'!A1 + 'Another Sheet'!B1"

  ' Mixed: Multi-ops, brackets, and single quotes
  TestAddOperatorWhiteSpaceExactHelper "='Data Sheet'!A1 + [Table1[Sales]] <= 'Summary'!B2", "='Data Sheet'!A1 + [Table1[Sales]] <= 'Summary'!B2"

  ' Brackets with operators
  TestAddOperatorWhiteSpaceExactHelper "=SUM([Revenue] - [Cost] / [Units])", "=SUM([Revenue] - [Cost] / [Units])"

  ' Brackets with operators without spaces
  TestAddOperatorWhiteSpaceExactHelper "=SUM([Revenue]-[Cost]/[Units])", "=SUM([Revenue] - [Cost] / [Units])"

  ' Brackets with operators without spaces
  TestAddOperatorWhiteSpaceExactHelper "=IF(AND([@[% In]],[@[Votes In]]),[@[Dem Votes]]/[@[Votes In]],0)", "=IF(AND([@[% In]],[@[Votes In]]),[@[Dem Votes]] / [@[Votes In]],0)"

  ' Miscellaneous case
  TestAddOperatorWhiteSpaceExactHelper "=IF(AND([@[% In]],[@[Votes In]]),[@[Total - 3rd]]/2+1, 0)", "=IF(AND([@[% In]],[@[Votes In]]),[@[Total - 3rd]] / 2 + 1, 0)"

  ' Multi-character operator followed by a space
  TestAddOperatorWhiteSpaceExactHelper "=IF(A1<= 100, ""Valid"", ""Invalid"")", "=IF(A1 <= 100, ""Valid"", ""Invalid"")"

  ' Double bracket nesting (complex table references)
  TestAddOperatorWhiteSpaceExactHelper "=XLOOKUP(A1, Table1[[Lookup Col1]:[Lookup Col2]], Table1[Return Col])", "=XLOOKUP(A1, Table1[[Lookup Col1]:[Lookup Col2]], Table1[Return Col])"

  ' Function inside single quotes with operators
  TestAddOperatorWhiteSpaceExactHelper "='Sales Report'!A1 & 'Sales Report'!B1", "='Sales Report'!A1 & 'Sales Report'!B1"

  'Long.
  TestAddOperatorWhiteSpaceExactHelper "=IF(TRIM([@[Time Loc Index]])<>"""",INDEX(tblTimeLocations[UTC Offset],[@[Time Loc Index]],1)+IF(INDEX(tblTimeLocations[DST True/False],[@[Time Loc Index]],1),INDEX(tblTimeLocations[DST Now],[@[Time Loc Index]],1),FALSE),"")", _
    "=IF(TRIM([@[Time Loc Index]]) <> """",INDEX(tblTimeLocations[UTC Offset],[@[Time Loc Index]],1) + IF(INDEX(tblTimeLocations[DST True/False],[@[Time Loc Index]],1),INDEX(tblTimeLocations[DST Now],[@[Time Loc Index]],1),FALSE),"")"

  TestAddOperatorWhiteSpaceExactHelper "=[@[Sales Total]]-[@[Expense Total]]", "=[@[Sales Total]] - [@[Expense Total]]"

  TestAddOperatorWhiteSpaceExactHelper "=LAMBDA(Input,Cnt,Pos,Incl,IFS((Cnt+Pos)<0,""[]"",(Cnt+Pos)>=8,""[]"",NOT(Incl),""[""&MID(Input,(Pos+Cnt),(ABS(Cnt)))&""]"",TRUE,""[""&MID(Input,(Pos+Cnt+1),(ABS(Cnt)))&""]""))($N70,$P70,$Q70,$O70)", _
    "=LAMBDA(Input,Cnt,Pos,Incl,IFS((Cnt + Pos) < 0,""[]"",(Cnt + Pos) >= 8,""[]"",NOT(Incl),""["" & MID(Input,(Pos + Cnt),(ABS(Cnt))) & ""]"",TRUE,""["" & MID(Input,(Pos + Cnt + 1),(ABS(Cnt))) & ""]""))($N70,$P70,$Q70,$O70)"

  ' Multiple sheet refs with & in name - second ref must not get spaces around &
  TestAddOperatorWhiteSpaceExactHelper "=IF(AND(Index!$K$2=""Forecast"",'Product Sales'!J$6>='FTM P&L'!$O$2),'Import Values from WB4'!J223,IF(AND(Index!$K$2=""Historical"",'Product Sales'!J$6<'FTM P&L'!$O$2),'Import Values from WB4'!J223,0))", _
    "=IF(AND(Index!$K$2 = ""Forecast"",'Product Sales'!J$6 >= 'FTM P&L'!$O$2),'Import Values from WB4'!J223,IF(AND(Index!$K$2 = ""Historical"",'Product Sales'!J$6 < 'FTM P&L'!$O$2),'Import Values from WB4'!J223,0))"

  ' Unary minus must not strip indentation or add left space
  TestAddOperatorWhiteSpaceExactHelper "=XLOOKUP(A1,B1:B10,C1:C10,0,-1)", "=XLOOKUP(A1,B1:B10,C1:C10,0,-1)"

  ' -- coercion operator must not strip indentation in multi-line context
  TestAddOperatorWhiteSpaceExactHelper "=SUMPRODUCT(--(tblData[Status]=""Active""),--(tblData[Amount]>100),tblData[Value])", "=SUMPRODUCT(--(tblData[Status] = ""Active""),--(tblData[Amount] > 100),tblData[Value])"
End Sub


Private Sub TestAddOperatorWhiteSpaceEdgeCases()
  If IsStandaloneTestRun() Then ClearImmediateWindow

  ' Multi-operator followed by space
  TestAddOperatorWhiteSpaceExactHelper "=IF(A1<>"""", CONCAT(""Prefix-"", TEXTJOIN("", "", TRUE, B1, C1, D1)), ""No Data"")", _
    "=IF(A1 <> """", CONCAT(""Prefix-"", TEXTJOIN("", "", TRUE, B1, C1, D1)), ""No Data"")"

  ' Nested brackets with multi-operators
  TestAddOperatorWhiteSpaceExactHelper "=XLOOKUP(A1, Table1[[Lookup Col1]:[Lookup Col2]], Table1[Return Col])", "=XLOOKUP(A1, Table1[[Lookup Col1]:[Lookup Col2]], Table1[Return Col])"

  ' Nested brackets inside an INDEX function
  TestAddOperatorWhiteSpaceExactHelper "=INDEX(Table1[[Column1]:[Column2]], MATCH(A1, Table1[Column3], 0))", "=INDEX(Table1[[Column1]:[Column2]], MATCH(A1, Table1[Column3], 0))"

  ' Table reference with subtraction and division (original issue)
  TestAddOperatorWhiteSpaceExactHelper "=SUM([Revenue] - [Cost] / [Units])", "=SUM([Revenue] - [Cost] / [Units])"

  ' Multi-character operator at the end of formula
  TestAddOperatorWhiteSpaceExactHelper "=IF(A1 <= 100, ""Valid"", ""Invalid"")", "=IF(A1 <= 100, ""Valid"", ""Invalid"")"

  ' Text containing an Excel operator should remain unchanged
  TestAddOperatorWhiteSpaceExactHelper "=TEXTJOIN("", "", TRUE, ""A+B"", ""C&D"", ""X<>Y"")", "=TEXTJOIN("", "", TRUE, ""A+B"", ""C&D"", ""X<>Y"")"

  ' Operator directly after brackets (ensuring no extra space is added)
  TestAddOperatorWhiteSpaceExactHelper "=SUM([Sales]+[Costs]/[Units])", "=SUM([Sales] + [Costs] / [Units])"

  ' Single quotes and table reference in same formula
  TestAddOperatorWhiteSpaceExactHelper "='Sales Report'!A1 & 'Sales Report'!B1", "='Sales Report'!A1 & 'Sales Report'!B1"

  ' Function inside single quotes with multi-operators
  TestAddOperatorWhiteSpaceExactHelper "='Data Sheet'!A1 + [Table1[Sales]] <= 'Summary'!B2", "='Data Sheet'!A1 + [Table1[Sales]] <= 'Summary'!B2"

  TestAddOperatorWhiteSpaceExactHelper "='Data Sheet'!A1 + [Table1[Sales]]<='Summary'!B2", "='Data Sheet'!A1 + [Table1[Sales]] <= 'Summary'!B2"

  ' make sure -- stays together
  TestAddOperatorWhiteSpaceExactHelper "=--C1", "=--C1"

  ' Operator next to a bracket (ensuring no double spaces)
  TestAddOperatorWhiteSpaceExactHelper "=IF(A1<>"""", [Table1[Column1]]+[Table1[Column2]], """")", _
    "=IF(A1 <> """", [Table1[Column1]] + [Table1[Column2]], """")"
End Sub



' "--" (double unary minus / boolean-to-number coercion, e.g. SUMPRODUCT's
' --(condition)) never gets surrounding spaces added, in either the unary
' ("=--C1") or binary/coercion ("=A1--B1") context -- the special-cased
' branch for it in AddOperatorWhiteSpace only conditionally trims a
' preceding space, unlike every other operator. Confirmed by running this
' test standalone: all cases below are the actual, unchanged output.
Private Sub TestAddOperatorWhiteSpaceDoubleNeg()
  If IsStandaloneTestRun() Then ClearImmediateWindow
  TestAddOperatorWhiteSpaceExactHelper "=--C1", "=--C1"
  TestAddOperatorWhiteSpaceExactHelper "=A1+--B1", "=A1 + --B1"
  TestAddOperatorWhiteSpaceExactHelper "=A1 * --B1", "=A1 * --B1"
  TestAddOperatorWhiteSpaceExactHelper "=A1 / --B1", "=A1 / --B1"
  TestAddOperatorWhiteSpaceExactHelper "=A1 ^ --B1", "=A1 ^ --B1"
  TestAddOperatorWhiteSpaceExactHelper "=--(A1 + B1)", "=--(A1 + B1)"
  TestAddOperatorWhiteSpaceExactHelper "=A1--B1", "=A1--B1"
  TestAddOperatorWhiteSpaceExactHelper "=IF(A1 - --B1, TRUE, FALSE)", "=IF(A1 - --B1, TRUE, FALSE)"
  TestAddOperatorWhiteSpaceExactHelper "=IF(A1 = --B1, TRUE, FALSE)", "=IF(A1 = --B1, TRUE, FALSE)"
End Sub








