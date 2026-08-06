Attribute VB_Name = "modAddOperatorWhiteSpace"
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

Private Sub TestIsOperatorHelper(ByVal AStr As String, ByVal AOps As Variant, ByVal AExpectedBool As Boolean, Optional ByVal ATestFilePathStr As String = vbNullString)
  Dim TmpAnswerBool As Boolean
  Dim TmpPassFailStr As String

  TmpAnswerBool = IsOperator(AStr, AOps)
  TmpPassFailStr = IIf(TmpAnswerBool = AExpectedBool, "PASS", "FAIL")

  TestLogLine AStr & " | " & TmpAnswerBool & " | " & AExpectedBool & " | " & TmpPassFailStr, ATestFilePathStr
End Sub

Private Sub TestIsMultiOperator(Optional ByVal ATestFilePathStr As String = vbNullString)
  Dim TmpMultiOps As Variant
  TmpMultiOps = Array("<=", ">=", "<>")

  TestIsOperatorHelper "<=", TmpMultiOps, True, ATestFilePathStr
  TestIsOperatorHelper ">=", TmpMultiOps, True, ATestFilePathStr
  TestIsOperatorHelper "<>", TmpMultiOps, True, ATestFilePathStr
  TestIsOperatorHelper "=", TmpMultiOps, False, ATestFilePathStr
  TestIsOperatorHelper "<X>", TmpMultiOps, False, ATestFilePathStr
  TestIsOperatorHelper ">>", TmpMultiOps, False, ATestFilePathStr
  TestIsOperatorHelper "<= ", TmpMultiOps, False, ATestFilePathStr 'Trailing space -- exact match required.
End Sub


Private Sub TestIsSingleOperator(Optional ByVal ATestFilePathStr As String = vbNullString)
  Dim TmpSingleOps As Variant
  TmpSingleOps = Array("+", "-", "*", "/", "^", "&", "=", "<", ">")

  TestIsOperatorHelper "+", TmpSingleOps, True, ATestFilePathStr
  TestIsOperatorHelper "-", TmpSingleOps, True, ATestFilePathStr
  TestIsOperatorHelper "*", TmpSingleOps, True, ATestFilePathStr
  TestIsOperatorHelper "/", TmpSingleOps, True, ATestFilePathStr
  TestIsOperatorHelper "^", TmpSingleOps, True, ATestFilePathStr
  TestIsOperatorHelper "&", TmpSingleOps, True, ATestFilePathStr
  TestIsOperatorHelper "=", TmpSingleOps, True, ATestFilePathStr
  TestIsOperatorHelper "<", TmpSingleOps, True, ATestFilePathStr
  TestIsOperatorHelper ">", TmpSingleOps, True, ATestFilePathStr
  TestIsOperatorHelper "@", TmpSingleOps, False, ATestFilePathStr 'Not a recognized operator character.
End Sub


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

Private Sub TestAddOperatorWhiteSpaceExactHelper(ByVal AFormula As String, ByVal AExpectedStr As String, Optional ByVal ATestFilePathStr As String = vbNullString)
  Dim TmpAnswerStr As String
  Dim TmpPassFailStr As String

  TmpAnswerStr = AddOperatorWhiteSpace(AFormula)
  TmpPassFailStr = IIf(TmpAnswerStr = AExpectedStr, "PASS", "FAIL")

  TestLogLine AFormula & " | " & TmpAnswerStr & " | " & AExpectedStr & " | " & TmpPassFailStr, ATestFilePathStr
End Sub

Private Sub TestAddOperatorWhiteSpaceExact(Optional ByVal ATestFilePathStr As String = vbNullString)
  ' Array literals only ever hold constants in Excel -- no operators except
  ' a leading "-" on a negative number literal. "{" wasn't recognized as a
  ' valid unary-minus context (unlike "(" and ","), so a negative number
  ' right after "{" used to get split into "{ - 1" instead of staying "{-1".
  TestAddOperatorWhiteSpaceExactHelper "={-1,2,-3}", "={-1,2,-3}", ATestFilePathStr
  TestAddOperatorWhiteSpaceExactHelper "={-1}", "={-1}", ATestFilePathStr
  TestAddOperatorWhiteSpaceExactHelper "={1,2,3}", "={1,2,3}", ATestFilePathStr

  ' Comma-space formatting (what PrettyPrint's own {} fix now produces) must
  ' not confuse the unary check -- IsUnaryMinusContext skips back past the
  ' space to find the comma, same as it always did outside of braces.
  TestAddOperatorWhiteSpaceExactHelper "={-1, 2, -3}", "={-1, 2, -3}", ATestFilePathStr

  ' A pre-existing space right after "{" (before this fix, the least likely
  ' real-world case, but the purest test of the whitespace-skip itself).
  ' The unary branch never trims -- unlike the binary-spacing branch, which
  ' calls TrimEnd -- so this space is preserved, not stripped.
  TestAddOperatorWhiteSpaceExactHelper "={ -1,2,-3}", "={ -1,2,-3}", ATestFilePathStr
End Sub

Private Sub TestAddOperatorWhiteSpace(Optional ByVal ATestFilePathStr As String = vbNullString)
  ' Simple formula with single-character operators
  TestAddOperatorWhiteSpaceExactHelper "=A1+B1*C1/D1^E1", "=A1 + B1 * C1 / D1 ^ E1", ATestFilePathStr

  ' Multi-character operators
  TestAddOperatorWhiteSpaceExactHelper "=A1<=B1>=C1<>D1", "=A1 <= B1 >= C1 <> D1", ATestFilePathStr

  ' Brackets (Table reference)
  TestAddOperatorWhiteSpaceExactHelper "=SUM([Sales])", "=SUM([Sales])", ATestFilePathStr

  ' Nested Brackets (Table references inside functions)
  TestAddOperatorWhiteSpaceExactHelper "=INDEX(Table1[Column1], MATCH(A1, Table1[Column2], 0))", "=INDEX(Table1[Column1], MATCH(A1, Table1[Column2], 0))", ATestFilePathStr

  ' Single quotes for sheet names with spaces
  TestAddOperatorWhiteSpaceExactHelper "='Sheet Name'!A1+'Another Sheet'!B1", "='Sheet Name'!A1 + 'Another Sheet'!B1", ATestFilePathStr

  ' Mixed: Multi-ops, brackets, and single quotes
  TestAddOperatorWhiteSpaceExactHelper "='Data Sheet'!A1 + [Table1[Sales]] <= 'Summary'!B2", "='Data Sheet'!A1 + [Table1[Sales]] <= 'Summary'!B2", ATestFilePathStr

  ' Brackets with operators
  TestAddOperatorWhiteSpaceExactHelper "=SUM([Revenue] - [Cost] / [Units])", "=SUM([Revenue] - [Cost] / [Units])", ATestFilePathStr

  ' Brackets with operators without spaces
  TestAddOperatorWhiteSpaceExactHelper "=SUM([Revenue]-[Cost]/[Units])", "=SUM([Revenue] - [Cost] / [Units])", ATestFilePathStr

  ' Brackets with operators without spaces
  TestAddOperatorWhiteSpaceExactHelper "=IF(AND([@[% In]],[@[Votes In]]),[@[Dem Votes]]/[@[Votes In]],0)", "=IF(AND([@[% In]],[@[Votes In]]),[@[Dem Votes]] / [@[Votes In]],0)", ATestFilePathStr

  ' Miscellaneous case
  TestAddOperatorWhiteSpaceExactHelper "=IF(AND([@[% In]],[@[Votes In]]),[@[Total - 3rd]]/2+1, 0)", "=IF(AND([@[% In]],[@[Votes In]]),[@[Total - 3rd]] / 2 + 1, 0)", ATestFilePathStr

  ' Multi-character operator followed by a space
  TestAddOperatorWhiteSpaceExactHelper "=IF(A1<= 100, ""Valid"", ""Invalid"")", "=IF(A1 <= 100, ""Valid"", ""Invalid"")", ATestFilePathStr

  ' Double bracket nesting (complex table references)
  TestAddOperatorWhiteSpaceExactHelper "=XLOOKUP(A1, Table1[[Lookup Col1]:[Lookup Col2]], Table1[Return Col])", "=XLOOKUP(A1, Table1[[Lookup Col1]:[Lookup Col2]], Table1[Return Col])", ATestFilePathStr

  ' Function inside single quotes with operators
  TestAddOperatorWhiteSpaceExactHelper "='Sales Report'!A1 & 'Sales Report'!B1", "='Sales Report'!A1 & 'Sales Report'!B1", ATestFilePathStr

  'Long.
  TestAddOperatorWhiteSpaceExactHelper "=IF(TRIM([@[Time Loc Index]])<>"""",INDEX(tblTimeLocations[UTC Offset],[@[Time Loc Index]],1)+IF(INDEX(tblTimeLocations[DST True/False],[@[Time Loc Index]],1),INDEX(tblTimeLocations[DST Now],[@[Time Loc Index]],1),FALSE),"")", _
    "=IF(TRIM([@[Time Loc Index]]) <> """",INDEX(tblTimeLocations[UTC Offset],[@[Time Loc Index]],1) + IF(INDEX(tblTimeLocations[DST True/False],[@[Time Loc Index]],1),INDEX(tblTimeLocations[DST Now],[@[Time Loc Index]],1),FALSE),"")", ATestFilePathStr

  TestAddOperatorWhiteSpaceExactHelper "=[@[Sales Total]]-[@[Expense Total]]", "=[@[Sales Total]] - [@[Expense Total]]", ATestFilePathStr

  TestAddOperatorWhiteSpaceExactHelper "=LAMBDA(Input,Cnt,Pos,Incl,IFS((Cnt+Pos)<0,""[]"",(Cnt+Pos)>=8,""[]"",NOT(Incl),""[""&MID(Input,(Pos+Cnt),(ABS(Cnt)))&""]"",TRUE,""[""&MID(Input,(Pos+Cnt+1),(ABS(Cnt)))&""]""))($N70,$P70,$Q70,$O70)", _
    "=LAMBDA(Input,Cnt,Pos,Incl,IFS((Cnt + Pos) < 0,""[]"",(Cnt + Pos) >= 8,""[]"",NOT(Incl),""["" & MID(Input,(Pos + Cnt),(ABS(Cnt))) & ""]"",TRUE,""["" & MID(Input,(Pos + Cnt + 1),(ABS(Cnt))) & ""]""))($N70,$P70,$Q70,$O70)", ATestFilePathStr

  ' Multiple sheet refs with & in name - second ref must not get spaces around &
  TestAddOperatorWhiteSpaceExactHelper "=IF(AND(Index!$K$2=""Forecast"",'Product Sales'!J$6>='FTM P&L'!$O$2),'Import Values from WB4'!J223,IF(AND(Index!$K$2=""Historical"",'Product Sales'!J$6<'FTM P&L'!$O$2),'Import Values from WB4'!J223,0))", _
    "=IF(AND(Index!$K$2 = ""Forecast"",'Product Sales'!J$6 >= 'FTM P&L'!$O$2),'Import Values from WB4'!J223,IF(AND(Index!$K$2 = ""Historical"",'Product Sales'!J$6 < 'FTM P&L'!$O$2),'Import Values from WB4'!J223,0))", ATestFilePathStr

  ' Unary minus must not strip indentation or add left space
  TestAddOperatorWhiteSpaceExactHelper "=XLOOKUP(A1,B1:B10,C1:C10,0,-1)", "=XLOOKUP(A1,B1:B10,C1:C10,0,-1)", ATestFilePathStr

  ' -- coercion operator must not strip indentation in multi-line context
  TestAddOperatorWhiteSpaceExactHelper "=SUMPRODUCT(--(tblData[Status]=""Active""),--(tblData[Amount]>100),tblData[Value])", "=SUMPRODUCT(--(tblData[Status] = ""Active""),--(tblData[Amount] > 100),tblData[Value])", ATestFilePathStr
End Sub


Private Sub TestAddOperatorWhiteSpaceEdgeCases(Optional ByVal ATestFilePathStr As String = vbNullString)
  ' Multi-operator followed by space
  TestAddOperatorWhiteSpaceExactHelper "=IF(A1<>"""", CONCAT(""Prefix-"", TEXTJOIN("", "", TRUE, B1, C1, D1)), ""No Data"")", _
    "=IF(A1 <> """", CONCAT(""Prefix-"", TEXTJOIN("", "", TRUE, B1, C1, D1)), ""No Data"")", ATestFilePathStr

  ' Nested brackets with multi-operators
  TestAddOperatorWhiteSpaceExactHelper "=XLOOKUP(A1, Table1[[Lookup Col1]:[Lookup Col2]], Table1[Return Col])", "=XLOOKUP(A1, Table1[[Lookup Col1]:[Lookup Col2]], Table1[Return Col])", ATestFilePathStr

  ' Nested brackets inside an INDEX function
  TestAddOperatorWhiteSpaceExactHelper "=INDEX(Table1[[Column1]:[Column2]], MATCH(A1, Table1[Column3], 0))", "=INDEX(Table1[[Column1]:[Column2]], MATCH(A1, Table1[Column3], 0))", ATestFilePathStr

  ' Table reference with subtraction and division (original issue)
  TestAddOperatorWhiteSpaceExactHelper "=SUM([Revenue] - [Cost] / [Units])", "=SUM([Revenue] - [Cost] / [Units])", ATestFilePathStr

  ' Multi-character operator at the end of formula
  TestAddOperatorWhiteSpaceExactHelper "=IF(A1 <= 100, ""Valid"", ""Invalid"")", "=IF(A1 <= 100, ""Valid"", ""Invalid"")", ATestFilePathStr

  ' Text containing an Excel operator should remain unchanged
  TestAddOperatorWhiteSpaceExactHelper "=TEXTJOIN("", "", TRUE, ""A+B"", ""C&D"", ""X<>Y"")", "=TEXTJOIN("", "", TRUE, ""A+B"", ""C&D"", ""X<>Y"")", ATestFilePathStr

  ' Operator directly after brackets (ensuring no extra space is added)
  TestAddOperatorWhiteSpaceExactHelper "=SUM([Sales]+[Costs]/[Units])", "=SUM([Sales] + [Costs] / [Units])", ATestFilePathStr

  ' Single quotes and table reference in same formula
  TestAddOperatorWhiteSpaceExactHelper "='Sales Report'!A1 & 'Sales Report'!B1", "='Sales Report'!A1 & 'Sales Report'!B1", ATestFilePathStr

  ' Function inside single quotes with multi-operators
  TestAddOperatorWhiteSpaceExactHelper "='Data Sheet'!A1 + [Table1[Sales]] <= 'Summary'!B2", "='Data Sheet'!A1 + [Table1[Sales]] <= 'Summary'!B2", ATestFilePathStr

  TestAddOperatorWhiteSpaceExactHelper "='Data Sheet'!A1 + [Table1[Sales]]<='Summary'!B2", "='Data Sheet'!A1 + [Table1[Sales]] <= 'Summary'!B2", ATestFilePathStr

  ' make sure -- stays together
  TestAddOperatorWhiteSpaceExactHelper "=--C1", "=--C1", ATestFilePathStr

  ' Operator next to a bracket (ensuring no double spaces)
  TestAddOperatorWhiteSpaceExactHelper "=IF(A1<>"""", [Table1[Column1]]+[Table1[Column2]], """")", _
    "=IF(A1 <> """", [Table1[Column1]] + [Table1[Column2]], """")", ATestFilePathStr
End Sub



' Helper function to compare expected vs. actual results.
' This is a very sweet test func look.
Private Sub TestAddOperatorWhiteSpaceDoubleNegHelper(ByVal AFrmStr As String)
  Dim TmpExpected As String
  Dim TmpResult As String
  Dim TmpMatch As Boolean

  ' Expected output should match input, except for proper spacing
  TmpExpected = AFrmStr
  TmpResult = AddOperatorWhiteSpace(AFrmStr)
  TmpMatch = (TmpResult = TmpExpected)

  Debug.Print "Before: "; AFrmStr, "After: "; TmpResult, "Expected: "; TmpExpected, "Match: "; TmpMatch
End Sub

'Many of these fail. These are pretty small edge cases. And it's pretty print and won't break the formula.
Private Sub TestAddOperatorWhiteSpaceDoubleNeg()
  Debug.Print "Testing AddOperatorWhiteSpace..."

  ' Simple double negative with no space
  TestAddOperatorWhiteSpaceDoubleNegHelper "=--C1"

  ' Double negative after addition
  TestAddOperatorWhiteSpaceDoubleNegHelper "=A1 + --B1"

  ' Double negative after multiplication
  TestAddOperatorWhiteSpaceDoubleNegHelper "=A1 * --B1"

  ' Double negative after division
  TestAddOperatorWhiteSpaceDoubleNegHelper "=A1 / --B1"

  ' Double negative after exponentiation
  TestAddOperatorWhiteSpaceDoubleNegHelper "=A1 ^ --B1"

  ' Double negative with parentheses
  TestAddOperatorWhiteSpaceDoubleNegHelper "=--(A1 + B1)"

  ' Edge case: Spacing between operators should be preserved
  TestAddOperatorWhiteSpaceDoubleNegHelper "=A1--B1" ' Should become `=A1 -- B1`

  ' Retain space before double negative inside IF
  TestAddOperatorWhiteSpaceDoubleNegHelper "=IF(A1 - --B1, TRUE, FALSE)"

  ' Check equality with double negative
  TestAddOperatorWhiteSpaceDoubleNegHelper "=IF(A1 = --B1, TRUE, FALSE)"

  Debug.Print "TestAddOperatorWhiteSpace Complete."
End Sub





