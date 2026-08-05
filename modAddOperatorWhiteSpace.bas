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


' True if '-' follows an operator, '(', ',', or start of formula - i.e. a unary sign
' (e.g. "-1", "(-A1)"), not a binary subtraction operator (e.g. "A1-B1").
Private Function IsUnaryMinusContext(ByVal AStr As String) As Boolean
  Dim TmpCnt As Long
  Dim TmpCh  As String

  For TmpCnt = Len(AStr) To 1 Step -1
    TmpCh = Mid$(AStr, TmpCnt, 1)
    If TmpCh <> " " And TmpCh <> vbLf Then
      Select Case TmpCh
        Case ",", "(", "+", "-", "*", "/", "^", "&", "=", "<", ">"
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

Private Sub TestIsMultiOperator()
  Dim TmpMultiOps As Variant
  Dim TmpResult As Boolean

  ' Define multi-character operators
  TmpMultiOps = Array("<=", ">=", "<>")

  ' Test cases
  Debug.Print "Testing IsOperator (multi)..."
  Debug.Print "'<='   -> " & IsOperator("<=", TmpMultiOps) ' Expected: True
  Debug.Print "'>='   -> " & IsOperator(">=", TmpMultiOps) ' Expected: True
  Debug.Print "'<>'   -> " & IsOperator("<>", TmpMultiOps) ' Expected: True
  Debug.Print "'='    -> " & IsOperator("=", TmpMultiOps) ' Expected: False
  Debug.Print "'<X>'  -> " & IsOperator("<X>", TmpMultiOps) ' Expected: False
  Debug.Print "'>>'   -> " & IsOperator(">>", TmpMultiOps) ' Expected: False
  Debug.Print "'<=' (extra space) -> " & IsOperator("<= ", TmpMultiOps) ' Expected: False

  Debug.Print "TestIsMultiOperator Complete."
End Sub


Private Sub TestIsSingleOperator()
  Dim TmpSingleOps As Variant
  Dim TmpResult As Boolean

  ' Define single-character operators
  TmpSingleOps = Array("+", "-", "*", "/", "^", "&", "=", "<", ">")

  ' Test cases for all valid single-character operators
  Debug.Print "Testing IsOperator (single)..."
  Debug.Print "'+'  -> " & IsOperator("+", TmpSingleOps) ' Expected: True
  Debug.Print "'-'  -> " & IsOperator("-", TmpSingleOps) ' Expected: True
  Debug.Print "'*'  -> " & IsOperator("*", TmpSingleOps) ' Expected: True
  Debug.Print "'/'  -> " & IsOperator("/", TmpSingleOps) ' Expected: True
  Debug.Print "'^'  -> " & IsOperator("^", TmpSingleOps) ' Expected: True
  Debug.Print "'&'  -> " & IsOperator("&", TmpSingleOps) ' Expected: True
  Debug.Print "'='  -> " & IsOperator("=", TmpSingleOps) ' Expected: True
  Debug.Print "'<'  -> " & IsOperator("<", TmpSingleOps) ' Expected: True
  Debug.Print "'>'  -> " & IsOperator(">", TmpSingleOps) ' Expected: True

  ' Test case for a character that is NOT a single operator
  Debug.Print "'@'  -> " & IsOperator("@", TmpSingleOps) ' Expected: False

  Debug.Print "TestIsSingleOperator Complete."
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

Private Sub TestAddOperatorWhiteSpaceHelper(ByVal AFormula As String)
  Debug.Print "Before: "; AFormula
  Debug.Print "After:  "; AddOperatorWhiteSpace(AFormula)
End Sub

Private Sub TestAddOperatorWhiteSpace()
  Debug.Print "Testing AddOperatorWhiteSpace..."

  ' Simple formula with single-character operators
  TestAddOperatorWhiteSpaceHelper "=A1+B1*C1/D1^E1"

  ' Multi-character operators
  TestAddOperatorWhiteSpaceHelper "=A1<=B1>=C1<>D1"

  ' Brackets (Table reference)
  TestAddOperatorWhiteSpaceHelper "=SUM([Sales])"

  ' Nested Brackets (Table references inside functions)
  TestAddOperatorWhiteSpaceHelper "=INDEX(Table1[Column1], MATCH(A1, Table1[Column2], 0))"

  ' Single quotes for sheet names with spaces
  TestAddOperatorWhiteSpaceHelper "='Sheet Name'!A1+'Another Sheet'!B1"

  ' Mixed: Multi-ops, brackets, and single quotes
  TestAddOperatorWhiteSpaceHelper "='Data Sheet'!A1 + [Table1[Sales]] <= 'Summary'!B2"

  ' Brackets with operators
  TestAddOperatorWhiteSpaceHelper "=SUM([Revenue] - [Cost] / [Units])"

  ' Brackets with operators without spaces
  TestAddOperatorWhiteSpaceHelper "=SUM([Revenue]-[Cost]/[Units])"

  ' Brackets with operators without spaces
  TestAddOperatorWhiteSpaceHelper "=IF(AND([@[% In]],[@[Votes In]]),[@[Dem Votes]]/[@[Votes In]],0)"

  ' Miscellaneous case
  TestAddOperatorWhiteSpaceHelper "=IF(AND([@[% In]],[@[Votes In]]),[@[Total - 3rd]]/2+1, 0)"

  ' Multi-character operator followed by a space
  TestAddOperatorWhiteSpaceHelper "=IF(A1<= 100, ""Valid"", ""Invalid"")"

  ' Double bracket nesting (complex table references)
  TestAddOperatorWhiteSpaceHelper "=XLOOKUP(A1, Table1[[Lookup Col1]:[Lookup Col2]], Table1[Return Col])"

  ' Function inside single quotes with operators
  TestAddOperatorWhiteSpaceHelper "='Sales Report'!A1 & 'Sales Report'!B1"

  'Long.
  TestAddOperatorWhiteSpaceHelper "=IF(TRIM([@[Time Loc Index]])<>"""",INDEX(tblTimeLocations[UTC Offset],[@[Time Loc Index]],1)+IF(INDEX(tblTimeLocations[DST True/False],[@[Time Loc Index]],1),INDEX(tblTimeLocations[DST Now],[@[Time Loc Index]],1),FALSE),"")"

  TestAddOperatorWhiteSpaceHelper "=[@[Sales Total]]-[@[Expense Total]]"

  TestAddOperatorWhiteSpaceHelper "=LAMBDA(Input,Cnt,Pos,Incl,IFS((Cnt+Pos)<0,""[]"",(Cnt+Pos)>=8,""[]"",NOT(Incl),""[""&MID(Input,(Pos+Cnt),(ABS(Cnt)))&""]"",TRUE,""[""&MID(Input,(Pos+Cnt+1),(ABS(Cnt)))&""]""))($N70,$P70,$Q70,$O70)"

  ' Multiple sheet refs with & in name - second ref must not get spaces around &
  TestAddOperatorWhiteSpaceHelper "=IF(AND(Index!$K$2=""Forecast"",'Product Sales'!J$6>='FTM P&L'!$O$2),'Import Values from WB4'!J223,IF(AND(Index!$K$2=""Historical"",'Product Sales'!J$6<'FTM P&L'!$O$2),'Import Values from WB4'!J223,0))"

  ' Unary minus must not strip indentation or add left space
  TestAddOperatorWhiteSpaceHelper "=XLOOKUP(A1,B1:B10,C1:C10,0,-1)"

  ' -- coercion operator must not strip indentation in multi-line context
  TestAddOperatorWhiteSpaceHelper "=SUMPRODUCT(--(tblData[Status]=""Active""),--(tblData[Amount]>100),tblData[Value])"
  Debug.Print "TestAddOperatorWhiteSpace Complete."
End Sub


Private Sub TestAddOperatorWhiteSpaceEdgeCases()
  Debug.Print "Testing AddOperatorWhiteSpace Edge Cases..."

  ' Multi-operator followed by space
  TestAddOperatorWhiteSpaceHelper "=IF(A1<>"""", CONCAT(""Prefix-"", TEXTJOIN("", "", TRUE, B1, C1, D1)), ""No Data"")"

  ' Nested brackets with multi-operators
  TestAddOperatorWhiteSpaceHelper "=XLOOKUP(A1, Table1[[Lookup Col1]:[Lookup Col2]], Table1[Return Col])"

  ' Nested brackets inside an INDEX function
  TestAddOperatorWhiteSpaceHelper "=INDEX(Table1[[Column1]:[Column2]], MATCH(A1, Table1[Column3], 0))"

  ' Table reference with subtraction and division (original issue)
  TestAddOperatorWhiteSpaceHelper "=SUM([Revenue] - [Cost] / [Units])"

  ' Multi-character operator at the end of formula
  TestAddOperatorWhiteSpaceHelper "=IF(A1 <= 100, ""Valid"", ""Invalid"")"

  ' Text containing an Excel operator should remain unchanged
  TestAddOperatorWhiteSpaceHelper "=TEXTJOIN("", "", TRUE, ""A+B"", ""C&D"", ""X<>Y"")"

  ' Operator directly after brackets (ensuring no extra space is added)
  TestAddOperatorWhiteSpaceHelper "=SUM([Sales]+[Costs]/[Units])"

  ' Single quotes and table reference in same formula
  TestAddOperatorWhiteSpaceHelper "='Sales Report'!A1 & 'Sales Report'!B1"

  ' Function inside single quotes with multi-operators
  TestAddOperatorWhiteSpaceHelper "='Data Sheet'!A1 + [Table1[Sales]] <= 'Summary'!B2"

  TestAddOperatorWhiteSpaceHelper "='Data Sheet'!A1 + [Table1[Sales]]<='Summary'!B2"

  ' make sure -- stays together
  TestAddOperatorWhiteSpaceHelper "=--C1"

  ' Operator next to a bracket (ensuring no double spaces)
  TestAddOperatorWhiteSpaceHelper "=IF(A1<>"""", [Table1[Column1]]+[Table1[Column2]], """")"

  Debug.Print "TestAddOperatorWhiteSpaceEdgeCases Complete."
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


