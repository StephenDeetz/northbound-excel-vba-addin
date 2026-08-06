Attribute VB_Name = "modPrettyPrint"
Option Explicit

Private Const kSimpleArgTextLen = 50
Private Const kSimpleArgCnt = 4

Private Const kLogFilePathStr As String = "C:\temp\temp.txt"


'----------------------------------------------------------------------'
'                       Test Log Functions                             '
'----------------------------------------------------------------------'

Private Sub LogClear()
  If Dir$(kLogFilePathStr) <> vbNullString Then Kill kLogFilePathStr
End Sub


Private Sub LogWrite(ByVal AStr As String)
  Dim TmpFileNbr As Integer

  TmpFileNbr = FreeFile
  Open kLogFilePathStr For Append As #TmpFileNbr
    Print #TmpFileNbr, AStr
  Close #TmpFileNbr
End Sub


'----------------------------------------------------------------------'
'                       Message Functions                              '
'----------------------------------------------------------------------'

Private Function ResultStrMinify(ByVal AResult As Boolean, _
                                 ByVal ASuccCnt As Long, _
                                 ByVal AErrCnt As Long) As String
  If AResult Then
    ResultStrMinify = "Finished. " & Format(ASuccCnt, "#,##0") & " Formula(s) Minified."
  Else
    ResultStrMinify = "Error. " & vbCrLf & _
                      Format(ASuccCnt, "#,##0") & " Formula(s) Minified." & vbCrLf & _
                      Format(AErrCnt, "#,##0") & " Formula(s) Unable to Change."
  End If

End Function

Private Function ResultStrPrettyPrint(ByVal AResult As Boolean, _
                                      ByVal ASuccCnt As Long, _
                                      ByVal AErrCnt As Long) As String
    
  If AResult Then
    ResultStrPrettyPrint = "Finished. " & Format(ASuccCnt, "#,##0") & " Formula(s) Pretty Printed."
  Else
    ResultStrPrettyPrint = "Error. " & vbCrLf & _
                           Format(ASuccCnt, "#,##0") & " Formula(s) Pretty Printed." & vbCrLf & _
                           Format(AErrCnt, "#,##0") & " Formula(s) Unable to Change."
  End If

End Function


'----------------------------------------------------------------------'
'                        Core Functions                                '
'----------------------------------------------------------------------'

'Note - Minify just calls ShortenFormula SF directly.

'1st Char (
Private Function FindMatchingCloseParen(ByVal AOpenParenStart As Long, ByVal AStr As String) As Long
 Dim TmpInsideQuotes As Boolean
  Dim i As Long
  Dim TmpMatchCnt As Long '( goes up, ) down. When 0, found closing paren.
  
  If Mid$(AStr, AOpenParenStart, 1) = "(" Then
    TmpMatchCnt = 1
  Else
    'Debug.Print "Error: FindMatchingCloseParen - No Starting Open Paren."
    Exit Function
  End If
  
  FindMatchingCloseParen = 0

  TmpInsideQuotes = False
  
  For i = AOpenParenStart + 1 To Len(AStr)
    If Mid(AStr, i, 1) = """" Then
      TmpInsideQuotes = Not TmpInsideQuotes
    ElseIf Not TmpInsideQuotes Then
      If Mid$(AStr, i, 1) = "(" Then
        Inc TmpMatchCnt
      End If
    
      If Mid$(AStr, i, 1) = ")" Then
        Dec TmpMatchCnt
        
        If TmpMatchCnt = 0 Then
          FindMatchingCloseParen = i
          Exit Function
        End If
      End If
    End If
  Next i
  
End Function

Private Sub TestFindMatchingCloseParen()
  Debug.Print FindMatchingCloseParen(5, "=@IF(TRIM($A2)<>"""",IF(IFERROR(FIND("" ) "",$A2),0)>0,INDEX(TEXTSPLIT($A2,""("","" - ""),1,2),INDEX(TEXTSPLIT($A2,""|""),2)),"""")")
  Debug.Print FindMatchingCloseParen(10, "=@IF(TRIM($A2)<>"""",IF(IFERROR(FIND("" ) "",$A2),0)>0,INDEX(TEXTSPLIT($A2,""("","" - ""),1,2),INDEX(TEXTSPLIT($A2,""|""),2)),"""")")
  Debug.Print FindMatchingCloseParen(22, "=@IF(TRIM($A2)<>"""",IF(IFERROR(FIND("" ) "",$A2),0)>0,INDEX(TEXTSPLIT($A2,""("","" - ""),1,2),INDEX(TEXTSPLIT($A2,""|""),2)),"""")")
End Sub


' ------------------ SimpleTextArg Documentation ---------------------'
' SimpleTextArg is run after each (.                                  '
' SimpleTextArg attempts to find ( that don't need a new line.        '
'                                                                     '
' Rules for Not Needing New Line:                                     '
' --------------------------------------------------------------      '
' 1. Single Arg, No Nested Parens or Func Calls                       '
' 2. No Nested Parens or Func Calls and kSimpleArgCnt args or less.   '
' 3. Only 1 function call and no longer than kSimpleArgTextLen chars. '
' --------------------------------------------------------------------'

Private Function SimpleArgText(ByVal AFrmStr As String, ByVal APos As Long) As String
  Dim TmpPosCloseParen As Long
  Dim TmpPosOpenParen As Long
  Dim TmpPosComma As Long
  Dim TmpIsSimpleArg As Boolean
  Dim TmpUntilCloseStr As String
  Dim TmpCommaCnt As Long
  Dim TmpOpenCnt As Long

  SimpleArgText = ""
  
  ' Find positions of key characters
  TmpPosCloseParen = FindMatchingCloseParen(APos, AFrmStr)
  TmpPosOpenParen = PosSkipQuotedSections(APos + 1, AFrmStr, "(")
  TmpPosComma = PosSkipQuotedSections(APos + 1, AFrmStr, ",")
  
  'Encounter a ) before ( or ,?
  TmpIsSimpleArg = (TmpPosCloseParen > 0) _
                   And (TmpPosOpenParen = 0 Or TmpPosOpenParen > TmpPosCloseParen) _
                   And (TmpPosComma = 0 Or TmpPosComma > TmpPosCloseParen)
                  
  'Single Arg, No Nested Parens or Func Calls
  If TmpIsSimpleArg Then
    SimpleArgText = Mid$(AFrmStr, APos + 1, TmpPosCloseParen - APos) 'includes closing paren.
  Else
    
    TmpUntilCloseStr = Mid$(AFrmStr, APos + 1, TmpPosCloseParen - APos)
    TmpCommaCnt = StrCnt(RemoveQuotedSections(TmpUntilCloseStr), ",")
    TmpOpenCnt = StrCnt(RemoveQuotedSections(TmpUntilCloseStr), "(")

    'No Nested Parens or Func Calls and 3 args or less. Make it 4 or 5?
    If TmpOpenCnt = 0 And _
       (TmpCommaCnt < kSimpleArgCnt) And _
       Len(TmpUntilCloseStr) < kSimpleArgTextLen Then
       
      TmpIsSimpleArg = True
    End If
     
    If TmpIsSimpleArg Then
      'Replace comma with comma space in non-quoted sections.
      SimpleArgText = ReplaceSkipQuotedSections(1, Mid$(AFrmStr, APos + 1, TmpPosCloseParen - APos), ",", ", ")
    Else
      'Only 1 function call and no longer than X chars.
      If TmpOpenCnt = 1 And Len(TmpUntilCloseStr) < kSimpleArgTextLen Then
        SimpleArgText = ReplaceSkipQuotedSections(1, Mid$(AFrmStr, APos + 1, TmpPosCloseParen - APos), ",", ", ")
      Else
        'Opportunity to construct other rules here.
        'Future - Could look for Operators and put them on a new line if line is too long.
      End If
    End If
  End If

End Function


Private Sub TestSimpleArgText()
  'Debug.Print SimpleArgText("=SUM(LEN(TRIM(Q$5#))-LEN(SUBSTITUTE(TRIM(Q$5#),$Z10,"")))+ROW($AD10)/10000", 14)
  'Debug.Print SimpleArgText("=@IF(TRIM($A2)<>"""",IF(IFERROR(FIND("" - "",$A2),0)>0,INDEX(TEXTSPLIT($A2,""|"","" - ""),1,2),INDEX(TEXTSPLIT($A2,""|""),2)),"""")", 30)
  'Debug.Print SimpleArgText("IF(IFERROR(FIND("" - "",$A2),0)>0,INDEX(TEXTSPLIT($A2,""|"","" - ""),1,2)", 11)

  Debug.Print SimpleArgText("=IF(TRIM([@[Time Loc Index]])<>"""",INDEX(tblTimeLocations[UTC Offset],[@[Time Loc Index]],1)+IF(INDEX(tblTimeLocations[DST True/False],[@[Time Loc Index]],1),INDEX(tblTimeLocations[DST Now],[@[Time Loc Index]],1),FALSE),"""")", 101)

End Sub


' ----------------------- FuncNameBeforeOpenParen Documentation -------------------------------'
'
' FuncNameBeforeOpenParen doesn't check for invalid beginning positions. An = should alway break us out.
'
'
'

Private Function FuncNameBeforeOpenParen(ByVal AFrmStr As String, _
                                         ByVal AOpenParenPos As Long) As String
                                         
  Dim TmpFuncNameCharCnt As Long
  Dim TmpCh  As String

  FuncNameBeforeOpenParen = vbNullString
  
  TmpFuncNameCharCnt = 0
  TmpCh = Mid$(AFrmStr, AOpenParenPos - TmpFuncNameCharCnt - 1, 1)
  Do While TmpCh Like "[A-Za-z0-9_.]"
    Inc TmpFuncNameCharCnt
    TmpCh = Mid$(AFrmStr, AOpenParenPos - TmpFuncNameCharCnt - 1, 1)
  Loop

  If TmpFuncNameCharCnt > 0 Then
    FuncNameBeforeOpenParen = Mid$(AFrmStr, AOpenParenPos - TmpFuncNameCharCnt, TmpFuncNameCharCnt)
  End If

End Function


Private Sub TestFuncNameBeforeOpenParen()

  ' Basic
  Debug.Print FuncNameBeforeOpenParen("=LET(x, A1 + B1)", 5)                  ' LET
  Debug.Print FuncNameBeforeOpenParen("=SUM(A1:A10)", 5)                     ' SUM
  Debug.Print FuncNameBeforeOpenParen("=AVERAGE(A1:A10)", 9)                 ' AVERAGE

  ' Nested functions
  Debug.Print FuncNameBeforeOpenParen("=IF(SUM(A1:A3)>0,1,0)", 4)             ' IF
  Debug.Print FuncNameBeforeOpenParen("=IFERROR(SUM(A1:A3),0)", 8)            ' IFERROR
  Debug.Print FuncNameBeforeOpenParen("=SUM(IF(A1>0,A1,0))", 5)               ' SUM

  ' Dotted / namespaced
  Debug.Print FuncNameBeforeOpenParen("=MyFunc.SubFunc(A1)", 15)              ' MyFunc.SubFunc
  Debug.Print FuncNameBeforeOpenParen("=NS1.NS2.Func123(A1)", 16)             ' NS1.NS2.Func123

  ' Underscores & numbers
  Debug.Print FuncNameBeforeOpenParen("=FUNC_1(A1)", 8)                       ' FUNC_1
  Debug.Print FuncNameBeforeOpenParen("=FUNC_TEST_99(A1)", 14)                ' FUNC_TEST_99

  ' Leading spaces
  Debug.Print FuncNameBeforeOpenParen("=  SUM(A1)", 7)                        ' SUM

  ' No function name
  Debug.Print FuncNameBeforeOpenParen("=(A1+A2)", 2)                          ' (empty)
  
  'Got rid of because Excel Formula has to begein with =. Then we wait for first (. Non-real scenario.
  'Debug.Print FuncNameBeforeOpenParen("=A1+A2", 0)                            ' (empty)

  ' Weird but legal
  Debug.Print FuncNameBeforeOpenParen("=_.__.(A1)", 6)                        ' _.__.

End Sub


Private Function PairArgPostProcess(ByVal AStr As String) As String
  Dim TmpFuncNameArr(100) As String
  Dim TmpArgNumArr(100)   As Long
  Dim TmpCnt              As Long
  Dim TmpChar             As String
  Dim TmpResult           As String
  Dim TmpDepth            As Long
  Dim TmpInQuotes         As Boolean
  Dim TmpInSingleQuotes   As Boolean
  Dim TmpInBrackets       As Boolean
  Dim TmpInBracketsCnt    As Long
  Dim TmpShouldPair       As Boolean

  PairArgPostProcess = AStr

  If InStr(1, AStr, "LET(", vbTextCompare) = 0 And _
     InStr(1, AStr, "IFS(", vbTextCompare) = 0 And _
     InStr(1, AStr, "SWITCH(", vbTextCompare) = 0 Then
    Exit Function
  End If

  TmpDepth = 0
  TmpResult = ""
  TmpCnt = 1

  Do While TmpCnt <= Len(AStr)
    TmpChar = Mid$(AStr, TmpCnt, 1)

    If TmpInQuotes Then
      TmpResult = TmpResult & TmpChar
      If TmpChar = """" Then TmpInQuotes = False

    ElseIf TmpInSingleQuotes Then
      TmpResult = TmpResult & TmpChar
      If TmpChar = "'" Then TmpInSingleQuotes = False

    ElseIf TmpInBrackets Then
      TmpResult = TmpResult & TmpChar
      If TmpChar = "]" Then
        Dec TmpInBracketsCnt
        TmpInBrackets = (TmpInBracketsCnt > 0)
      End If

    Else

      Select Case TmpChar
        Case "("
          Inc TmpDepth
          TmpFuncNameArr(TmpDepth) = UCase$(FuncNameBeforeOpenParen(AStr, TmpCnt))
          TmpArgNumArr(TmpDepth) = 1
          TmpResult = TmpResult & TmpChar

        Case ")"
          Dec TmpDepth
          TmpResult = TmpResult & TmpChar

        Case """"
          TmpInQuotes = True
          TmpResult = TmpResult & TmpChar

        Case "'"
          TmpInSingleQuotes = True
          TmpResult = TmpResult & TmpChar

        Case "["
          Inc TmpInBracketsCnt
          TmpInBrackets = True
          TmpResult = TmpResult & TmpChar

        Case ","
          TmpResult = TmpResult & TmpChar

          If TmpDepth > 0 Then
            Inc TmpArgNumArr(TmpDepth)

            TmpShouldPair = False
            Select Case TmpFuncNameArr(TmpDepth)
              Case "LET", "IFS"
                TmpShouldPair = (TmpArgNumArr(TmpDepth) Mod 2 = 0)
              Case "SWITCH"
                TmpShouldPair = (TmpArgNumArr(TmpDepth) Mod 2 = 1) And TmpArgNumArr(TmpDepth) > 1
            End Select

            If TmpShouldPair Then
              TmpCnt = TmpCnt + 1
              Do While TmpCnt <= Len(AStr)
                TmpChar = Mid$(AStr, TmpCnt, 1)
                If TmpChar = vbLf Or TmpChar = " " Then
                  TmpCnt = TmpCnt + 1
                Else
                  Exit Do
                End If
              Loop
              TmpResult = TmpResult & " "
              TmpCnt = TmpCnt - 1
            End If
          End If

        Case Else
          TmpResult = TmpResult & TmpChar
      End Select

    End If

    TmpCnt = TmpCnt + 1
  Loop

  PairArgPostProcess = TmpResult

End Function


Private Sub TestPairArgPostProcess()
  TestPrettyPrintHelper "=LET(a, 1, b, 2, a + b)"
  TestPrettyPrintHelper "=IFS(A1 > 10, ""High"", A1 > 5, ""Mid"", TRUE, ""Low"")"
  TestPrettyPrintHelper "=SWITCH(A1, 1, ""One"", 2, ""Two"", ""Other"")"
  TestPrettyPrintHelper "=SWITCH(A1, 1, ""One"", 2, ""Two"", ""Default"")"
End Sub


' ----------------------- PrettyPrint Pipeline -----------------------------------'
'                                                                                  '
' Step 1 - ShortenFormula                                                          '
'   Strips all whitespace outside quoted/bracketed sections (minify).              '
'                                                                                  '
' Step 2 - PrettyPrint (char-by-char)                                              '
'   Tracks context: double-quoted strings, single-quoted sheet names, brackets.   '
'   Inserts newlines and 4-space indentation on:                                   '
'     "(" - calls SimpleArgText; if args are simple, keeps on one line.            '
'     ")" - outdents.                                                              '
'     "," - new line at current indent level.                                      '
'   Pushes function names onto CStack as each "(" is entered.                     '
'                                                                                  '
' Step 3 - AddOperatorWhiteSpace                                                   '
'   Adds spaces around binary operators (+, -, *, /, ^, &, =, <, >, <=, >=, <>). '
'   Skips quoted strings, sheet names, and brackets.                               '
'   Detects unary minus (preceded by comma, open paren, or operator) and          '
'   appends it without stripping indentation.                                      '
'                                                                                  '
' Step 4 - PairArgPostProcess                                                      '
'   Collapses paired arguments for LET, IFS, and SWITCH at any nesting depth:     '
'     LET    - pulls even args (value) up onto the same line as the preceding     '
'              odd arg (name). Last (calc) arg stays on its own line.              '
'     IFS    - pulls even args (result) up onto the condition line.               '
'     SWITCH - pulls odd args >1 (result) up onto the value line.                 '
'              Arg 1 (expression) and any default (even, unpaired) stay alone.    '
'                                                                                  '
' -------------------------------------------------------------------------------- '

Private Function PrettyPrint(ByVal AFrmStr As String) As String
  Dim TmpLevel As Long
  Dim TmpInQuotes As Boolean
  Dim TmpInSingleQuotes As Boolean
  Dim TmpResult As String
  Dim TmpCnt As Long
  Dim TmpChar As String
  Dim TmpFrmStr As String
  Dim TmpInBrackets As Boolean
  Dim TmpInBracketsCnt As Long
  Dim TmpInBraces As Boolean
  Dim TmpInBracesCnt As Long
  Dim TmpSimpleArgText As String
  Dim TmpFuncName As String
  Dim TmpFuncNameCharCnt As Long
  Dim TmpStk As CStack
  Dim TmpParenDepth As Long
  
  TmpFrmStr = ShortenFormula(AFrmStr)

  TmpLevel = 0
  TmpInQuotes = False
  TmpInSingleQuotes = False
  TmpInBrackets = False
  TmpInBraces = False
  TmpCnt = 1
  Set TmpStk = New CStack
  For TmpCnt = 1 To Len(TmpFrmStr)
    TmpChar = Mid$(TmpFrmStr, TmpCnt, 1)

    If TmpInQuotes Then

      TmpResult = TmpResult & TmpChar
      If TmpChar = """" Then TmpInQuotes = False

    ElseIf TmpInSingleQuotes Then

      TmpResult = TmpResult & TmpChar
      If TmpChar = "'" Then TmpInSingleQuotes = False

    ElseIf TmpInBrackets Then

      TmpResult = TmpResult & TmpChar
      If TmpChar = "]" Then
        Dec TmpInBracketsCnt
        TmpInBrackets = (TmpInBracketsCnt > 0)
      End If

    ElseIf TmpInBraces Then 'Array literal, e.g. {"A","B","C"} -- Excel silently

      TmpResult = TmpResult & TmpChar 'collapses these back to one line, so never break inside one.
      If TmpChar = "}" Then
        Dec TmpInBracesCnt
        TmpInBraces = (TmpInBracesCnt > 0)
      End If

    Else

      Select Case TmpChar
        Case "("
          Inc TmpParenDepth
     
          TmpFuncName = FuncNameBeforeOpenParen(TmpFrmStr, TmpCnt)
        
          If TmpFuncName <> vbNullString Then
            TmpStk.Push TmpFuncName
          End If
        
          TmpResult = TmpResult & TmpChar
          
          TmpSimpleArgText = SimpleArgText(TmpFrmStr, TmpCnt)
          
          If TmpSimpleArgText <> vbNullString Then
            TmpResult = TmpResult & TmpSimpleArgText
            TmpCnt = TmpCnt + Len(TmpSimpleArgText) - StrCnt(RemoveQuotedSections(TmpSimpleArgText), ", ") 'Space inserted by SimpleArgText
          Else
            Inc TmpLevel

            TmpResult = TmpResult & vbLf
            TmpResult = TmpResult & String(TmpLevel * 4, " ")
          
          End If
          
        Case ")"
          Dec TmpParenDepth
          Dec TmpLevel
          TmpStk.Pop
          TmpResult = TmpResult & vbLf
          TmpResult = TmpResult & String(TmpLevel * 4, " ")
          TmpResult = TmpResult & TmpChar

        Case """"
          TmpInQuotes = Not TmpInQuotes
          TmpResult = TmpResult & TmpChar

        Case "'"
          TmpInSingleQuotes = Not TmpInSingleQuotes
          TmpResult = TmpResult & TmpChar

        Case ","
          TmpResult = TmpResult & TmpChar
          TmpResult = TmpResult & vbLf
          TmpResult = TmpResult & String(TmpLevel * 4, " ")

        Case "["
          Inc TmpInBracketsCnt
          TmpInBrackets = True
          TmpResult = TmpResult & TmpChar

        Case "{"
          Inc TmpInBracesCnt
          TmpInBraces = True
          TmpResult = TmpResult & TmpChar

        Case Else
          TmpResult = TmpResult & TmpChar
      End Select

      End If

  Next TmpCnt

  TmpResult = AddOperatorWhiteSpace(TmpResult)
  TmpResult = PairArgPostProcess(TmpResult)

  Set TmpStk = Nothing

  PrettyPrint = TmpResult

End Function


'True if pretty-printing AFrmStr would actually change it. Simple formulas
'(e.g. =SUM(A1:A10)) map to themselves under Pretty Print regardless of
'whether they started minified or not, so this is False for them.
Private Function PrettyPrintWouldChangeFormula(ByVal AFrmStr As String) As Boolean
  PrettyPrintWouldChangeFormula = (PrettyPrint(AFrmStr) <> AFrmStr)
End Function


Private Sub TestPrettyPrintWouldChangeFormulaHelper(ByVal AFrmStr As String, ByVal AExpectedBool As Boolean)
  Dim TmpAnswerBool As Boolean
  Dim TmpPassFailStr As String

  TmpAnswerBool = PrettyPrintWouldChangeFormula(AFrmStr)
  TmpPassFailStr = IIf(TmpAnswerBool = AExpectedBool, "PASS", "FAIL")

  Debug.Print AFrmStr; " | "; TmpAnswerBool; " | "; AExpectedBool; " | "; TmpPassFailStr
End Sub

Private Sub TestPrettyPrintWouldChangeFormula()
  ' Simple formula: Pretty Print maps it to itself either way.
  TestPrettyPrintWouldChangeFormulaHelper "=SUM(A1:A10)", False

  ' Complex, still minified: Pretty Print would add line breaks.
  TestPrettyPrintWouldChangeFormulaHelper "=IF(A1>0,SUM(A1:A10),0)", True

  ' Complex, already pretty-printed: idempotent, nothing left to do.
  TestPrettyPrintWouldChangeFormulaHelper PrettyPrint("=IF(A1>0,SUM(A1:A10),0)"), False

  ' Array literal, already pretty-printed: must stay idempotent -- {} content
  ' must never get a line break, since Excel silently collapses it back to
  ' one line and would otherwise make this loop forever.
  TestPrettyPrintWouldChangeFormulaHelper _
    PrettyPrint("=CHOOSE(MATCH(Data!E3,{""A"",""B"",""C""},0),""Alpha"",""Beta"",""Gamma"")"), False
End Sub


Private Sub TestPrettyPrintHelper(ByVal AFrmStr As String)
  LogWrite "-------------------------------------------------"
  LogWrite AFrmStr
  LogWrite PrettyPrint(AFrmStr)
End Sub


Private Sub TestPrettyPrint()
  Dim TmpStr As String

  ClearImmediateWindow
  LogClear

  TestPrettyPrintHelper "=IF(A1>0, SUM((A1:B1) + (C1:D1)), ""Test,,"")" & vbCrLf & vbCrLf

  TestPrettyPrintHelper "=CONCAT(""Hello """"World,"""""",""End"")"

  TestPrettyPrintHelper "=IF(TRIM([@[UTC Offset]])<>"""", $B$2 + ([@[UTC Offset]]/24), """")"
  TestPrettyPrintHelper "=IF(TRUE,IF(A1=1,B1+C1,D1),FALSE)"

  TestPrettyPrintHelper "=LAMBDA(Char,CharStatus,Guess,GameArr,CharStatusArr,SUM(MAKEARRAY(1,5,LAMBDA(r,C,IF(AND(INDEX(GameArr,r,C)=Char,INDEX(CharStatusArr,r,C)<>""Not in Word""),1,0)))))(Game!D5,D5,$AP5,Game!$B5:$F5,Interpret!$B5:$F5)"

  TestPrettyPrintHelper "=IFERROR(IF($H18<>"""",XLOOKUP($H18,'Client Forecast'!$K:$K,'Client Forecast'!C:C),""""),"""")"

  TestPrettyPrintHelper "=IF(TRIM($A2)<>"""",INDEX(TEXTSPLIT($A2,""|""),1,3),"""")"
  TestPrettyPrintHelper "=@IF($AF2,IF($AJ2,IF($AP2>0,INDEX(CSV_Std_Translations,1,$AP2),""Blank""),IF(TRIM($C2)<>"""",$C2,""Blank"")),"""")"
  TestPrettyPrintHelper "=IF(IFERROR(FIND("" - "",INDEX(TEXTSPLIT($A2,""|""),1,2)),0)>=0,INDEX(TEXTSPLIT($A2,""|""),1,2),"""")"
  TestPrettyPrintHelper "=IF(A9,IF(K9<>0,TRIM(MID(B9,J9,FIND(K9,B9,J9)-J9)),MID(B9,J9,2)),"""")"

  TestPrettyPrintHelper "=IF(A1<>"""", IF(B1>0, IF(C1=1, ""Valid"", ""Invalid""), ""Error""), ""Blank"")"
  TestPrettyPrintHelper "=IF(A1 <> """", CONCAT(""Prefix-"", TEXTJOIN("", "", TRUE, B1, C1, D1)), ""No Data"")"
  TestPrettyPrintHelper "=SUM(XLOOKUP(A1, Table1[Lookup], Table1[Value], 0) * B1, C1)"
  TestPrettyPrintHelper "=IF([@[Net Profit]] > 0, ""Profitable"", IF([@[Net Profit]] = 0, ""Break Even"", ""Loss""))"
  TestPrettyPrintHelper "=INDIRECT(ADDRESS(A1, B1, 4))"
  TestPrettyPrintHelper "=SUM(SEQUENCE(5, 1, A1, 1) * B1)"
  TestPrettyPrintHelper "=LET(x, A1 + B1, y, x * 2, y - C1)"

  TestPrettyPrintHelper "=CHOOSE(A1, 10, 20, ""Text Value"", ""Another Text"")"
  TestPrettyPrintHelper "=IF(AND([@[% In]],[@[Votes In]]),[@[Rep Need]]/[@[VL-3VL]],0)"

  TestPrettyPrintHelper "=LAMBDA(Input,Cnt,Pos,Incl,IFS((Cnt+Pos)<0,""[]"",(Cnt+Pos)>=8,""[]"",NOT(Incl),""[""&MID(Input,(Pos+Cnt),(ABS(Cnt)))&""]"",TRUE,""[""&MID(Input,(Pos+Cnt+1),(ABS(Cnt)))&""]""))($N70,$P70,$Q70,$O70)"

  TestPrettyPrintHelper "=LET(x, ((A1+B1)), y, IF(x>0, LET(u, x*2, v, u+3, v), ((x))), y)"
  TestPrettyPrintHelper "=LET(x, (A1 + B1) * ((C1 + D1)), y, LET(u, x + 1, v, u * 3, v), y + ((x)))"
  TestPrettyPrintHelper "=LET(x, A1 + (B1 + (C1)), y, x * 2, y - ((C1)))"

  TestPrettyPrintHelper "=IF([@Active]=FALSE,0,IFERROR(LET(Rid,[@Id],Keep,(tblRecipeBOM[Recipe Id]=Rid)*(1-((tblRecipeBOM[Item Type]=""Recipe"")*(tblRecipeBOM[Item Id]=Rid))),Qty,FILTER(tblRecipeBOM[Qty],Keep),Type,FILTER(tblRecipeBOM[Item Type],Keep),ID,FILTER(tblRecipeBOM[Item Id],Keep),val,SWITCH(Type,""Ingredient"",XLOOKUP(ID,tblIngredients[Id],tblIngredients[Calories],0),""Recipe"",XLOOKUP(ID,[Id],[Calories],0),0),ROUND(SUM(Qty*val),1)),""""))"

  TestPrettyPrintHelper "=IF(EOMONTH(K$2,0)<=EOMONTH('FTM P&L'!$O$2,0),XLOOKUP(EOMONTH(K$2,0),'Import Values from WB4'!$D$3300:$AM$3300,'Import Values from WB4'!$D3309:$AM3309,0,-1),0)"

  TestPrettyPrintHelper "=SUMPRODUCT(--(tblData[Status]=""Active""),--(tblData[Amount]>100),tblData[Value])"

  'Comma and open-paren inside a bracketed structured reference must stay on
  'one line -- bracket-tracking must protect them like quotes do.
  TestPrettyPrintHelper "=SUM([@[Revenue, Total]])"
  TestPrettyPrintHelper "=SUM([@[Total (Net)]])"

  'Array literal must never get a line break inside it -- Excel silently
  'collapses {} back to one line, which would otherwise break idempotence.
  TestPrettyPrintHelper "=CHOOSE(MATCH(Data!E3,{""A"",""B"",""C""},0),""Alpha"",""Beta"",""Gamma"")"

  TestPairArgPostProcess

End Sub


'Covers two bugs: (1) ResultStr* used to render blank instead of "0" for
'zero counts, (2) *Rng used to return False (Error) when a range had no
'formulas at all, rather than True (nothing to do).
Private Sub TestPrettyPrintNoFormulaRng()
  Dim TmpSuccCnt As Long
  Dim TmpErrCnt As Long
  Dim TmpResult As Boolean
  Dim TmpRng As Range

  ClearImmediateWindow

  'Assumes A1 on the active sheet has no formula.
  Set TmpRng = ActiveSheet.Range("A1")
  If TmpRng.HasFormula Then
    Debug.Print "SKIPPED: A1 has a formula; re-run against a non-formula cell."
    Exit Sub
  End If

  TmpResult = PrettyPrintRng(TmpRng, TmpSuccCnt, TmpErrCnt)
  Debug.Print "PrettyPrintRng (no formulas) Result: "; TmpResult; " SuccCnt: "; TmpSuccCnt; " ErrCnt: "; TmpErrCnt
  Debug.Print "Expect True/0/0. ResultStr: " & ResultStrPrettyPrint(TmpResult, TmpSuccCnt, TmpErrCnt)

  TmpSuccCnt = 0
  TmpErrCnt = 0
  TmpResult = MinifyRng(TmpRng, TmpSuccCnt, TmpErrCnt)
  Debug.Print "MinifyRng (no formulas) Result: "; TmpResult; " SuccCnt: "; TmpSuccCnt; " ErrCnt: "; TmpErrCnt
  Debug.Print "Expect True/0/0. ResultStr: " & ResultStrMinify(TmpResult, TmpSuccCnt, TmpErrCnt)

  Debug.Print "Zero-count formatting (Error branch): " & ResultStrPrettyPrint(False, 0, 0)
  Debug.Print "Expect the two 0 counts to display as literal ""0"", not blank."

End Sub


'----------------------------------------------------------------------'
'                        Cell Functions                                '
'----------------------------------------------------------------------'

'AFrmCell is presumed to hold a formula.
Private Function MinifyCell(ByVal AFrmCell As Range) As Boolean
  Dim TmpOrigStr As String
  Dim TmpMinifyStr As String

  MinifyCell = False

  'Saves
  TmpOrigStr = AFrmCell.Formula2
  
  'Get Minify Str
  TmpMinifyStr = ShortenFormula(TmpOrigStr)

  If CellSetFormula2Safe(AFrmCell, TmpMinifyStr) = False Then
    Debug.Print "--------------- Minify Cell Error ---------------------" & vbCrLf
    Debug.Print ShtFormulaNameStr(AFrmCell.Worksheet.Name) & "!" & AFrmCell.Address
    Debug.Print "Original: "
    Debug.Print TmpOrigStr
    Debug.Print "Minified: "
    Debug.Print TmpMinifyStr & vbCrLf
    Debug.Print "--------------------------------------------------------" & vbCrLf
  End If

  MinifyCell = True

End Function


'Does it really matter if the cell wasn't a formula to begin with? No.
Private Function CellSetFormula2Safe(ByVal ACell As Range, ByVal AFrmStr As String) As Boolean
  Dim TmpAppEvents As Boolean
  Dim TmpNbrFmtStr As String

  CellSetFormula2Safe = False

  'Saves
  TmpAppEvents = Application.EnableEvents
  TmpNbrFmtStr = ACell.NumberFormat
  
  'Set Off
  Application.EnableEvents = False
  Application.DisplayAlerts = False
  
  'Set Error Handling
  On Error GoTo OnError
  
  'Set to General Format to Prevent Format from Changing to Text.
  'Quick Note: If NumberFormat is Text, resetting Formula2 cause the formula to show as text.
  ACell.NumberFormat = "General"
  
  'Set Formula2
  ACell.Formula2 = AFrmStr

  'Should now have a Formula
  If ACell.HasFormula = False Then
    err.Raise Number:=65536, Description:="Formula Changed to Text."
  End If

  'No Errors and is a Formula
  CellSetFormula2Safe = True
  
Finally:
  
  'Restores
  ACell.NumberFormat = TmpNbrFmtStr 'Note: Even if manually changing from date to formula, NumberFormat doesn't change.
  Application.EnableEvents = TmpAppEvents
  Application.DisplayAlerts = True

  Exit Function
OnError:

  Debug.Print "--------------- CellSetFormula2Safe Error ---------------------" & vbCrLf
  Debug.Print "Error " & err.Number & ": " & err.Description
  Debug.Print ShtFormulaNameStr(ACell.Worksheet.Name) & "!" & ACell.Address
  Debug.Print "AFrmStr: "
  Debug.Print AFrmStr & vbCrLf
  Debug.Print "--------------------------------------------------------" & vbCrLf

  On Error GoTo 0
  
  Resume Finally
End Function


Private Sub TestCellSetFormula2Safe()
  Dim TmpResult As Boolean
  Dim TmpFrmStr As String

  TmpFrmStr = "=SUM(A2:A10)"

  TmpResult = CellSetFormula2Safe(ActiveSheet.Range("A1"), TmpFrmStr)
  
  Debug.Print "Result: "; TmpResult
End Sub



'AFrmCell is presumed to hold a formula.
Private Function PrettyPrintCell(ByVal AFrmCell As Range) As Boolean
  Dim TmpOrigStr As String
  Dim TmpPPStr As String
  
  PrettyPrintCell = False
  
  'Saves
  TmpOrigStr = AFrmCell.Formula2 'For Error Printing.
  
  'Get Pretty Print Formula String
  TmpPPStr = PrettyPrint(TmpOrigStr)
  
  'Set Formula2
  If CellSetFormula2Safe(AFrmCell, TmpPPStr) = False Then
    Debug.Print "--------------- Pretty Print Error ---------------------" & vbCrLf
    Debug.Print ShtFormulaNameStr(AFrmCell.Worksheet.Name) & "!" & AFrmCell.Address
    Debug.Print "Original: "
    Debug.Print TmpOrigStr
    Debug.Print "Pretty Print: "
    Debug.Print TmpPPStr & vbCrLf
    Debug.Print "--------------------------------------------------------" & vbCrLf
   
    Exit Function
  End If
  
  PrettyPrintCell = True
End Function


'----------------------------------------------------------------------'
'                         Range Level Functions                        '
'----------------------------------------------------------------------'

Private Function MinifyRng(ByVal ARng As Range, _
                           ByRef ASuccCnt As Long, _
                           ByRef AErrCnt As Long) As Boolean
  Dim TmpCell As Range
  Dim TmpRng As Range
  Dim TmpOrigErrCnt As Long

  MinifyRng = True

  'Get Around 1 Cell Selected Searches Whole Sheet
  If ARng.Count = 1 Then
    If ARng.HasFormula Then
      Set TmpRng = ARng
    Else
      Exit Function
    End If
  Else
    Set TmpRng = SpecialCellsSafe(ARng, XlCellType.xlCellTypeFormulas)
  End If

  If TmpRng Is Nothing Then Exit Function
   
   
  TmpOrigErrCnt = AErrCnt
  For Each TmpCell In TmpRng
    If MinifyCell(TmpCell) Then Inc ASuccCnt Else Inc AErrCnt
  Next
  
  MinifyRng = (TmpOrigErrCnt = AErrCnt)
  
  Set TmpRng = Nothing
End Function


Private Function PrettyPrintRng(ByVal ARng As Range, _
                                ByRef ASuccCnt As Long, _
                                ByRef AErrCnt As Long) As Boolean
  Dim TmpCell As Range
  Dim TmpRng As Range
  Dim TmpOrigErrCnt As Long

  PrettyPrintRng = True

  'Get Around 1 Cell Selected Searches Whole Sheet
  If ARng.Count = 1 Then
    If ARng.HasFormula Then
      Set TmpRng = ARng
    Else
      Exit Function
    End If
  Else
    Set TmpRng = SpecialCellsSafe(ARng, XlCellType.xlCellTypeFormulas)
  End If

  If TmpRng Is Nothing Then Exit Function

  TmpOrigErrCnt = AErrCnt
  For Each TmpCell In TmpRng
    If PrettyPrintCell(TmpCell) Then Inc ASuccCnt Else Inc AErrCnt
  Next

  PrettyPrintRng = (TmpOrigErrCnt = AErrCnt)

  Set TmpRng = Nothing

End Function


'True if any formula cell in ARng would actually change under Pretty Print.
'Used by ToggleSel to decide Minify vs Pretty Print with a definitive test
'("is there anything left for Pretty Print to do?") rather than inferring
'from vbLf presence, which simple formulas can't distinguish.
Private Function PrettyPrintWouldChangeRng(ByVal ARng As Range) As Boolean
  Dim TmpCell As Range
  Dim TmpRng As Range

  PrettyPrintWouldChangeRng = False

  'Get Around 1 Cell Selected Searches Whole Sheet
  If ARng.Count = 1 Then
    If ARng.HasFormula Then
      PrettyPrintWouldChangeRng = PrettyPrintWouldChangeFormula(ARng.Formula2)
    End If
    Exit Function
  End If

  Set TmpRng = SpecialCellsSafe(ARng, XlCellType.xlCellTypeFormulas)
  If TmpRng Is Nothing Then Exit Function

  For Each TmpCell In TmpRng
    If PrettyPrintWouldChangeFormula(TmpCell.Formula2) Then
      PrettyPrintWouldChangeRng = True
      Exit Function
    End If
  Next TmpCell

End Function

'----------------------------------------------------------------------'
'                          Sheet Level Functions                       '
'----------------------------------------------------------------------'

Private Function MinifySht(ByVal ASht As Worksheet, _
                           ByRef ASuccCnt As Long, _
                           ByRef AErrCnt As Long) As Boolean
  MinifySht = False
  
  'Exit Conditions
  If ASht Is Nothing Then Exit Function
  
  'Process Used Range
  MinifySht = MinifyRng(ASht.UsedRange, ASuccCnt, AErrCnt)
  
End Function

Private Function PrettyPrintSht(ByVal ASht As Worksheet, _
                                ByRef ASuccCnt As Long, _
                                ByRef AErrCnt As Long) As Boolean
  PrettyPrintSht = False

  'Exit Conditions
  If ASht Is Nothing Then Exit Function
  
  'Process Used Range
  PrettyPrintSht = PrettyPrintRng(ASht.UsedRange, ASuccCnt, AErrCnt)

End Function


'----------------------------------------------------------------------'
'                       Workbook Level Functions                       '
'----------------------------------------------------------------------'

Private Function MinifyWbk(ByVal AWbk As Workbook, _
                           ByRef ASuccCnt As Long, _
                           ByRef AErrCnt As Long) As Boolean
  Dim TmpSht As Worksheet
  Dim TmpOrigErrCnt As Long

  MinifyWbk = False

  'Exit Conditions
  If AWbk Is Nothing Then Exit Function

  TmpOrigErrCnt = AErrCnt

  For Each TmpSht In AWbk.Worksheets
    MinifySht TmpSht, ASuccCnt, AErrCnt
  Next TmpSht

  MinifyWbk = (TmpOrigErrCnt = AErrCnt)

End Function


Private Function PrettyPrintWbk(ByVal AWbk As Workbook, _
                                ByRef ASuccCnt As Long, _
                                ByRef AErrCnt As Long) As Boolean
  Dim TmpSht As Worksheet
  Dim TmpOrigErrCnt As Long

  PrettyPrintWbk = False

  'Exit Conditions
  If AWbk Is Nothing Then Exit Function

  TmpOrigErrCnt = AErrCnt

  For Each TmpSht In AWbk.Worksheets
    PrettyPrintSht TmpSht, ASuccCnt, AErrCnt
  Next TmpSht

  PrettyPrintWbk = (TmpOrigErrCnt = AErrCnt)

End Function


'----------------------------------------------------------------------'
'                         Public Functions                             '
'----------------------------------------------------------------------'

'Ribbon default action: Pretty Print if anything in the selection would
'actually change under Pretty Print; Minify if nothing would (i.e. every
'formula is already as pretty as it gets).
Public Sub ToggleSel()
  If Selection Is Nothing Then Exit Sub

  If PrettyPrintWouldChangeRng(Selection) Then
    PrettyPrintSel
  Else
    MinifySel
  End If
End Sub

Public Sub MinifySel()
  Dim TmpSuccCnt As Long
  Dim TmpErrCnt As Long
  Dim TmpResultStr As String
  Dim TmpResult As Boolean

  If AllClear2(AIsWbkReadOnlyWarning:=True, _
               AWbkProtectedStructWarning:=False, _
               AIsActiveShtWorksheetWarning:=True, _
               AProtectedActiveShtWarning:=True, _
               AAnyProtectedShtWarning:=False, _
               AIsSelRngWarning:=True, _
               AIsSelSingleAreaWarning:=False) = False Then Exit Sub

  MacroSpeedup msInit
  TmpResult = MinifyRng(Selection, TmpSuccCnt, TmpErrCnt)
  MacroSpeedup msClear
  
  TmpResultStr = ResultStrMinify(TmpResult, TmpSuccCnt, TmpErrCnt)
End Sub


Public Sub PrettyPrintSel()
  Dim TmpSuccCnt As Long
  Dim TmpErrCnt As Long
  Dim TmpResultStr As String
  Dim TmpResult As Boolean

  If AllClear2(AIsWbkReadOnlyWarning:=True, _
               AWbkProtectedStructWarning:=False, _
               AIsActiveShtWorksheetWarning:=True, _
               AProtectedActiveShtWarning:=True, _
               AAnyProtectedShtWarning:=False, _
               AIsSelRngWarning:=True, _
               AIsSelSingleAreaWarning:=False) = False Then Exit Sub

  MacroSpeedup msInit
  TmpResult = PrettyPrintRng(Selection, TmpSuccCnt, TmpErrCnt)
  MacroSpeedup msClear
  
  TmpResultStr = ResultStrPrettyPrint(TmpResult, TmpSuccCnt, TmpErrCnt)
End Sub

Public Sub MinifyActiveSht()
  Dim TmpSuccCnt As Long
  Dim TmpErrCnt As Long
  Dim TmpResultStr As String
  Dim TmpResult As Boolean

  If AllClear2(AIsWbkReadOnlyWarning:=True, _
               AWbkProtectedStructWarning:=False, _
               AIsActiveShtWorksheetWarning:=True, _
               AProtectedActiveShtWarning:=True, _
               AAnyProtectedShtWarning:=False, _
               AIsSelRngWarning:=False, _
               AIsSelSingleAreaWarning:=False) = False Then Exit Sub

  MacroSpeedup msInit
  TmpResult = MinifySht(ActiveSheet, TmpSuccCnt, TmpErrCnt)
  MacroSpeedup msClear
  
  TmpResultStr = ResultStrMinify(TmpResult, TmpSuccCnt, TmpErrCnt)
  
  MsgBox TmpResultStr
End Sub

Public Sub PrettyPrintActiveSht()
  Dim TmpSuccCnt   As Long
  Dim TmpErrCnt    As Long
  Dim TmpResultStr As String
  Dim TmpResult    As Boolean

  If AllClear2(AIsWbkReadOnlyWarning:=True, _
               AWbkProtectedStructWarning:=False, _
               AIsActiveShtWorksheetWarning:=True, _
               AProtectedActiveShtWarning:=True, _
               AAnyProtectedShtWarning:=False, _
               AIsSelRngWarning:=False, _
               AIsSelSingleAreaWarning:=False) = False Then Exit Sub

  MacroSpeedup msInit
  TmpResult = PrettyPrintSht(ActiveSheet, TmpSuccCnt, TmpErrCnt)
  MacroSpeedup msClear
  
  TmpResultStr = ResultStrPrettyPrint(TmpResult, TmpSuccCnt, TmpErrCnt)

  MsgBox TmpResultStr
End Sub


Public Sub MinifyActiveWbk()
  Dim TmpSuccCnt   As Long
  Dim TmpErrCnt    As Long
  Dim TmpResultStr As String
  Dim TmpResult    As Boolean

  If AllClear2(AIsWbkReadOnlyWarning:=True, _
               AWbkProtectedStructWarning:=False, _
               AIsActiveShtWorksheetWarning:=False, _
               AProtectedActiveShtWarning:=True, _
               AAnyProtectedShtWarning:=True, _
               AIsSelRngWarning:=False, _
               AIsSelSingleAreaWarning:=False) = False Then Exit Sub

  MacroSpeedup msInit
  TmpResult = MinifyWbk(ActiveWorkbook, TmpSuccCnt, TmpErrCnt)
  MacroSpeedup msClear

  TmpResultStr = ResultStrMinify(TmpResult, TmpSuccCnt, TmpErrCnt)

  MsgBox TmpResultStr
End Sub


Public Sub PrettyPrintActiveWbk()
  Dim TmpSuccCnt   As Long
  Dim TmpErrCnt    As Long
  Dim TmpResultStr As String
  Dim TmpResult    As Boolean

  If AllClear2(AIsWbkReadOnlyWarning:=True, _
               AWbkProtectedStructWarning:=False, _
               AIsActiveShtWorksheetWarning:=False, _
               AProtectedActiveShtWarning:=True, _
               AAnyProtectedShtWarning:=True, _
               AIsSelRngWarning:=False, _
               AIsSelSingleAreaWarning:=False) = False Then Exit Sub

  MacroSpeedup msInit
  TmpResult = PrettyPrintWbk(ActiveWorkbook, TmpSuccCnt, TmpErrCnt)
  MacroSpeedup msClear
  
  TmpResultStr = ResultStrPrettyPrint(TmpResult, TmpSuccCnt, TmpErrCnt)

  MsgBox TmpResultStr
End Sub





