Attribute VB_Name = "modSmallFunctions"
Option Explicit
' Shared low-level string/formula helpers used by the Pretty Print feature
' (modPrettyPrint.bas, modAddOperatorWhiteSpace.bas).
'
' Public declarations (in source order):
'   Inc
'   Dec
'   StrCnt
'   RemoveQuotedSections
'   PosSkipQuotedSections
'   ReplaceSkipQuotedSections
'   ShtNameRequiresSingleQuotes
'   ShtFormulaNameStr
'   ShortenFormula
'   ClearImmediateWindow
'   SpecialCellsSafe

Public Sub Inc(ByRef AVal As Variant)
  AVal = AVal + 1
End Sub

Public Sub Dec(ByRef AVal As Variant)
  AVal = AVal - 1
End Sub

Public Function GetAppDataDir() As String
  GetAppDataDir = Environ$("AppData")
End Function

Function DirExists(ByVal AFileNameStr As String) As Boolean
  On Error Resume Next
  DirExists = ((GetAttr(AFileNameStr) And vbDirectory) = vbDirectory)
End Function

Function IncludeTrailingBackslash(ADirStr As String)
  If Right$(ADirStr, 1) <> "\" Then
    IncludeTrailingBackslash = ADirStr & "\"
  Else
    IncludeTrailingBackslash = ADirStr
  End If
End Function

Public Function StrCnt(AStr As String, ASubStr As String) As Long
  Dim TmpPos As Long
  StrCnt = 0

  TmpPos = 1

  Do
    TmpPos = InStr(TmpPos, AStr, ASubStr)
    If TmpPos > 0 Then
      StrCnt = StrCnt + 1
      TmpPos = TmpPos + Len(ASubStr) ' Move to the next position after the found occurrence
    End If
  Loop While TmpPos > 0
End Function


Private Sub TestStrCntHelper(ByVal AStr As String, ByVal ASubStr As String, ByVal AExpectedCnt As Long, Optional ByVal ATestFilePathStr As String = vbNullString)
  Dim TmpAnswerCnt As Long
  Dim TmpPassFailStr As String

  TmpAnswerCnt = StrCnt(AStr, ASubStr)
  TmpPassFailStr = IIf(TmpAnswerCnt = AExpectedCnt, "PASS", "FAIL")

  TestLogLine AStr & " | " & ASubStr & " | " & TmpAnswerCnt & " | " & AExpectedCnt & " | " & TmpPassFailStr, ATestFilePathStr
End Sub

Private Sub TestStrCnt(Optional ByVal ATestFilePathStr As String = vbNullString)
  TestStrCntHelper "/DDDD//", "//", 1, ATestFilePathStr
End Sub


Public Function RemoveQuotedSections(ByVal AStr As String) As String
  Dim TmpResult As String
  Dim TmpInsideQuotes As Boolean
  Dim TmpInBracketsCnt As Long
  Dim TmpInBraces As Boolean
  Dim TmpChar As String
  Dim TmpIsDelimiter As Boolean
  Dim i As Long

  TmpResult = vbNullString
  TmpInsideQuotes = False

  For i = 1 To Len(AStr)
    TmpChar = Mid$(AStr, i, 1)
    TmpIsDelimiter = False

    If TmpChar = """" Then
      TmpInsideQuotes = Not TmpInsideQuotes
      TmpIsDelimiter = True
    ElseIf Not TmpInsideQuotes And TmpChar = "[" Then
      TmpInBracketsCnt = TmpInBracketsCnt + 1
      TmpIsDelimiter = True
    ElseIf Not TmpInsideQuotes And TmpChar = "]" And TmpInBracketsCnt > 0 Then
      TmpInBracketsCnt = TmpInBracketsCnt - 1
      TmpIsDelimiter = True
    ElseIf Not TmpInsideQuotes And Not TmpInBraces And TmpChar = "{" Then
      TmpInBraces = True 'Array literal, e.g. {"A","B","C"} -- treat as one
      TmpIsDelimiter = True 'opaque unit, same as a quoted string (can't nest
    ElseIf Not TmpInsideQuotes And TmpInBraces And TmpChar = "}" Then 'in Excel).
      TmpInBraces = False
      TmpIsDelimiter = True
    End If

    If Not TmpIsDelimiter And Not TmpInsideQuotes And TmpInBracketsCnt = 0 And Not TmpInBraces Then
      TmpResult = TmpResult & TmpChar
    End If
  Next i

  RemoveQuotedSections = TmpResult
End Function


Private Sub TestRemoveQuotedSectionsHelper(ByVal AOriginalStr As String, ByVal AExpectedStr As String, Optional ByVal ATestFilePathStr As String = vbNullString)
  Dim TmpAnswerStr As String
  Dim TmpPassFailStr As String

  TmpAnswerStr = RemoveQuotedSections(AOriginalStr)
  TmpPassFailStr = IIf(TmpAnswerStr = AExpectedStr, "PASS", "FAIL")

  TestLogLine AOriginalStr & " | " & TmpAnswerStr & " | " & AExpectedStr & " | " & TmpPassFailStr, ATestFilePathStr
End Sub

Private Sub TestRemoveQuotedSections(Optional ByVal ATestFilePathStr As String = vbNullString)
  TestRemoveQuotedSectionsHelper "Test(" & """This is the middle""" & ")", "Test()", ATestFilePathStr
  TestRemoveQuotedSectionsHelper "SUM([@[Revenue, Total]])", "SUM()", ATestFilePathStr
  TestRemoveQuotedSectionsHelper "A,[Col1],B", "A,,B", ATestFilePathStr
  TestRemoveQuotedSectionsHelper "[Col1],[Col2]", ",", ATestFilePathStr
  TestRemoveQuotedSectionsHelper """[literal],text""" & ",A2", ",A2", ATestFilePathStr
  TestRemoveQuotedSectionsHelper "[]", vbNullString, ATestFilePathStr
  TestRemoveQuotedSectionsHelper "Data!E3,{""A"",""B"",""C""},0", "Data!E3,,0", ATestFilePathStr 'Array literal's internal commas excluded.
  TestRemoveQuotedSectionsHelper "A,{""X,Y"",""Z""},B", "A,,B", ATestFilePathStr 'Comma inside a quoted array element also excluded.
End Sub


Public Function PosSkipQuotedSections(ByVal AStart As Long, ByVal AStr As String, ByVal ASubStr As String) As Long
  Dim TmpInsideQuotes As Boolean
  Dim TmpInBrackets As Boolean
  Dim TmpInBracketsCnt As Long
  Dim TmpInBraces As Boolean
  Dim TmpChar As String
  Dim i As Long

  PosSkipQuotedSections = 0

  TmpInsideQuotes = False

  For i = AStart To Len(AStr)
    TmpChar = Mid$(AStr, i, 1)

    If TmpChar = """" Then
      TmpInsideQuotes = Not TmpInsideQuotes
    ElseIf Not TmpInsideQuotes And Not TmpInBrackets And Not TmpInBraces Then
      If Mid$(AStr, i, Len(ASubStr)) = ASubStr Then
        PosSkipQuotedSections = i
        Exit Function
      End If
    End If

    If Not TmpInsideQuotes Then
      If TmpChar = "[" Then TmpInBracketsCnt = TmpInBracketsCnt + 1
      If TmpChar = "]" Then TmpInBracketsCnt = TmpInBracketsCnt - 1
      TmpInBrackets = TmpInBracketsCnt > 0

      'Array literal, e.g. {"A","B","C"} -- can't nest in Excel, so a simple
      'toggle (not a counter) is enough.
      If TmpChar = "{" Then TmpInBraces = True
      If TmpChar = "}" Then TmpInBraces = False
    End If
  Next i

End Function


Private Sub TestPosSkipQuotedSectionsHelper(ByVal AStart As Long, ByVal AStr As String, ByVal ASubStr As String, ByVal AExpectedPos As Long, Optional ByVal ATestFilePathStr As String = vbNullString)
  Dim TmpAnswerPos As Long
  Dim TmpPassFailStr As String

  TmpAnswerPos = PosSkipQuotedSections(AStart, AStr, ASubStr)
  TmpPassFailStr = IIf(TmpAnswerPos = AExpectedPos, "PASS", "FAIL")

  TestLogLine AStart & " | " & AStr & " | " & ASubStr & " | " & TmpAnswerPos & " | " & AExpectedPos & " | " & TmpPassFailStr, ATestFilePathStr
End Sub

Private Sub TestPosSkipQuotedSections(Optional ByVal ATestFilePathStr As String = vbNullString)
  TestPosSkipQuotedSectionsHelper 1, "=IF(A1=""("",1,2)", ")", 15, ATestFilePathStr
  TestPosSkipQuotedSectionsHelper 1, "SUM([@[Revenue, Total]])", ",", 0, ATestFilePathStr 'Comma is inside brackets.
  TestPosSkipQuotedSectionsHelper 1, "Data!E3,{""A"",""B"",""C""},0", ",", 8, ATestFilePathStr 'Finds the real separator, skips the array's internal commas.
End Sub


Public Function ReplaceSkipQuotedSections(ByVal AStart As Long, _
                                          ByVal AStr As String, _
                                          ByVal ASubStr As String, _
                                          ByVal ARepStr As String) As String
  Dim TmpInsideQuotes As Boolean
  Dim TmpInBrackets As Boolean
  Dim TmpInBracketsCnt As Long
  Dim TmpChar As String
  Dim i As Long
  Dim TmpResult As String

  TmpResult = vbNullString
  TmpInsideQuotes = False

  For i = AStart To Len(AStr)
    TmpChar = Mid$(AStr, i, 1)

    If TmpChar = """" Then
      TmpInsideQuotes = Not TmpInsideQuotes
      TmpResult = TmpResult & TmpChar ' Always include quotes
    ElseIf TmpInsideQuotes Or TmpInBrackets Then
      TmpResult = TmpResult & TmpChar 'Always include bracketed content (structured references).
    Else
      If Mid$(AStr, i, Len(ASubStr)) = ASubStr Then
        ' Append up to the current position before skipping the substring
        TmpResult = TmpResult + ARepStr
      Else
        TmpResult = TmpResult & TmpChar
      End If
    End If

    If Not TmpInsideQuotes Then
      If TmpChar = "[" Then TmpInBracketsCnt = TmpInBracketsCnt + 1
      If TmpChar = "]" Then TmpInBracketsCnt = TmpInBracketsCnt - 1
      TmpInBrackets = TmpInBracketsCnt > 0
    End If
  Next i

  ReplaceSkipQuotedSections = TmpResult
End Function


Private Sub TestReplaceSkipQuotedSectionsHelper(ByVal AStart As Long, _
                                                ByVal AStr As String, _
                                                ByVal ASubStr As String, _
                                                ByVal ARepStr As String, _
                                                ByVal AExpectedStr As String, _
                                                Optional ByVal ATestFilePathStr As String = vbNullString)
  Dim TmpAnswerStr As String
  Dim TmpPassFailStr As String

  TmpAnswerStr = ReplaceSkipQuotedSections(AStart, AStr, ASubStr, ARepStr)
  TmpPassFailStr = IIf(TmpAnswerStr = AExpectedStr, "PASS", "FAIL")

  TestLogLine AStart & " | " & AStr & " | " & ASubStr & " | " & ARepStr & " | " & TmpAnswerStr & " | " & AExpectedStr & " | " & TmpPassFailStr, ATestFilePathStr
End Sub

Private Sub TestReplaceSkipQuotedSections(Optional ByVal ATestFilePathStr As String = vbNullString)
  TestReplaceSkipQuotedSectionsHelper 1, "=IF(A1,""a,b"",2,3)", ",", ", ", "=IF(A1, ""a,b"", 2, 3)", ATestFilePathStr
  TestReplaceSkipQuotedSectionsHelper 1, "SUM([@[Revenue, Total]])", ",", ", ", "SUM([@[Revenue, Total]])", ATestFilePathStr
End Sub


Public Function ShtNameRequiresSingleQuotes(ByVal AShtNameStr As String) As Boolean

  Const kSpecialChars As String = " -',:[]()!&^%$#@{}=+<>?/\"
  Dim TmpChar As String
  Dim TmpCnt As Integer

  ShtNameRequiresSingleQuotes = False

  For TmpCnt = 1 To Len(AShtNameStr)
    TmpChar = Mid$(AShtNameStr, TmpCnt, 1)
    If InStr(kSpecialChars, TmpChar) > 0 Then
      ShtNameRequiresSingleQuotes = True
      Exit Function
    End If
  Next TmpCnt

End Function


Private Sub TestShtNameRequiresSingleQuotesHelper(ByVal AShtNameStr As String, ByVal AExpectedBool As Boolean, Optional ByVal ATestFilePathStr As String = vbNullString)
  Dim TmpAnswerBool As Boolean
  Dim TmpPassFailStr As String

  TmpAnswerBool = ShtNameRequiresSingleQuotes(AShtNameStr)
  TmpPassFailStr = IIf(TmpAnswerBool = AExpectedBool, "PASS", "FAIL")

  TestLogLine AShtNameStr & " | " & TmpAnswerBool & " | " & AExpectedBool & " | " & TmpPassFailStr, ATestFilePathStr
End Sub

Private Sub TestShtNameRequiresSingleQuotes(Optional ByVal ATestFilePathStr As String = vbNullString)
  ' False -- no character from kSpecialChars present.
  TestShtNameRequiresSingleQuotesHelper "Sheet1", False, ATestFilePathStr
  TestShtNameRequiresSingleQuotesHelper "Sheet123", False, ATestFilePathStr
  TestShtNameRequiresSingleQuotesHelper "Sheet_1", False, ATestFilePathStr 'Underscore is notably absent from kSpecialChars.
  TestShtNameRequiresSingleQuotesHelper "", False, ATestFilePathStr 'Empty name: loop never runs, default holds.

  ' True -- one case per character in kSpecialChars (" -',:[]()!&^%$#@{}=+<>?/\"),
  ' so a future edit to that constant can't silently drop coverage.
  TestShtNameRequiresSingleQuotesHelper "Sheet 1", True, ATestFilePathStr
  TestShtNameRequiresSingleQuotesHelper "Sheet-1", True, ATestFilePathStr
  TestShtNameRequiresSingleQuotesHelper "Sheet'1", True, ATestFilePathStr
  TestShtNameRequiresSingleQuotesHelper "Sheet,1", True, ATestFilePathStr
  TestShtNameRequiresSingleQuotesHelper "Sheet:1", True, ATestFilePathStr
  TestShtNameRequiresSingleQuotesHelper "Sheet[1", True, ATestFilePathStr
  TestShtNameRequiresSingleQuotesHelper "Sheet]1", True, ATestFilePathStr
  TestShtNameRequiresSingleQuotesHelper "Sheet(1", True, ATestFilePathStr
  TestShtNameRequiresSingleQuotesHelper "Sheet)1", True, ATestFilePathStr
  TestShtNameRequiresSingleQuotesHelper "Sheet!1", True, ATestFilePathStr
  TestShtNameRequiresSingleQuotesHelper "Sheet&1", True, ATestFilePathStr
  TestShtNameRequiresSingleQuotesHelper "Sheet^1", True, ATestFilePathStr
  TestShtNameRequiresSingleQuotesHelper "Sheet%1", True, ATestFilePathStr
  TestShtNameRequiresSingleQuotesHelper "Sheet$1", True, ATestFilePathStr
  TestShtNameRequiresSingleQuotesHelper "Sheet#1", True, ATestFilePathStr
  TestShtNameRequiresSingleQuotesHelper "Sheet@1", True, ATestFilePathStr
  TestShtNameRequiresSingleQuotesHelper "Sheet{1", True, ATestFilePathStr
  TestShtNameRequiresSingleQuotesHelper "Sheet}1", True, ATestFilePathStr
  TestShtNameRequiresSingleQuotesHelper "Sheet=1", True, ATestFilePathStr
  TestShtNameRequiresSingleQuotesHelper "Sheet+1", True, ATestFilePathStr
  TestShtNameRequiresSingleQuotesHelper "Sheet<1", True, ATestFilePathStr
  TestShtNameRequiresSingleQuotesHelper "Sheet>1", True, ATestFilePathStr
  TestShtNameRequiresSingleQuotesHelper "Sheet?1", True, ATestFilePathStr
  TestShtNameRequiresSingleQuotesHelper "Sheet/1", True, ATestFilePathStr
  TestShtNameRequiresSingleQuotesHelper "Sheet\1", True, ATestFilePathStr
End Sub


Public Function ShtFormulaNameStr(ByVal AShtNameStr As String) As String
  If ShtNameRequiresSingleQuotes(AShtNameStr) Then
    ShtFormulaNameStr = "'" & AShtNameStr & "'"
  Else
    ShtFormulaNameStr = AShtNameStr
  End If
End Function


Private Sub TestShtFormulaNameStrHelper(ByVal AShtNameStr As String, ByVal AExpectedStr As String, Optional ByVal ATestFilePathStr As String = vbNullString)
  Dim TmpAnswerStr As String
  Dim TmpPassFailStr As String

  TmpAnswerStr = ShtFormulaNameStr(AShtNameStr)
  TmpPassFailStr = IIf(TmpAnswerStr = AExpectedStr, "PASS", "FAIL")

  TestLogLine AShtNameStr & " | " & TmpAnswerStr & " | " & AExpectedStr & " | " & TmpPassFailStr, ATestFilePathStr
End Sub

Private Sub TestShtFormulaNameStr(Optional ByVal ATestFilePathStr As String = vbNullString)
  TestShtFormulaNameStrHelper "Sheet1", "Sheet1", ATestFilePathStr
  TestShtFormulaNameStrHelper "Sheet 1", "'Sheet 1'", ATestFilePathStr
End Sub


Public Function ShortenFormula(ByVal AStr As String) As String
  Dim TmpCnt As Long
  Dim TmpStr As String
  Dim TmpChar As String
  Dim TmpInQuotes As Boolean
  Dim TmpInSingleQuotes As Boolean
  Dim TmpInBrackets As Boolean
  Dim TmpInBracketsCnt As Long

  ShortenFormula = vbNullString
  TmpInQuotes = False
  TmpInBrackets = False
  TmpInBracketsCnt = 0
  For TmpCnt = 1 To Len(AStr)
    TmpChar = Mid$(AStr, TmpCnt, 1)

    If TmpInQuotes Then 'Always add chars between quotes.
      TmpStr = TmpStr + TmpChar

    ElseIf TmpInSingleQuotes Then 'Always add chars between single quotes (tab names).
      TmpStr = TmpStr + TmpChar

    ElseIf TmpInBrackets Then
      TmpStr = TmpStr + TmpChar 'Brackets Represent Table Column Names that can have spaces.
    Else
      If (TmpChar <> " ") And (TmpChar <> Chr(10)) And (TmpChar <> Chr(13)) Then
        TmpStr = TmpStr + TmpChar
      End If
    End If

    If TmpChar = """" Then
      TmpInQuotes = Not TmpInQuotes
    End If

    If Not TmpInQuotes Then
      If TmpChar = "'" Then
        TmpInSingleQuotes = Not TmpInSingleQuotes
      End If

      If TmpChar = "[" Then
        TmpInBracketsCnt = TmpInBracketsCnt + 1
      End If

      If TmpChar = "]" Then
        TmpInBracketsCnt = TmpInBracketsCnt - 1
      End If

      TmpInBrackets = TmpInBracketsCnt > 0
    End If

  Next TmpCnt

  ShortenFormula = TmpStr
End Function


Private Sub TestShortenFormulaHelper(ByVal AStr As String, ByVal AExpectedStr As String, Optional ByVal ATestFilePathStr As String = vbNullString)
  Dim TmpAnswerStr As String
  Dim TmpPassFailStr As String

  TmpAnswerStr = ShortenFormula(AStr)
  TmpPassFailStr = IIf(TmpAnswerStr = AExpectedStr, "PASS", "FAIL")

  TestLogLine AStr & " | " & TmpAnswerStr & " | " & AExpectedStr & " | " & TmpPassFailStr, ATestFilePathStr
End Sub

Private Sub TestShortenFormula(Optional ByVal ATestFilePathStr As String = vbNullString)
  TestShortenFormulaHelper "=SUM( A1 , B1 )", "=SUM(A1,B1)", ATestFilePathStr

  'A trailing CRLF-CRLF (e.g. blank lines left over from a previous Pretty
  'Print) must be fully stripped, not just its LF half -- leaving orphaned
  'CRs behind previously survived into PrettyPrint's output uncaught.
  TestShortenFormulaHelper "=SUM(A1,B1)" & vbCrLf & vbCrLf, "=SUM(A1,B1)", ATestFilePathStr
End Sub


Public Sub ClearImmediateWindow()
  If Not Application.VBE.MainWindow.Visible Then Exit Sub
  Application.VBE.Windows("Immediate").SetFocus
  DoEvents
  Application.SendKeys "^g^a{DEL}", True
  DoEvents
End Sub


Public Function SpecialCellsSafe(ByVal ARng As Range, _
                                 ByVal ACellType As XlCellType, _
                                 Optional ByVal AValue As Variant) As Range

  Set SpecialCellsSafe = Nothing

  If ARng Is Nothing Then Exit Function

  'Error triggered when no cells found, so don't log error.
  On Error GoTo OnError
  If IsMissing(AValue) Then
    Set SpecialCellsSafe = ARng.SpecialCells(ACellType)
  Else
    Set SpecialCellsSafe = ARng.SpecialCells(ACellType, AValue)
  End If
  On Error GoTo 0

Finally:

  Exit Function
OnError:
  err.Clear
  Resume Finally

End Function


Public Sub SpecialCellsSafeTest()
  Dim rngToSearch As Range

  On Error GoTo OnError
  Set rngToSearch = SpecialCellsSafe(ActiveSheet.UsedRange, _
                                     XlCellType.xlCellTypeConstants, _
                                     XlSpecialCellsValue.xlTextValues)

  Debug.Print "Range: " + rngToSearch.Address

Finally:
  Set rngToSearch = Nothing
  Exit Sub

OnError:

  Debug.Print "SpecialCellsSafeTest: " + CStr(err.Number) + err.Description
  GoTo Finally
End Sub

Public Function AnySheetsProtected(ByVal AWkBook As Workbook) As Boolean
  Dim TmpWorksheet As Worksheet

  AnySheetsProtected = False

  For Each TmpWorksheet In AWkBook.Worksheets
    If TmpWorksheet.ProtectContents = True Then
      AnySheetsProtected = True
      Exit Function
    End If
  Next TmpWorksheet
End Function

Function GetFolder(AStartFolderStr As String, Optional ByVal ATitleStr As Variant) As String
    Dim Fldr As FileDialog
    Dim sItem As String
    Set Fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With Fldr
        If IsMissing(ATitleStr) Then
            .Title = "Select a Folder"
        Else
            .Title = ATitleStr
        End If
        
        .AllowMultiSelect = False

        If (AStartFolderStr <> vbNullString) And DirExists(AStartFolderStr) Then
          If Right(AStartFolderStr, 1) = "\" Then
            AStartFolderStr = Left(AStartFolderStr, Len(AStartFolderStr) - 1)
          End If
          .InitialFileName = AStartFolderStr & "\"
        Else
          .InitialFileName = Application.DefaultFilePath & "\"
        End If

        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
NextCode:
    GetFolder = sItem
    Set Fldr = Nothing
End Function


Public Sub OpenPathInExplorer(ByVal APathStr As String)
  Dim TmpStr As String
  
  If APathStr = vbNullString Then
    MsgBox "Path Not Provided.", vbExclamation
    Exit Sub
  End If
  
  TmpStr = "Explorer.exe " & """" & APathStr & """"
  
  Shell TmpStr, vbNormalFocus
End Sub


