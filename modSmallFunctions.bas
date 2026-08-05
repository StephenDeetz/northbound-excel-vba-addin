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


Private Sub TestStrCnt()
  Debug.Print StrCnt("/DDDD//", "//")
End Sub


Public Function RemoveQuotedSections(ByVal AStr As String) As String
  Dim TmpResult As String
  Dim TmpInsideQuotes As Boolean
  Dim i As Long

  TmpResult = vbNullString
  TmpInsideQuotes = False

  For i = 1 To Len(AStr)
    If Mid(AStr, i, 1) = """" Then
      TmpInsideQuotes = Not TmpInsideQuotes
    ElseIf Not TmpInsideQuotes Then
      TmpResult = TmpResult & Mid(AStr, i, 1)
    End If
  Next i

  RemoveQuotedSections = TmpResult
End Function


Private Sub TestRemoveQuotedSections()
  Debug.Print RemoveQuotedSections("Test(" & """This is the middle""" & ")")
End Sub


Public Function PosSkipQuotedSections(ByVal AStart As Long, ByVal AStr As String, ByVal ASubStr As String) As Long
  Dim TmpInsideQuotes As Boolean
  Dim i As Long

  PosSkipQuotedSections = 0

  TmpInsideQuotes = False

  For i = AStart To Len(AStr)
    If Mid(AStr, i, 1) = """" Then
      TmpInsideQuotes = Not TmpInsideQuotes
    ElseIf Not TmpInsideQuotes Then
      If Mid$(AStr, i, Len(ASubStr)) = ASubStr Then
        PosSkipQuotedSections = i
        Exit Function
      End If
    End If
  Next i

End Function


Private Sub TestPosSkipQuotedSections()
  Debug.Print PosSkipQuotedSections(1, "=IF(A1=""("",1,2)", ")")
End Sub


Public Function ReplaceSkipQuotedSections(ByVal AStart As Long, _
                                          ByVal AStr As String, _
                                          ByVal ASubStr As String, _
                                          ByVal ARepStr As String) As String
  Dim TmpInsideQuotes As Boolean
  Dim i As Long
  Dim TmpResult As String

  TmpResult = vbNullString
  TmpInsideQuotes = False

  For i = AStart To Len(AStr)
    If Mid$(AStr, i, 1) = """" Then
      TmpInsideQuotes = Not TmpInsideQuotes
      TmpResult = TmpResult & Mid(AStr, i, 1) ' Always include quotes
    ElseIf Not TmpInsideQuotes Then
      If Mid$(AStr, i, Len(ASubStr)) = ASubStr Then
        ' Append up to the current position before skipping the substring
        TmpResult = TmpResult + ARepStr
      Else
        TmpResult = TmpResult & Mid$(AStr, i, 1)
      End If
    Else
      TmpResult = TmpResult & Mid(AStr, i, 1)
    End If
  Next i

  ReplaceSkipQuotedSections = TmpResult
End Function


Private Sub TestReplaceSkipQuotedSections()
  Debug.Print ReplaceSkipQuotedSections(1, "=IF(A1,""a,b"",2,3)", ",", ", ")
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


Private Sub TestShtNameRequiresSingleQuotes()
  Debug.Print ShtNameRequiresSingleQuotes("Sheet1")        ' Should be False
  Debug.Print ShtNameRequiresSingleQuotes("Sheet 1")       ' Should be True
End Sub


Public Function ShtFormulaNameStr(ByVal AShtNameStr As String) As String
  If ShtNameRequiresSingleQuotes(AShtNameStr) Then
    ShtFormulaNameStr = "'" & AShtNameStr & "'"
  Else
    ShtFormulaNameStr = AShtNameStr
  End If
End Function


Private Sub TestShtFormulaNameStr()
  Debug.Print ShtFormulaNameStr("Sheet1")        ' Should be Sheet1
  Debug.Print ShtFormulaNameStr("Sheet 1")       ' Should be 'Sheet 1'
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
      If (TmpChar <> " ") And (TmpChar <> Chr(10)) Then
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
