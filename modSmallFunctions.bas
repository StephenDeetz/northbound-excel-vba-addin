Attribute VB_Name = "modSmallFunctions"
' Copyright (c) 2026 Northbound Group
' Contact: stephendeetz@northboundgroup.com
' SPDX-License-Identifier: MIT
Option Explicit
' Shared low-level string/formula helpers used by the Pretty Print feature
' (modPrettyPrint.bas, modAddOperatorWhiteSpace.bas).
'
' Public declarations (in source order):
'   Inc
'   Dec
'   StrCnt
'   ShtNameRequiresSingleQuotes
'   ShtFormulaNameStr
'   ShortenFormula
'   ClearImmediateWindow
'   SpecialCellsSafe

' Purpose: Increment a numeric variable by 1.
Public Sub Inc(ByRef AVal As Variant)
  AVal = AVal + 1
End Sub

' Purpose: Decrement a numeric variable by 1.
Public Sub Dec(ByRef AVal As Variant)
  AVal = AVal - 1
End Sub

' Purpose: Return the current user's AppData directory path.
Public Function GetAppDataDir() As String
  GetAppDataDir = Environ$("AppData")
End Function

' Purpose: Return True if the given path exists and is a directory.
Public Function DirExists(ByVal AFileNameStr As String) As Boolean
  On Error Resume Next
  DirExists = ((GetAttr(AFileNameStr) And vbDirectory) = vbDirectory)
End Function

' Purpose: Append a trailing backslash to a directory path if it lacks one.
Public Function IncludeTrailingBackslash(ByVal ADirStr As String) As String
  If Right$(ADirStr, 1) <> "\" Then
    IncludeTrailingBackslash = ADirStr & "\"
  Else
    IncludeTrailingBackslash = ADirStr
  End If
End Function

' Purpose: Count non-overlapping occurrences of a substring within a string.
Public Function StrCnt(ByVal AStr As String, ByVal ASubStr As String) As Long
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


Private Sub TestStrCntHelper(ByVal AStr As String, ByVal ASubStr As String, ByVal AExpectedCnt As Long)
  Dim TmpAnswerCnt As Long
  Dim TmpPassFailStr As String

  TmpAnswerCnt = StrCnt(AStr, ASubStr)
  TmpPassFailStr = IIf(TmpAnswerCnt = AExpectedCnt, "PASS", "FAIL")

  TestLogLine AStr & " | " & ASubStr & " | " & TmpAnswerCnt & " | " & AExpectedCnt & " | " & TmpPassFailStr
End Sub

Private Sub TestStrCnt()
  If IsStandaloneTestRun() Then ClearImmediateWindow
  TestStrCntHelper "/DDDD//", "//", 1
End Sub


' Purpose: Return True if a sheet name requires single-quoting in a formula reference.
Public Function ShtNameRequiresSingleQuotes(ByVal AShtNameStr As String) As Boolean

  Const kSpecialChars As String = " -',:[]()!&^%$#@{}=+<>?/\"
  Dim TmpChar As String
  Dim TmpCnt As Long

  ShtNameRequiresSingleQuotes = False

  For TmpCnt = 1 To Len(AShtNameStr)
    TmpChar = Mid$(AShtNameStr, TmpCnt, 1)
    If InStr(kSpecialChars, TmpChar) > 0 Then
      ShtNameRequiresSingleQuotes = True
      Exit Function
    End If
  Next TmpCnt

End Function


Private Sub TestShtNameRequiresSingleQuotesHelper(ByVal AShtNameStr As String, ByVal AExpectedBool As Boolean)
  Dim TmpAnswerBool As Boolean
  Dim TmpPassFailStr As String

  TmpAnswerBool = ShtNameRequiresSingleQuotes(AShtNameStr)
  TmpPassFailStr = IIf(TmpAnswerBool = AExpectedBool, "PASS", "FAIL")

  TestLogLine AShtNameStr & " | " & TmpAnswerBool & " | " & AExpectedBool & " | " & TmpPassFailStr
End Sub

Private Sub TestShtNameRequiresSingleQuotes()
  If IsStandaloneTestRun() Then ClearImmediateWindow

  ' False -- no character from kSpecialChars present.
  TestShtNameRequiresSingleQuotesHelper "Sheet1", False
  TestShtNameRequiresSingleQuotesHelper "Sheet123", False
  TestShtNameRequiresSingleQuotesHelper "Sheet_1", False 'Underscore is notably absent from kSpecialChars.
  TestShtNameRequiresSingleQuotesHelper "", False 'Empty name: loop never runs, default holds.

  ' True -- one case per character in kSpecialChars (" -',:[]()!&^%$#@{}=+<>?/\"),
  ' so a future edit to that constant can't silently drop coverage.
  TestShtNameRequiresSingleQuotesHelper "Sheet 1", True
  TestShtNameRequiresSingleQuotesHelper "Sheet-1", True
  TestShtNameRequiresSingleQuotesHelper "Sheet'1", True
  TestShtNameRequiresSingleQuotesHelper "Sheet,1", True
  TestShtNameRequiresSingleQuotesHelper "Sheet:1", True
  TestShtNameRequiresSingleQuotesHelper "Sheet[1", True
  TestShtNameRequiresSingleQuotesHelper "Sheet]1", True
  TestShtNameRequiresSingleQuotesHelper "Sheet(1", True
  TestShtNameRequiresSingleQuotesHelper "Sheet)1", True
  TestShtNameRequiresSingleQuotesHelper "Sheet!1", True
  TestShtNameRequiresSingleQuotesHelper "Sheet&1", True
  TestShtNameRequiresSingleQuotesHelper "Sheet^1", True
  TestShtNameRequiresSingleQuotesHelper "Sheet%1", True
  TestShtNameRequiresSingleQuotesHelper "Sheet$1", True
  TestShtNameRequiresSingleQuotesHelper "Sheet#1", True
  TestShtNameRequiresSingleQuotesHelper "Sheet@1", True
  TestShtNameRequiresSingleQuotesHelper "Sheet{1", True
  TestShtNameRequiresSingleQuotesHelper "Sheet}1", True
  TestShtNameRequiresSingleQuotesHelper "Sheet=1", True
  TestShtNameRequiresSingleQuotesHelper "Sheet+1", True
  TestShtNameRequiresSingleQuotesHelper "Sheet<1", True
  TestShtNameRequiresSingleQuotesHelper "Sheet>1", True
  TestShtNameRequiresSingleQuotesHelper "Sheet?1", True
  TestShtNameRequiresSingleQuotesHelper "Sheet/1", True
  TestShtNameRequiresSingleQuotesHelper "Sheet\1", True
End Sub


' Purpose: Return a sheet name formatted for use in a formula, single-quoted if needed.
Public Function ShtFormulaNameStr(ByVal AShtNameStr As String) As String
  If ShtNameRequiresSingleQuotes(AShtNameStr) Then
    ShtFormulaNameStr = "'" & AShtNameStr & "'"
  Else
    ShtFormulaNameStr = AShtNameStr
  End If
End Function


Private Sub TestShtFormulaNameStrHelper(ByVal AShtNameStr As String, ByVal AExpectedStr As String)
  Dim TmpAnswerStr As String
  Dim TmpPassFailStr As String

  TmpAnswerStr = ShtFormulaNameStr(AShtNameStr)
  TmpPassFailStr = IIf(TmpAnswerStr = AExpectedStr, "PASS", "FAIL")

  TestLogLine AShtNameStr & " | " & TmpAnswerStr & " | " & AExpectedStr & " | " & TmpPassFailStr
End Sub

Private Sub TestShtFormulaNameStr()
  If IsStandaloneTestRun() Then ClearImmediateWindow
  TestShtFormulaNameStrHelper "Sheet1", "Sheet1"
  TestShtFormulaNameStrHelper "Sheet 1", "'Sheet 1'"
End Sub


' Purpose: Strip unnecessary spaces and line breaks from a formula string outside quoted sections.
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


Private Sub TestShortenFormulaHelper(ByVal AStr As String, ByVal AExpectedStr As String)
  Dim TmpAnswerStr As String
  Dim TmpPassFailStr As String

  TmpAnswerStr = ShortenFormula(AStr)
  TmpPassFailStr = IIf(TmpAnswerStr = AExpectedStr, "PASS", "FAIL")

  TestLogLine AStr & " | " & TmpAnswerStr & " | " & AExpectedStr & " | " & TmpPassFailStr
End Sub

Private Sub TestShortenFormula()
  If IsStandaloneTestRun() Then ClearImmediateWindow
  TestShortenFormulaHelper "=SUM( A1 , B1 )", "=SUM(A1,B1)"

  'A trailing CRLF-CRLF (e.g. blank lines left over from a previous Pretty
  'Print) must be fully stripped, not just its LF half -- leaving orphaned
  'CRs behind previously survived into PrettyPrint's output uncaught.
  TestShortenFormulaHelper "=SUM(A1,B1)" & vbCrLf & vbCrLf, "=SUM(A1,B1)"
End Sub


' Purpose: Clear all text from the VBE Immediate window.
Public Sub ClearImmediateWindow()
  Const kImmediateWindowMaxLines  As Long = 200
  Dim TmpCnt As Long
  
  If Not Application.VBE.MainWindow.Visible Then Exit Sub
  If Not Application.VBE.Windows("Immediate").Visible Then Exit Sub

  For TmpCnt = 1 To kImmediateWindowMaxLines
    Debug.Print ""
  Next TmpCnt

End Sub


' Purpose: Wrap Range.SpecialCells, returning Nothing instead of raising when no cells match.
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


' Purpose: Manual smoke test for SpecialCellsSafe against the active sheet's used range.
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

' Purpose: Return True if any worksheet in the given workbook has protected contents.
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

' Purpose: Show a folder picker, returning the chosen path, or vbNullString if canceled.
Public Function GetFolder(ByVal AStartFolderStr As String, _
                          Optional ByVal ATitleStr As Variant) As String
  Dim TmpFldr As FileDialog
  Dim TmpItemStr As String

  GetFolder = vbNullString

  Set TmpFldr = Application.FileDialog(msoFileDialogFolderPicker)
  With TmpFldr
    If IsMissing(ATitleStr) Then
      .Title = "Select a Folder"
    Else
      .Title = ATitleStr
    End If

    .AllowMultiSelect = False

    If (AStartFolderStr <> vbNullString) And DirExists(AStartFolderStr) Then
      If Right$(AStartFolderStr, 1) = "\" Then
        AStartFolderStr = Left$(AStartFolderStr, Len(AStartFolderStr) - 1)
      End If
      .InitialFileName = AStartFolderStr & "\"
    Else
      .InitialFileName = Application.DefaultFilePath & "\"
    End If

    If .Show = -1 Then TmpItemStr = .SelectedItems(1)
  End With

  GetFolder = TmpItemStr
End Function


' Purpose: Open the given path in a Windows Explorer window.
Public Sub OpenPathInExplorer(ByVal APathStr As String)
  Dim TmpStr As String
  
  If APathStr = vbNullString Then
    MsgBox "Path Not Provided.", vbExclamation
    Exit Sub
  End If
  
  TmpStr = "Explorer.exe " & """" & APathStr & """"
  
  Shell TmpStr, vbNormalFocus
End Sub







