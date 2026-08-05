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


Private Sub TestStrCntHelper(ByVal AStr As String, ByVal ASubStr As String, ByVal AExpectedCnt As Long)
  Dim TmpAnswerCnt As Long
  Dim TmpPassFailStr As String

  TmpAnswerCnt = StrCnt(AStr, ASubStr)
  TmpPassFailStr = IIf(TmpAnswerCnt = AExpectedCnt, "PASS", "FAIL")

  Debug.Print AStr; " | "; ASubStr; " | "; TmpAnswerCnt; " | "; AExpectedCnt; " | "; TmpPassFailStr
End Sub

Private Sub TestStrCnt()
  TestStrCntHelper "/DDDD//", "//", 1
End Sub


Public Function RemoveQuotedSections(ByVal AStr As String) As String
  Dim TmpResult As String
  Dim TmpInsideQuotes As Boolean
  Dim TmpInBracketsCnt As Long
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
    End If

    If Not TmpIsDelimiter And Not TmpInsideQuotes And TmpInBracketsCnt = 0 Then
      TmpResult = TmpResult & TmpChar
    End If
  Next i

  RemoveQuotedSections = TmpResult
End Function


Private Sub TestRemoveQuotedSectionsHelper(ByVal AOriginalStr As String, ByVal AExpectedStr As String)
  Dim TmpAnswerStr As String
  Dim TmpPassFailStr As String

  TmpAnswerStr = RemoveQuotedSections(AOriginalStr)
  TmpPassFailStr = IIf(TmpAnswerStr = AExpectedStr, "PASS", "FAIL")

  Debug.Print AOriginalStr; " | "; TmpAnswerStr; " | "; AExpectedStr; " | "; TmpPassFailStr
End Sub

Private Sub TestRemoveQuotedSections()
  TestRemoveQuotedSectionsHelper "Test(" & """This is the middle""" & ")", "Test()"
  TestRemoveQuotedSectionsHelper "SUM([@[Revenue, Total]])", "SUM()"
  TestRemoveQuotedSectionsHelper "A,[Col1],B", "A,,B"
  TestRemoveQuotedSectionsHelper "[Col1],[Col2]", ","
  TestRemoveQuotedSectionsHelper """[literal],text""" & ",A2", ",A2"
  TestRemoveQuotedSectionsHelper "[]", vbNullString
End Sub


Public Function PosSkipQuotedSections(ByVal AStart As Long, ByVal AStr As String, ByVal ASubStr As String) As Long
  Dim TmpInsideQuotes As Boolean
  Dim TmpInBrackets As Boolean
  Dim TmpInBracketsCnt As Long
  Dim TmpChar As String
  Dim i As Long

  PosSkipQuotedSections = 0

  TmpInsideQuotes = False

  For i = AStart To Len(AStr)
    TmpChar = Mid$(AStr, i, 1)

    If TmpChar = """" Then
      TmpInsideQuotes = Not TmpInsideQuotes
    ElseIf Not TmpInsideQuotes And Not TmpInBrackets Then
      If Mid$(AStr, i, Len(ASubStr)) = ASubStr Then
        PosSkipQuotedSections = i
        Exit Function
      End If
    End If

    If Not TmpInsideQuotes Then
      If TmpChar = "[" Then TmpInBracketsCnt = TmpInBracketsCnt + 1
      If TmpChar = "]" Then TmpInBracketsCnt = TmpInBracketsCnt - 1
      TmpInBrackets = TmpInBracketsCnt > 0
    End If
  Next i

End Function


Private Sub TestPosSkipQuotedSectionsHelper(ByVal AStart As Long, ByVal AStr As String, ByVal ASubStr As String, ByVal AExpectedPos As Long)
  Dim TmpAnswerPos As Long
  Dim TmpPassFailStr As String

  TmpAnswerPos = PosSkipQuotedSections(AStart, AStr, ASubStr)
  TmpPassFailStr = IIf(TmpAnswerPos = AExpectedPos, "PASS", "FAIL")

  Debug.Print AStart; " | "; AStr; " | "; ASubStr; " | "; TmpAnswerPos; " | "; AExpectedPos; " | "; TmpPassFailStr
End Sub

Private Sub TestPosSkipQuotedSections()
  TestPosSkipQuotedSectionsHelper 1, "=IF(A1=""("",1,2)", ")", 15
  TestPosSkipQuotedSectionsHelper 1, "SUM([@[Revenue, Total]])", ",", 0 'Comma is inside brackets.
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
                                                ByVal AExpectedStr As String)
  Dim TmpAnswerStr As String
  Dim TmpPassFailStr As String

  TmpAnswerStr = ReplaceSkipQuotedSections(AStart, AStr, ASubStr, ARepStr)
  TmpPassFailStr = IIf(TmpAnswerStr = AExpectedStr, "PASS", "FAIL")

  Debug.Print AStart; " | "; AStr; " | "; ASubStr; " | "; ARepStr; " | "; TmpAnswerStr; " | "; AExpectedStr; " | "; TmpPassFailStr
End Sub

Private Sub TestReplaceSkipQuotedSections()
  TestReplaceSkipQuotedSectionsHelper 1, "=IF(A1,""a,b"",2,3)", ",", ", ", "=IF(A1, ""a,b"", 2, 3)"
  TestReplaceSkipQuotedSectionsHelper 1, "SUM([@[Revenue, Total]])", ",", ", ", "SUM([@[Revenue, Total]])"
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


Private Sub TestShtNameRequiresSingleQuotesHelper(ByVal AShtNameStr As String, ByVal AExpectedBool As Boolean)
  Dim TmpAnswerBool As Boolean
  Dim TmpPassFailStr As String

  TmpAnswerBool = ShtNameRequiresSingleQuotes(AShtNameStr)
  TmpPassFailStr = IIf(TmpAnswerBool = AExpectedBool, "PASS", "FAIL")

  Debug.Print AShtNameStr; " | "; TmpAnswerBool; " | "; AExpectedBool; " | "; TmpPassFailStr
End Sub

Private Sub TestShtNameRequiresSingleQuotes()
  TestShtNameRequiresSingleQuotesHelper "Sheet1", False
  TestShtNameRequiresSingleQuotesHelper "Sheet 1", True
End Sub


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

  Debug.Print AShtNameStr; " | "; TmpAnswerStr; " | "; AExpectedStr; " | "; TmpPassFailStr
End Sub

Private Sub TestShtFormulaNameStr()
  TestShtFormulaNameStrHelper "Sheet1", "Sheet1"
  TestShtFormulaNameStrHelper "Sheet 1", "'Sheet 1'"
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

