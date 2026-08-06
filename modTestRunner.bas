Attribute VB_Name = "modTestRunner"
Option Explicit

' ==========================================================================================
' Module: modTestRunner
'
' Purpose:
'   Orchestrates every Test* sub across this add-in's modules into one run,
'   writing PASS/FAIL results to a shared file instead of the Immediate
'   window, so failures can be found with a single text search instead of
'   running and re-reading each Test* sub by hand.
'
' Usage:
'   RunAllTests                          ' runs every wired-in test, writes
'                                         ' TestResults.txt to the remembered
'                                         ' test folder (prompts once, then
'                                         ' remembers it via NBPub_TestDir.txt)
'
'   Any individual Test* sub still works standalone exactly as before -- e.g.
'   running TestPrettyPrint directly from the Immediate window omits the
'   optional ATestFilePathStr argument, which falls back to Debug.Print.
' ==========================================================================================

Private Const kTestDirPathFile As String = "NBPub_TestDir.txt"
Private Const kTestResultsFileName As String = "TestResults.txt"

Private Const kGoldenDirPathFile As String = "NBPub_GoldenDir.txt"
Private Const kGoldenFileName As String = "PrettyPrintGolden.txt"

'Writes ALine to ATestFilePathStr if given, else to the Immediate window.
Public Sub TestLogLine(ByVal ALine As String, Optional ByVal ATestFilePathStr As String = vbNullString)
  Dim TmpFileNbr As Integer

  If LenB(ATestFilePathStr) = 0 Then
    Debug.Print ALine
    Exit Sub
  End If

  TmpFileNbr = FreeFile
  Open ATestFilePathStr For Append As #TmpFileNbr
  Print #TmpFileNbr, ALine
  Close #TmpFileNbr
End Sub

'Remembered test-results folder (same pattern as AddinExportCodeModules),
'plus the results filename. Returns vbNullString if the user cancels.
Private Function ResolveTestLogPathStr() As String
  Dim TmpDirStr As String

  ResolveTestLogPathStr = vbNullString

  TmpDirStr = AppDataFileReadStr(kTestDirPathFile)
  If LenB(TmpDirStr) = 0 Then TmpDirStr = GetFolder(ThisWorkbook.Path)
  If LenB(TmpDirStr) = 0 Then Exit Function

  AppDataFileWriteStr TmpDirStr, kTestDirPathFile
  ResolveTestLogPathStr = IncludeTrailingBackslash(TmpDirStr) & kTestResultsFileName
End Function

'Remembered golden-file folder -- a separate remembered path from
'ResolveTestLogPathStr's, since this one must point at the git repo (the
'golden file is a version-controlled fixture), not a scratch results folder.
Public Function ResolveGoldenFilePathStr() As String
  Dim TmpDirStr As String

  ResolveGoldenFilePathStr = vbNullString

  TmpDirStr = AppDataFileReadStr(kGoldenDirPathFile)
  If LenB(TmpDirStr) = 0 Then TmpDirStr = GetFolder(ThisWorkbook.Path)
  If LenB(TmpDirStr) = 0 Then Exit Function

  AppDataFileWriteStr TmpDirStr, kGoldenDirPathFile
  ResolveGoldenFilePathStr = IncludeTrailingBackslash(TmpDirStr) & kGoldenFileName
End Function

'Logs a section header, runs ATestNameStr via Application.Run (works on
'Private subs in other modules), and traps any crash so one bad test can't
'halt the rest of the run or leave Excel state stuck mid-test.
Private Sub RunOneTest(ByVal ATestNameStr As String, ByVal ATestFilePathStr As String)
  TestLogLine "===== " & ATestNameStr & " =====", ATestFilePathStr

  On Error Resume Next
  Application.Run ATestNameStr, ATestFilePathStr
  If err.Number <> 0 Then
    TestLogLine "CRASHED: " & ATestNameStr & " - " & err.Description, ATestFilePathStr
    err.Clear
  End If
  On Error GoTo 0

  TestLogLine vbNullString, ATestFilePathStr 'Blank separator line between sections.
End Sub

Public Sub RunAllTests()
  Dim TmpPathStr As String

  TmpPathStr = ResolveTestLogPathStr()
  If LenB(TmpPathStr) > 0 Then
    If Dir$(TmpPathStr) <> vbNullString Then Kill TmpPathStr
  End If

  ' modSmallFunctions.bas
  RunOneTest "TestShortenFormula", TmpPathStr
  RunOneTest "TestStrCnt", TmpPathStr
  RunOneTest "TestRemoveQuotedSections", TmpPathStr
  RunOneTest "TestPosSkipQuotedSections", TmpPathStr
  RunOneTest "TestReplaceSkipQuotedSections", TmpPathStr
  RunOneTest "TestShtNameRequiresSingleQuotes", TmpPathStr
  RunOneTest "TestShtFormulaNameStr", TmpPathStr

  ' modAddOperatorWhiteSpace.bas
  RunOneTest "TestIsMultiOperator", TmpPathStr
  RunOneTest "TestIsSingleOperator", TmpPathStr
  RunOneTest "TestAddOperatorWhiteSpace", TmpPathStr
  RunOneTest "TestAddOperatorWhiteSpaceExact", TmpPathStr
  RunOneTest "TestAddOperatorWhiteSpaceEdgeCases", TmpPathStr
  ' TestAddOperatorWhiteSpaceDoubleNeg -- excluded: different, non-standardized
  ' output format, with documented known-failing edge cases.

  ' modPrettyPrint.bas
  RunOneTest "TestPrettyPrintWouldChangeFormula", TmpPathStr
  RunOneTest "TestPrettyPrintExact", TmpPathStr
  RunOneTest "TestGoldenEscapeStr", TmpPathStr
  RunOneTest "TestGoldenReadEntries", TmpPathStr
  RunOneTest "TestPrettyPrintGolden", TmpPathStr
  ' TestPrettyPrint -- excluded: no PASS/FAIL, just Input/Answer for manual
  ' review; run it standalone from the Immediate window instead.
  ' TestPrettyPrintCaptureGolden -- excluded: deliberately mutates the golden
  ' file, never run automatically.
  ' TestCellSetFormula2Safe -- excluded: mutates ActiveSheet.Range("A1").

  If LenB(TmpPathStr) > 0 Then
    MsgBox "Results written to:" & vbCrLf & TmpPathStr
  End If
End Sub
