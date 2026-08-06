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
'   running TestPrettyPrint directly from the Immediate window (or via F5 in
'   the VBE, since Test* subs take no arguments) falls back to Debug.Print.
'   TestLogLine decides Debug.Print vs. file-append by checking TestFilePathStr,
'   a module-level flag RunAllTests sets for the duration of its run -- Test*
'   subs used to take this as an optional argument, but that made VBA treat
'   them as "not runnable with F5" (VBA's F5/Run only works parameter-free).
' ==========================================================================================

Private Const kTestDirPathFile As String = "NBPub_TestDir.txt"
Private Const kTestResultsFileName As String = "TestResults.txt"

Private Const kGoldenDirPathFile As String = "NBPub_GoldenDir.txt"
Private Const kGoldenFileName As String = "PrettyPrintGolden.txt"

Private TestFilePathStr As String 'Set by RunAllTests for its duration; vbNullString otherwise.

'Writes ALine to the file RunAllTests is currently targeting, or the
'Immediate window if this Test* sub is running standalone.
Public Sub TestLogLine(ByVal ALine As String)
  Dim TmpFileNbr As Integer

  If LenB(TestFilePathStr) = 0 Then
    Debug.Print ALine
    Exit Sub
  End If

  TmpFileNbr = FreeFile
  Open TestFilePathStr For Append As #TmpFileNbr
  Print #TmpFileNbr, ALine
  Close #TmpFileNbr
End Sub

'Remembered test-results folder (same pattern as AddinExportCodeModules),
'plus the results filename. Returns vbNullString if the user cancels.
Private Function ResolveTestLogPathStr() As String
  Dim TmpDirStr As String

  ResolveTestLogPathStr = vbNullString

  TmpDirStr = AppDataFileReadStr(kTestDirPathFile)
  If LenB(TmpDirStr) = 0 Then TmpDirStr = GetFolder(ActiveWorkbook.Path)
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
  If LenB(TmpDirStr) = 0 Then TmpDirStr = GetFolder(ActiveWorkbook.Path)
  If LenB(TmpDirStr) = 0 Then Exit Function

  AppDataFileWriteStr TmpDirStr, kGoldenDirPathFile
  ResolveGoldenFilePathStr = IncludeTrailingBackslash(TmpDirStr) & kGoldenFileName
End Function

'Logs a section header, runs ATestNameStr via Application.Run (works on
'Private subs in other modules), and traps any crash so one bad test can't
'halt the rest of the run or leave Excel state stuck mid-test.
Private Sub RunOneTest(ByVal ATestNameStr As String)
  TestLogLine "===== " & ATestNameStr & " ====="

  On Error Resume Next
  Application.Run ATestNameStr
  If err.Number <> 0 Then
    TestLogLine "CRASHED: " & ATestNameStr & " - " & err.Description
    err.Clear
  End If
  On Error GoTo 0

  TestLogLine vbNullString 'Blank separator line between sections.
End Sub

Public Sub RunAllTests()
  Dim TmpPathStr As String

  TmpPathStr = ResolveTestLogPathStr()
  If LenB(TmpPathStr) > 0 Then
    If Dir$(TmpPathStr) <> vbNullString Then Kill TmpPathStr
  End If

  TestFilePathStr = TmpPathStr
  On Error GoTo Finally

  ' modSmallFunctions.bas
  RunOneTest "TestShortenFormula"
  RunOneTest "TestStrCnt"
  RunOneTest "TestRemoveQuotedSections"
  RunOneTest "TestPosSkipQuotedSections"
  RunOneTest "TestReplaceSkipQuotedSections"
  RunOneTest "TestShtNameRequiresSingleQuotes"
  RunOneTest "TestShtFormulaNameStr"

  ' modMacroSpeedup.bas
  RunOneTest "TestMacroSpeedup"
  RunOneTest "Test_MacroSpeedup_AllItemsPushPop"
  RunOneTest "Test_MacroSpeedup_NestedDepth"
  RunOneTest "Test_MacroSpeedup_ManualBaseline"
  RunOneTest "Test_MacroSpeedup_Interleaved"
  RunOneTest "Test_MacroSpeedup_IdempotentPush"
  RunOneTest "Test_MacroSpeedup_MidScopeOverride"
  RunOneTest "Test_MacroSpeedup_ErrorPathFinalizer"

  ' modAddOperatorWhiteSpace.bas
  RunOneTest "TestIsMultiOperator"
  RunOneTest "TestIsSingleOperator"
  RunOneTest "TestAddOperatorWhiteSpace"
  RunOneTest "TestAddOperatorWhiteSpaceExact"
  RunOneTest "TestAddOperatorWhiteSpaceEdgeCases"
  ' TestAddOperatorWhiteSpaceDoubleNeg -- excluded: different, non-standardized
  ' output format, with documented known-failing edge cases.

  ' modPrettyPrint.bas
  RunOneTest "TestPrettyPrintWouldChangeFormula"
  RunOneTest "TestPrettyPrintExact"
  RunOneTest "TestGoldenEscapeStr"
  RunOneTest "TestGoldenReadEntries"
  RunOneTest "TestPrettyPrintGolden"
  ' TestPrettyPrint -- excluded: no PASS/FAIL, just Input/Answer for manual
  ' review; run it standalone from the Immediate window instead.
  ' TestPrettyPrintCaptureGolden -- excluded: deliberately mutates the golden
  ' file, never run automatically.
  ' TestCellSetFormula2Safe -- excluded: mutates ActiveSheet.Range("A1").

Finally:
  TestFilePathStr = vbNullString 'Standalone Test* runs afterward fall back to Debug.Print.

  If LenB(TmpPathStr) > 0 Then
    MsgBox "Results written to:" & vbCrLf & TmpPathStr
  End If
End Sub
