Attribute VB_Name = "modTestRunner"
' Copyright (c) 2026 Northbound Group
' Contact: stephendeetz@northboundgroup.com
' SPDX-License-Identifier: MIT
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

Private Const kAnswersDirPathFile As String = "NBPub_AnswersDir.txt"
Private Const kAnswersFileName As String = "PrettyPrintAnswers.txt"

Private TestFilePathStr As String 'Set by RunAllTests for its duration; vbNullString otherwise.

'True when no Test* sub is currently running under RunAllTests -- i.e. this
'Test* sub was launched standalone (F5 or Immediate window). Top-level Test*
'subs call ClearImmediateWindow when this is True, so standalone re-runs
'don't accumulate old Debug.Print output.
Public Function IsStandaloneTestRun() As Boolean
  IsStandaloneTestRun = (LenB(TestFilePathStr) = 0)
End Function

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

'Remembered answers-file folder -- a separate remembered path from
'ResolveTestLogPathStr's, since this one must point at the git repo (the
'answers file is a version-controlled fixture), not a scratch results folder.
Public Function ResolveAnswersFilePathStr() As String
  Dim TmpDirStr As String

  ResolveAnswersFilePathStr = vbNullString

  TmpDirStr = AppDataFileReadStr(kAnswersDirPathFile)
  If LenB(TmpDirStr) = 0 Then TmpDirStr = GetFolder(ActiveWorkbook.Path)
  If LenB(TmpDirStr) = 0 Then Exit Function

  AppDataFileWriteStr TmpDirStr, kAnswersDirPathFile
  ResolveAnswersFilePathStr = IncludeTrailingBackslash(TmpDirStr) & kAnswersFileName
End Function

'Counts result lines in APathStr by their trailing "| PASS" / "| FAIL" marker.
'(Header/section-title lines don't contain either substring, so they're
'skipped without needing a separate "is this a result line" check.)
Private Sub CountPassFail(ByVal APathStr As String, ByRef APassCnt As Long, ByRef AFailCnt As Long)
  Dim TmpFileNbr As Integer
  Dim TmpLine As String

  APassCnt = 0
  AFailCnt = 0
  If Dir$(APathStr) = vbNullString Then Exit Sub

  TmpFileNbr = FreeFile
  Open APathStr For Input As #TmpFileNbr
  Do Until EOF(TmpFileNbr)
    Line Input #TmpFileNbr, TmpLine
    If InStr(TmpLine, "| FAIL") > 0 Then
      AFailCnt = AFailCnt + 1
    ElseIf InStr(TmpLine, "| PASS") > 0 Then
      APassCnt = APassCnt + 1
    End If
  Loop
  Close #TmpFileNbr
End Sub

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

' Purpose: Run every wired-in Test* sub across all modules and log PASS/FAIL results to TestResults.txt.
Public Sub RunAllTests()
  Dim TmpPathStr As String
  Dim TmpPassCnt As Long
  Dim TmpFailCnt As Long
  Dim TmpSummaryStr As String

  TmpPathStr = ResolveTestLogPathStr()
  If LenB(TmpPathStr) > 0 Then
    If Dir$(TmpPathStr) <> vbNullString Then Kill TmpPathStr
  End If

  TestFilePathStr = TmpPathStr
  On Error GoTo Finally

  TestLogLine "Most tests log one line per case in a 4-part format:"
  TestLogLine "  Input | Answer | Expected | PASS or FAIL"
  TestLogLine "(some cases log multiple Input fields when the sub under test takes"
  TestLogLine "more than one argument, but Answer/Expected/PASS-FAIL are always the"
  TestLogLine "last 3.)"
  TestLogLine vbNullString

  ' modSmallFunctions.bas
  RunOneTest "TestShortenFormula"
  RunOneTest "TestStrCnt"
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
  RunOneTest "TestAddOperatorWhiteSpaceDoubleNeg"
  RunOneTest "TestAddOperatorWhiteSpaceScientificNotation"
  RunOneTest "TestAddOperatorWhiteSpaceErrorLiterals"
  RunOneTest "TestAddOperatorWhiteSpacePercent"
  RunOneTest "TestAddOperatorWhiteSpaceExternalRefs"
  RunOneTest "TestAddOperatorWhiteSpaceImplicitIntersection"
  RunOneTest "TestAddOperatorWhiteSpaceSpilledRange"

  ' modPrettyPrint.bas
  RunOneTest "TestRemoveQuotedSections"
  RunOneTest "TestPosSkipQuotedSections"
  RunOneTest "TestReplaceSkipQuotedSections"
  RunOneTest "TestPrettyPrintWouldChangeFormula"
  RunOneTest "TestPrettyPrintExact"
  RunOneTest "TestAnswersEscapeStr"
  RunOneTest "TestAnswersReadEntries"
  RunOneTest "TestPrettyPrintMatchAnswers"
  ' Test functions explicitly skipped.
  ' TestPrettyPrint -- excluded: no PASS/FAIL, just Input/Answer for manual
  ' review; run it standalone from the Immediate window instead.
  ' TestPrettyPrintWriteAnswers -- excluded: deliberately mutates the answers
  ' file, never run automatically.
  ' TestCellSetFormula2Safe -- excluded: mutates ActiveSheet.Range("A1").

  If LenB(TmpPathStr) > 0 Then
    CountPassFail TmpPathStr, TmpPassCnt, TmpFailCnt
    TmpSummaryStr = "PASS: " & TmpPassCnt & " | FAIL: " & TmpFailCnt

    TestLogLine "===== SUMMARY " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ====="
    TestLogLine TmpSummaryStr
  End If

Finally:
  TestFilePathStr = vbNullString 'Standalone Test* runs afterward fall back to Debug.Print.

  If LenB(TmpPathStr) > 0 Then
    MsgBox "Results written to:" & vbCrLf & TmpPathStr & vbCrLf & vbCrLf & TmpSummaryStr
  End If
End Sub
