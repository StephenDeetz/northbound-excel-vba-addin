Attribute VB_Name = "modUtilitiesAddin"
' Copyright (c) 2026 Northbound Group
' Contact: stephendeetz@northboundgroup.com
' SPDX-License-Identifier: MIT
Option Explicit

Public Const kAddinFileNameStr = "Northbound.xlam"
Private Const kAppDataSubDir  As String = "Northbound"

' Purpose: Return (creating if needed) the add-in's settings directory under AppData.
Public Function GetUtilitiesAddInSettingsDir() As String
  Dim TmpDir As String
  Const kUtilitiesAddInSubDir = "Northbound Add-in"

  GetUtilitiesAddInSettingsDir = vbNullString

  TmpDir = GetAppDataDir

  If DirExists(TmpDir) = False Then
    Exit Function
  End If

  TmpDir = IncludeTrailingBackslash(TmpDir) + kUtilitiesAddInSubDir

  If DirExists(TmpDir) = False Then
    MkDir TmpDir

    If DirExists(TmpDir) = False Then
      Exit Function
    End If
  End If

  GetUtilitiesAddInSettingsDir = TmpDir

End Function

Private Sub TestGetUtilitiesAddInSettingsDir()
  If IsStandaloneTestRun() Then ClearImmediateWindow
  Debug.Print GetUtilitiesAddInSettingsDir
End Sub

' Purpose: Run pre-flight guard checks (read-only, protection, sheet type, selection shape) before a bulk operation.
Public Function AllClear2(ByVal AIsWbkReadOnlyWarning As Boolean, _
                          ByVal AWbkProtectedStructWarning As Boolean, _
                          ByVal AIsActiveShtWorksheetWarning As Boolean, _
                          ByVal AProtectedActiveShtWarning As Boolean, _
                          ByVal AAnyProtectedShtWarning As Boolean, _
                          ByVal AIsSelRngWarning As Boolean, _
                          ByVal AIsSelSingleAreaWarning As Boolean) As Boolean
  AllClear2 = False
  
  If ActiveWorkbook Is Nothing Then
    MsgBox "No active workbook. Please open or activate a workbook."
    Exit Function
  End If

  If AIsWbkReadOnlyWarning Then
    If ActiveWorkbook.ReadOnly Then
      MsgBox "Workbook is read-only."
      Exit Function
    End If
  End If

  If AWbkProtectedStructWarning Then
    If ActiveWorkbook.ProtectStructure Then
      MsgBox "Please Unprotect Workbook Structure Before Using."
      Exit Function
    End If
  End If

  If AIsActiveShtWorksheetWarning Then
    If TypeName(ActiveSheet) <> "Worksheet" Then
      MsgBox "Active sheet must be a worksheet."
      Exit Function
    End If
  End If

  If AProtectedActiveShtWarning Then
    If ActiveSheet.ProtectContents Then
      MsgBox "Please Unprotect Active Sheet Before Using."
      Exit Function
    End If
  End If

  If AAnyProtectedShtWarning Then
    If AnySheetsProtected(ActiveWorkbook) Then
      MsgBox "Please Unprotect Sheets Before Using."
      Exit Function
    End If
  End If

  If AIsSelRngWarning Then
    If Not TypeOf Selection Is Range Then
      MsgBox "Selection must be a range."
      Exit Function
    End If
  End If

  If AIsSelSingleAreaWarning Then
    If Not TypeOf Selection Is Range Then
      MsgBox "Selection must be a range."
      Exit Function
    End If
    If Selection.Areas.Count > 1 Then
      MsgBox "Please select a single area."
      Exit Function
    End If
  End If

  AllClear2 = True
End Function


Private Function AppDataDirStr() As String
  Dim TmpDir As String

  AppDataDirStr = vbNullString

  TmpDir = Environ$("APPDATA")
  If LenB(TmpDir) = 0 Then Exit Function
  If Right$(TmpDir, 1) <> "\" Then TmpDir = TmpDir & "\"
  TmpDir = TmpDir & kAppDataSubDir

  If Dir(TmpDir, vbDirectory) = vbNullString Then
    On Error Resume Next
    MkDir TmpDir
    On Error GoTo 0
    If Dir(TmpDir, vbDirectory) = vbNullString Then Exit Function
  End If

  AppDataDirStr = TmpDir & "\"
End Function

Private Sub FileWriteStr(ByVal APathStr As String, ByVal AStr As String)
  Dim TmpFileNbr As Integer
  TmpFileNbr = FreeFile
  Open APathStr For Output As #TmpFileNbr
  Print #TmpFileNbr, AStr
  Close #TmpFileNbr
End Sub

Private Function FileReadStr(ByVal APathStr As String) As String
  Dim TmpFileNbr As Integer
  Dim TmpLine As String
  Dim TmpResult As String

  FileReadStr = vbNullString
  If Dir(APathStr) = vbNullString Then Exit Function

  TmpFileNbr = FreeFile
  Open APathStr For Input As #TmpFileNbr
  Do Until EOF(TmpFileNbr)
    Line Input #TmpFileNbr, TmpLine
    TmpResult = TmpResult & TmpLine & vbCrLf
  Loop
  Close #TmpFileNbr

  If Right$(TmpResult, 2) = vbCrLf Then
    TmpResult = Left$(TmpResult, Len(TmpResult) - 2)
  End If
  FileReadStr = TmpResult
End Function

' Purpose: Write a string to a file in the add-in's AppData settings directory.
Public Sub AppDataFileWriteStr(ByVal AStr As String, ByVal AFileNameStr As String)
  Dim TmpDir As String
  TmpDir = AppDataDirStr
  If LenB(TmpDir) = 0 Then Exit Sub
  FileWriteStr TmpDir & AFileNameStr, AStr
End Sub

' Purpose: Read a string from a file in the add-in's AppData settings directory.
Public Function AppDataFileReadStr(ByVal AFileNameStr As String) As String
  Dim TmpDir As String
  TmpDir = AppDataDirStr
  If LenB(TmpDir) = 0 Then Exit Function
  AppDataFileReadStr = FileReadStr(TmpDir & AFileNameStr)
End Function

Private Sub TestAppDataFileWriteStr()
  If IsStandaloneTestRun() Then ClearImmediateWindow
  AppDataFileWriteStr "TestValue", "TestAppData.txt"
  Debug.Print "TestAppDataFileWriteStr: wrote 'TestValue' to TestAppData.txt"
End Sub

Private Sub TestAppDataFileReadStr()
  If IsStandaloneTestRun() Then ClearImmediateWindow
  Debug.Print "TestAppDataFileReadStr: " & AppDataFileReadStr("TestAppData.txt")
End Sub








