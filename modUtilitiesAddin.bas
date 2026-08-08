Attribute VB_Name = "modUtilitiesAddin"
Option Explicit

Public Const kAddinFileNameStr = "Northbound.xlam"

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






