Attribute VB_Name = "modCodeExport"
Option Explicit

Private Const kModulePathFile As String = "NBPub_AddinModulePath.txt"

' Purpose: Export every VBA module/class/form in a workbook's VBA project to a folder.
Public Function WbkExportVBAModules(ByVal AWbk As Workbook, _
                                    ByVal ADirStr As String, _
                                    Optional ByRef AErrStr As Variant) As Long
  'Dim TmpVBProj   As Object
  Dim TmpVBComp   As Object
  Dim TmpFNameStr As String
  Dim TmpFPathStr As String
  Dim TmpCnt      As Long
  
  WbkExportVBAModules = 0
  
  If AWbk Is Nothing Then
    If IsMissing(AErrStr) = False Then
      AErrStr = "Workbook is Null."
    End If
    
    Exit Function
  End If
  
  If AWbk.Path = vbNullString Then
    If IsMissing(AErrStr) = False Then
      AErrStr = "Save Workbook Before Code Export."
    End If
    
    Exit Function
  End If
  
  On Error GoTo OnError
  
  TmpFPathStr = ADirStr
  If Right$(TmpFPathStr, 1) <> "\" Then TmpFPathStr = TmpFPathStr & "\"
  
  'Set TmpVBProj = AWbk.VBProject
  
  TmpCnt = 0
  For Each TmpVBComp In AWbk.VBProject.VBComponents
    Dim TmpShouldExport As Boolean
    TmpShouldExport = False

    Select Case TmpVBComp.Type
      Case 1 ' StdModule: skip if no procedures
        TmpFNameStr = TmpFPathStr & TmpVBComp.Name & ".bas"
        TmpShouldExport = TmpVBComp.CodeModule.CountOfLines > TmpVBComp.CodeModule.CountOfDeclarationLines
      Case 2 ' ClassModule: export if any lines (declarations are meaningful in classes)
        TmpFNameStr = TmpFPathStr & TmpVBComp.Name & ".cls"
        TmpShouldExport = TmpVBComp.CodeModule.CountOfLines > 0
      Case 100 ' Document (ThisWorkbook, Sheet modules): skip if no procedures
        TmpFNameStr = TmpFPathStr & TmpVBComp.Name & ".cls"
        TmpShouldExport = TmpVBComp.CodeModule.CountOfLines > TmpVBComp.CodeModule.CountOfDeclarationLines
      Case 3 ' MSForm
        TmpFNameStr = TmpFPathStr & TmpVBComp.Name & ".frm"
        TmpShouldExport = True
      Case Else
        TmpFNameStr = vbNullString
    End Select

    If TmpFNameStr <> vbNullString Then
      If TmpShouldExport Then
        TmpVBComp.Export TmpFNameStr
        TmpCnt = TmpCnt + 1
      End If
    End If
  Next TmpVBComp
  
  WbkExportVBAModules = TmpCnt
  Exit Function

OnError:
  If IsMissing(AErrStr) = False Then
    AErrStr = "Error " & err.Number & ": " & err.Description
  End If
  err.Clear
End Function

Private Sub TestWbkExportVBAModules()
  Dim TmpResult As Boolean
  Dim TmpErrStr As String

  If IsStandaloneTestRun() Then ClearImmediateWindow

  '@Ignore UnassignedVariableUsage
  TmpResult = WbkExportVBAModules(ThisWorkbook, "C:\Temp\", TmpErrStr)
  
  If TmpResult Then
    Debug.Print "Export successful"
  Else
    '@Ignore UnassignedVariableUsage
    Debug.Print "Export failed: " & TmpErrStr
  End If
End Sub

' Purpose: Prompt for a folder (remembering the last one used) and export ThisWorkbook's VBA modules to it.
Public Sub AddinExportCodeModules()
  Dim TmpErrStr As String
  Dim TmpResult As Long
  Dim TmpDirStr As String
  Dim TmpLastPathStr As String

  TmpLastPathStr = AppDataFileReadStr(kModulePathFile)
  If LenB(TmpLastPathStr) > 0 Then
    TmpDirStr = GetFolder(TmpLastPathStr)
  Else
    TmpDirStr = GetFolder(ThisWorkbook.Path)
  End If
  If TmpDirStr = "" Then Exit Sub
  AppDataFileWriteStr TmpDirStr, kModulePathFile
  
  TmpResult = WbkExportVBAModules(ThisWorkbook, TmpDirStr, TmpErrStr)
  
  If TmpResult Then
    If MsgBox(CStr(TmpResult) & " module(s) exported." & vbCrLf & vbCrLf & "Open in Explorer?", _
          vbQuestion + vbYesNo) = vbYes Then
      OpenPathInExplorer TmpDirStr
    End If
  Else
    If TmpErrStr <> vbNullString Then
      MsgBox "Export Failed: " & TmpErrStr, vbExclamation
    Else
      MsgBox "Export Failed: No VBA Modules or Forms Found.", vbExclamation
    End If
  End If
  
End Sub








