Attribute VB_Name = "modCodeImport"
Option Explicit

' Self-contained: no compile-time references to other imported modules.
'
' Design notes:
'
'   - Transfer sequence: blank source (Name1) FIRST to drop its public symbols,
'     then blank+rewrite destination, then Remove the empty temp. This prevents
'     a duplicate-symbol compile error while both modules coexist in the project.
'
'   - VBComponents.Remove silently fails when other modules hold compile-time refs
'     to the target's public symbols -- in-place transfer avoids that entirely.
'
'   - Code pane for the existing module is closed before Import runs so VBE does
'     not display a stale window during or after the transfer.

Private Const kAppDataSubDir  As String = "Northbound"
Private Const kModulePathFile As String = "NBPub_AddinModulePath.txt"

' -- AppData helpers ----------------------------------------------------------

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

Public Sub AppDataFileWriteStr(ByVal AStr As String, ByVal AFileNameStr As String)
  Dim TmpDir As String
  TmpDir = AppDataDirStr
  If LenB(TmpDir) = 0 Then Exit Sub
  FileWriteStr TmpDir & AFileNameStr, AStr
End Sub

Public Function AppDataFileReadStr(ByVal AFileNameStr As String) As String
  Dim TmpDir As String
  TmpDir = AppDataDirStr
  If LenB(TmpDir) = 0 Then Exit Function
  AppDataFileReadStr = FileReadStr(TmpDir & AFileNameStr)
End Function

Private Sub TestAppDataFileWriteStr()
  AppDataFileWriteStr "TestValue", "TestAppData.txt"
  Debug.Print "TestAppDataFileWriteStr: wrote 'TestValue' to TestAppData.txt"
End Sub

Private Sub TestAppDataFileReadStr()
  Debug.Print "TestAppDataFileReadStr: " & AppDataFileReadStr("TestAppData.txt")
End Sub

' -- VBE helpers --------------------------------------------------------------

Private Function FileModuleNameStr(ByVal APathStr As String) As String
  Dim TmpName As String
  TmpName = Mid$(APathStr, InStrRev(APathStr, "\") + 1)
  FileModuleNameStr = Left$(TmpName, InStrRev(TmpName, ".") - 1)
End Function

Private Function FindVBComp(ByVal AWbk As Workbook, ByVal ANameStr As String) As Object
  Dim TmpIdx As Long
  Set FindVBComp = Nothing
  For TmpIdx = 1 To AWbk.VBProject.VBComponents.Count
    If AWbk.VBProject.VBComponents(TmpIdx).Name = ANameStr Then
      Set FindVBComp = AWbk.VBProject.VBComponents(TmpIdx)
      Exit Function
    End If
  Next TmpIdx
End Function

Private Sub CloseModuleCodePane(ByVal AWbk As Workbook, ByVal AComp As Object)
  Dim TmpCnt As Long
  On Error Resume Next
  For TmpCnt = AWbk.VBProject.VBE.CodePanes.Count To 1 Step -1
    If AWbk.VBProject.VBE.CodePanes(TmpCnt).CodeModule Is AComp.CodeModule Then
      AWbk.VBProject.VBE.CodePanes(TmpCnt).Window.Close
    End If
  Next TmpCnt
  On Error GoTo 0
End Sub

Private Sub TransferCodeModule(ByVal ASrc As Object, ByVal ADst As Object)
  ' Safe sequence: blank the source FIRST so its public symbols are gone
  ' before we write them into the destination. This prevents duplicate-symbol
  ' compile errors during the transfer and makes Remove more likely to succeed.
  Dim TmpCode As String

  If ASrc.CountOfLines > 0 Then
    TmpCode = ASrc.Lines(1, ASrc.CountOfLines)
    ASrc.DeleteLines 1, ASrc.CountOfLines
  End If

  If ADst.CountOfLines > 0 Then ADst.DeleteLines 1, ADst.CountOfLines

  If LenB(TmpCode) > 0 Then ADst.InsertLines 1, TmpCode
End Sub

' -- Import -------------------------------------------------------------------

Private Sub ImportOneModule(ByVal AWbk As Workbook, _
                            ByVal AFilePathStr As String, _
                            ByRef AResultStr As String)
  Dim TmpNameStr As String
  Dim TmpExisting As Object
  Dim TmpImported As Object

  On Error GoTo OnError

  TmpNameStr = FileModuleNameStr(AFilePathStr)

  If TmpNameStr = "modCodeImport" Then
    AResultStr = AResultStr & "SKIP: modCodeImport" & vbLf
    Exit Sub
  End If

  Set TmpExisting = FindVBComp(AWbk, TmpNameStr)

  If Not TmpExisting Is Nothing Then
    CloseModuleCodePane AWbk, TmpExisting
    ' Blank the destination BEFORE importing. Public Enum members are
    ' project-global (unlike Sub/Function names, they can't be qualified by
    ' module), so if the incoming file declares one, importing it while the
    ' existing module still has its copy creates an instant "Ambiguous name
    ' detected" COMPILE error -- which On Error cannot catch, since the code
    ' never gets to run. Clearing the destination first means only one copy
    ' of any such symbol ever exists at a time.
    If TmpExisting.CodeModule.CountOfLines > 0 Then
      TmpExisting.CodeModule.DeleteLines 1, TmpExisting.CodeModule.CountOfLines
    End If
  End If

  Set TmpImported = AWbk.VBProject.VBComponents.Import(AFilePathStr)

  If TmpExisting Is Nothing Then
    AResultStr = AResultStr & "OK:   " & TmpImported.Name & vbLf
    Exit Sub
  End If

  ' Existing module: VBE renamed the import to TmpNameStr & "1", and the
  ' destination is already blank (above). Transfer the code across, then
  ' remove the now-empty temp module.
  TransferCodeModule TmpImported.CodeModule, TmpExisting.CodeModule

  On Error Resume Next
  AWbk.VBProject.VBComponents.Remove TmpImported
  On Error GoTo OnError

  AResultStr = AResultStr & "OK:   " & TmpNameStr & " (updated)" & vbLf
  Exit Sub

OnError:
  AResultStr = AResultStr & "ERR:  " & TmpNameStr & " - " & err.Description & vbLf
  err.Clear
End Sub

Public Sub AddinUpdateCodeModules()
  Dim TmpResultStr As String
  Dim TmpLastPathStr As String
  Dim TmpItem As Variant

  On Error GoTo Finally

  Application.EnableEvents = False

  TmpLastPathStr = AppDataFileReadStr(kModulePathFile)

  With Application.FileDialog(msoFileDialogFilePicker)
    .Title = "Select Code Modules to Import"
    .InitialFileName = TmpLastPathStr
    If LenB(.InitialFileName) = 0 Then .InitialFileName = ThisWorkbook.Path & "\"
    .Filters.Clear
    .Filters.Add "VBA Code Files", "*.bas;*.cls;*.frm"
    .AllowMultiSelect = True
    If .Show = False Then GoTo Finally
    AppDataFileWriteStr Left$(.SelectedItems(1), InStrRev(.SelectedItems(1), "\")), _
                        kModulePathFile
    For Each TmpItem In .SelectedItems
      ImportOneModule ThisWorkbook, CStr(TmpItem), TmpResultStr
    Next TmpItem
  End With

Finally:
  Application.EnableEvents = True
  If LenB(TmpResultStr) > 0 Then
    MsgBox TmpResultStr, vbInformation, "Import Code Modules"
  End If
End Sub
