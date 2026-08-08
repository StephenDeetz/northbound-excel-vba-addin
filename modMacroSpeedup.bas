Attribute VB_Name = "modMacroSpeedup"

' ==========================================================================================
' Module: modMacroSpeedup
'
' Purpose:
'   Provides a one-line API for engaging and restoring Excel performance settings.
'   This module manages and restores the following Application properties:
'       - ScreenUpdating
'       - Calculation
'       - EnableEvents
'       - DisplayAlerts
'       - Cursor
'
'   Typical pattern:
'       MacroSpeedup                     ' Init + PushAll (default is msInit)
'           ... your macro code here ...
'       MacroSpeedup msClear             ' Restores original settings
'
'   Fine-grained control:
'       MacroSpeedup msPush, msDisplayAlerts
'           ... perform silent deletes ...
'       MacroSpeedup msPop,  msDisplayAlerts
'
'   Example Calls:
'       MacroSpeedup
'       MacroSpeedup msPush, msEnableEvents    'Saves Current Enable Events
'
'           ' Something that needs Application.EnableEvents on or off...
'           ' (events suppressed or temporarily allowed)
'
'       MacroSpeedup msPop, msEnableEvents     'Restores Last Enable Events Pushed
'       MacroSpeedup msClear
'
' Details:
'   This module manages a singleton CMacroSpeedup instance that internally keeps
'   a stack for each tracked Application property. Each Push saves the current value
'   and sets a "fast" value. Each Pop restores the most recent saved value.
'
'   Multiple nested routines can push and pop safely without overwriting each other's
'   settings. Each property's stack operates independently. msClear unwinds all stacks
'   until empty, ensuring Excel always returns to its starting configuration even if
'   intermediate routines forget to restore their changes.
'
'   The MacroSpeedup API extends naturally to nested and partial operations:
'       MacroSpeedup msPush, msEnableEvents
'           ' turn off events temporarily
'       MacroSpeedup msPop, msEnableEvents
'
'   No object variables, no setup, and no cleanup required -- one call in, one call out.
' ==========================================================================================

Option Explicit

' ===== Enums =====

Public Enum msStackOptions
  msInit
  msPush
  msPop
  msClear
End Enum

Public Enum msMacroItems
  msAllItems
  msScreenUpdating
  msCalculation
  msEnableEvents
  msDisplayAlerts
  msCursor
End Enum

' ===== Singleton access =====
Private MacroSpeedupObj As CMacroSpeedup

Private Function MacroSpeedupState() As CMacroSpeedup
  If MacroSpeedupObj Is Nothing Then
    Set MacroSpeedupObj = New CMacroSpeedup
  End If
  Set MacroSpeedupState = MacroSpeedupObj
End Function

' ===== Front Door =====
' Purpose: Push or pop saved Excel state (screen updating, calculation, events, alerts, cursor) around bulk operations.
Public Sub MacroSpeedup(Optional ByVal AStackOptions As msStackOptions = msInit, _
                        Optional ByVal AMacroItems As msMacroItems = msAllItems)
  
  Select Case AStackOptions
    Case msInit
      MacroSpeedupState.Init
      MacroSpeedupState.PushAll
    Case msPush
      If AMacroItems = msAllItems Then
        MacroSpeedupState.PushAll
      Else
        MacroSpeedupState.PushItem AMacroItems
      End If
    Case msPop
      If AMacroItems = msAllItems Then
        MacroSpeedupState.PopAll
      Else
        MacroSpeedupState.PopItem AMacroItems
      End If
    Case msClear
      MacroSpeedupState.ClearAll
    Case Else
      ' no-op
  End Select
End Sub

' Purpose: Print the live Application settings MacroSpeedup manages to the Immediate window.
Public Sub MSState()
  Dim TmpCalcStr As String
  Dim TmpCursorStr As String
  
  ' Current settings snapshot
  Select Case Application.Calculation
    Case xlCalculationAutomatic:     TmpCalcStr = "Automatic"
    Case xlCalculationManual:        TmpCalcStr = "Manual"
    Case xlCalculationSemiautomatic: TmpCalcStr = "Semiautomatic"
    Case Else:                       TmpCalcStr = CStr(Application.Calculation)
  End Select
  
  Select Case Application.Cursor
    Case xlDefault:        TmpCursorStr = "Default"
    Case xlWait:           TmpCursorStr = "Wait"
    Case xlNorthwestArrow: TmpCursorStr = "NorthwestArrow"
    Case xlIBeam:          TmpCursorStr = "IBeam"
    Case Else:             TmpCursorStr = CStr(Application.Cursor)
  End Select
  
  Debug.Print String(60, "=")
  Debug.Print "MacroSpeedup State (live Application values)"
  Debug.Print String(60, "-")
  Debug.Print "ScreenUpdating.....: "; Application.ScreenUpdating
  Debug.Print "Calculation........: "; TmpCalcStr
  Debug.Print "EnableEvents.......: "; Application.EnableEvents
  Debug.Print "DisplayAlerts......: "; Application.DisplayAlerts
  Debug.Print "Cursor.............: "; TmpCursorStr
  Debug.Print String(60, "=")
End Sub



' ============================
' Test: MacroSpeedup Test Helpers
' ============================

' Pretty-print a single assertion row
Private Sub PPResult(ByVal ACase As String, _
                     ByVal AExpected As String, _
                     ByVal AActual As String)
  Dim TmpPassFailStr As String
  TmpPassFailStr = IIf(StrComp(AExpected, AActual, vbBinaryCompare) = 0, "PASS", "FAIL")
  TestLogLine ACase & " | " & AActual & " | " & AExpected & " | " & TmpPassFailStr
End Sub

' Helpers to stringify the properties we test
Private Function BStr(ByVal ABool As Boolean) As String
  BStr = IIf(ABool, "True", "False")
End Function

Private Function CalcStr(ByVal ACalc As XlCalculation) As String
  Select Case ACalc
    Case xlCalculationAutomatic: CalcStr = "Automatic"
    Case xlCalculationManual:    CalcStr = "Manual"
    Case xlCalculationSemiautomatic: CalcStr = "Semiautomatic"
    Case Else: CalcStr = CStr(ACalc)
  End Select
End Function

Private Function CursorStr(ByVal ACursor As XlMousePointer) As String
  Select Case ACursor
    Case xlDefault: CursorStr = "Default"
    Case xlWait:    CursorStr = "Wait"
    Case xlNorthwestArrow: CursorStr = "NWArrow"
    Case xlIBeam:   CursorStr = "IBeam"
    Case Else:      CursorStr = CStr(ACursor)
  End Select
End Function

Private Sub SnapshotState(ByRef OutScreenUpdating As Boolean, _
                          ByRef OutCalc As XlCalculation, _
                          ByRef OutEvents As Boolean, _
                          ByRef OutAlerts As Boolean, _
                          ByRef OutCursor As XlMousePointer)
  OutScreenUpdating = Application.ScreenUpdating
  OutCalc = Application.Calculation
  OutEvents = Application.EnableEvents
  OutAlerts = Application.DisplayAlerts
  OutCursor = Application.Cursor
End Sub

' ============================
' Test: MacroSpeedup end-to-end
' ============================
Private Sub TestMacroSpeedup()
  Dim TmpBaselineScreenUpdating As Boolean
  Dim TmpBaselineCalculation    As XlCalculation
  Dim TmpBaselineEnableEvents   As Boolean
  Dim TmpBaselineDisplayAlerts  As Boolean
  Dim TmpBaselineCursor         As XlMousePointer

  Dim TmpPrePushCalculation     As XlCalculation
  Dim TmpPrePushDisplayAlerts   As Boolean

  On Error GoTo OnError

  ' 1) Capture baseline
  SnapshotState TmpBaselineScreenUpdating, TmpBaselineCalculation, TmpBaselineEnableEvents, _
                TmpBaselineDisplayAlerts, TmpBaselineCursor

  ' 2) Init (push-all) ? expect fast values
  MacroSpeedup msInit

  PPResult "Init.ScreenUpdating", "False", BStr(Application.ScreenUpdating)
  PPResult "Init.Calculation", "Manual", CalcStr(Application.Calculation)
  PPResult "Init.EnableEvents", "False", BStr(Application.EnableEvents)
  PPResult "Init.DisplayAlerts", "False", BStr(Application.DisplayAlerts)
  PPResult "Init.Cursor", "Wait", CursorStr(Application.Cursor)

  ' 3) Nested push of a single item (EnableEvents), then temporary override, then pop ? should restore to False
  MacroSpeedup msPush, msEnableEvents
  ' temporarily override to True (simulating a routine that needed one event fire)
  Application.EnableEvents = True
  MacroSpeedup msPop, msEnableEvents
  PPResult "Nested.Events.RestoredToFast", "False", BStr(Application.EnableEvents)

  ' 4) Mixed-order push/pop of two items should restore to pre-push state for each
  ' Capture current states (fast mode still on from Init)
  TmpPrePushCalculation = Application.Calculation
  TmpPrePushDisplayAlerts = Application.DisplayAlerts

  MacroSpeedup msPush, msCalculation
  MacroSpeedup msPush, msDisplayAlerts
  ' Pop in reverse (calc first), ensure each restores to the exact prior of its own push
  MacroSpeedup msPop, msCalculation
  PPResult "MixedOrder.CalcRestored", CalcStr(TmpPrePushCalculation), CalcStr(Application.Calculation)
  MacroSpeedup msPop, msDisplayAlerts
  PPResult "MixedOrder.AlertsRestored", BStr(TmpPrePushDisplayAlerts), BStr(Application.DisplayAlerts)

  ' 5) Pop on empty (safety): should be no change
  ' First, clear all to restore to baseline
  MacroSpeedup msClear
  PPResult "Clear.ScreenUpdating", BStr(TmpBaselineScreenUpdating), BStr(Application.ScreenUpdating)
  PPResult "Clear.Calculation", CalcStr(TmpBaselineCalculation), CalcStr(Application.Calculation)
  PPResult "Clear.EnableEvents", BStr(TmpBaselineEnableEvents), BStr(Application.EnableEvents)
  PPResult "Clear.DisplayAlerts", BStr(TmpBaselineDisplayAlerts), BStr(Application.DisplayAlerts)
  PPResult "Clear.Cursor", CursorStr(TmpBaselineCursor), CursorStr(Application.Cursor)

  ' Try popping when stacks are empty
  MacroSpeedup msPop, msScreenUpdating
  MacroSpeedup msPop, msCalculation
  MacroSpeedup msPop, msEnableEvents
  MacroSpeedup msPop, msDisplayAlerts
  MacroSpeedup msPop, msCursor

  PPResult "PopEmpty.ScreenUpdatingNoChange", BStr(TmpBaselineScreenUpdating), BStr(Application.ScreenUpdating)
  PPResult "PopEmpty.CalculationNoChange", CalcStr(TmpBaselineCalculation), CalcStr(Application.Calculation)
  PPResult "PopEmpty.EnableEventsNoChange", BStr(TmpBaselineEnableEvents), BStr(Application.EnableEvents)
  PPResult "PopEmpty.DisplayAlertsNoChange", BStr(TmpBaselineDisplayAlerts), BStr(Application.DisplayAlerts)
  PPResult "PopEmpty.CursorNoChange", CursorStr(TmpBaselineCursor), CursorStr(Application.Cursor)

Finally:
  Exit Sub

OnError:
  MacroSpeedup msClear
  TestLogLine "CRASHED: TestMacroSpeedup - " & err.Description
  err.Clear
  Resume Finally
End Sub

' ============================
' Extra Tests: place below existing tests
' ============================

' Test A: Push/Pop all-items explicitly (msPush/msPop)
Private Sub Test_MacroSpeedup_AllItemsPushPop()
  Dim TmpBaselineScreenUpdating As Boolean
  Dim TmpBaselineCalculation    As XlCalculation
  Dim TmpBaselineEnableEvents   As Boolean
  Dim TmpBaselineDisplayAlerts  As Boolean
  Dim TmpBaselineCursor         As XlMousePointer

  On Error GoTo OnError

  SnapshotState TmpBaselineScreenUpdating, TmpBaselineCalculation, TmpBaselineEnableEvents, _
                TmpBaselineDisplayAlerts, TmpBaselineCursor

  ' Push ALL (default msAllItems when omitted)
  MacroSpeedup msPush
  PPResult "PushAll.ScreenUpdating", "False", BStr(Application.ScreenUpdating)
  PPResult "PushAll.Calculation", "Manual", CalcStr(Application.Calculation)
  PPResult "PushAll.EnableEvents", "False", BStr(Application.EnableEvents)
  PPResult "PushAll.DisplayAlerts", "False", BStr(Application.DisplayAlerts)
  PPResult "PushAll.Cursor", "Wait", CursorStr(Application.Cursor)

  ' Pop ALL
  MacroSpeedup msPop
  PPResult "PopAll.ScreenUpdatingRestored", BStr(TmpBaselineScreenUpdating), BStr(Application.ScreenUpdating)
  PPResult "PopAll.CalculationRestored", CalcStr(TmpBaselineCalculation), CalcStr(Application.Calculation)
  PPResult "PopAll.EnableEventsRestored", BStr(TmpBaselineEnableEvents), BStr(Application.EnableEvents)
  PPResult "PopAll.DisplayAlertsRestored", BStr(TmpBaselineDisplayAlerts), BStr(Application.DisplayAlerts)
  PPResult "PopAll.CursorRestored", CursorStr(TmpBaselineCursor), CursorStr(Application.Cursor)

Finally:
  Exit Sub

OnError:
  MacroSpeedup msClear
  TestLogLine "CRASHED: Test_MacroSpeedup_AllItemsPushPop - " & err.Description
  err.Clear
  Resume Finally
End Sub

' Test B: Deep nesting (multiple pushes each item)
Private Sub Test_MacroSpeedup_NestedDepth()
  Dim TmpBaselineScreenUpdating As Boolean
  Dim TmpBaselineCalculation    As XlCalculation
  Dim TmpBaselineEnableEvents   As Boolean
  Dim TmpBaselineDisplayAlerts  As Boolean
  Dim TmpBaselineCursor         As XlMousePointer

  On Error GoTo OnError

  SnapshotState TmpBaselineScreenUpdating, TmpBaselineCalculation, TmpBaselineEnableEvents, _
                TmpBaselineDisplayAlerts, TmpBaselineCursor

  MacroSpeedup msInit                      ' level 1 (all)
  MacroSpeedup msPush, msEnableEvents      ' level 2 (events)
  MacroSpeedup msPush, msEnableEvents      ' level 3 (events again)
  MacroSpeedup msPush, msCalculation       ' level 4 (calc)
  MacroSpeedup msPush, msDisplayAlerts     ' level 5 (alerts)

  ' Partial unwind in mixed order
  MacroSpeedup msPop, msEnableEvents
  MacroSpeedup msPop, msDisplayAlerts
  MacroSpeedup msPop, msCalculation
  MacroSpeedup msPop, msEnableEvents

  ' Final unwind
  MacroSpeedup msClear

  PPResult "NestedDepth.ScreenUpdatingRestored", BStr(TmpBaselineScreenUpdating), BStr(Application.ScreenUpdating)
  PPResult "NestedDepth.CalculationRestored", CalcStr(TmpBaselineCalculation), CalcStr(Application.Calculation)
  PPResult "NestedDepth.EnableEventsRestored", BStr(TmpBaselineEnableEvents), BStr(Application.EnableEvents)
  PPResult "NestedDepth.DisplayAlertsRestored", BStr(TmpBaselineDisplayAlerts), BStr(Application.DisplayAlerts)
  PPResult "NestedDepth.CursorRestored", CursorStr(TmpBaselineCursor), CursorStr(Application.Cursor)

Finally:
  Exit Sub

OnError:
  MacroSpeedup msClear
  TestLogLine "CRASHED: Test_MacroSpeedup_NestedDepth - " & err.Description
  err.Clear
  Resume Finally
End Sub

' Test C: Respect user's Manual baseline
Private Sub Test_MacroSpeedup_ManualBaseline()
  Dim TmpBaselineCalculation As XlCalculation

  On Error GoTo OnError

  ' Force a non-default baseline deliberately
  TmpBaselineCalculation = Application.Calculation
  Application.Calculation = xlCalculationManual

  MacroSpeedup msInit       ' stores Manual, sets Manual (idempotent)
  MacroSpeedup msClear      ' must restore to Manual (not Automatic)

  PPResult "ManualBaseline.Restored", "Manual", CalcStr(Application.Calculation)

Finally:
  ' restore original baseline
  Application.Calculation = TmpBaselineCalculation
  Exit Sub

OnError:
  MacroSpeedup msClear
  TestLogLine "CRASHED: Test_MacroSpeedup_ManualBaseline - " & err.Description
  err.Clear
  Resume Finally
End Sub

' Test D: Interleaved sequences + ClearAll as finalizer
Private Sub Test_MacroSpeedup_Interleaved()
  Dim TmpBaselineScreenUpdating As Boolean
  Dim TmpBaselineCalculation    As XlCalculation
  Dim TmpBaselineEnableEvents   As Boolean
  Dim TmpBaselineDisplayAlerts  As Boolean
  Dim TmpBaselineCursor         As XlMousePointer

  On Error GoTo OnError

  SnapshotState TmpBaselineScreenUpdating, TmpBaselineCalculation, TmpBaselineEnableEvents, _
                TmpBaselineDisplayAlerts, TmpBaselineCursor

  MacroSpeedup msPush, msScreenUpdating
  MacroSpeedup msPush, msDisplayAlerts
  MacroSpeedup msPush, msCalculation

  ' Intentionally pop only some
  MacroSpeedup msPop, msDisplayAlerts
  MacroSpeedup msPush, msEnableEvents
  MacroSpeedup msPush, msCursor

  ' Finalizer must unwind everything
  MacroSpeedup msClear

  PPResult "Interleaved.ScreenUpdatingRestored", BStr(TmpBaselineScreenUpdating), BStr(Application.ScreenUpdating)
  PPResult "Interleaved.CalculationRestored", CalcStr(TmpBaselineCalculation), CalcStr(Application.Calculation)
  PPResult "Interleaved.EnableEventsRestored", BStr(TmpBaselineEnableEvents), BStr(Application.EnableEvents)
  PPResult "Interleaved.DisplayAlertsRestored", BStr(TmpBaselineDisplayAlerts), BStr(Application.DisplayAlerts)
  PPResult "Interleaved.CursorRestored", CursorStr(TmpBaselineCursor), CursorStr(Application.Cursor)

Finally:
  Exit Sub

OnError:
  MacroSpeedup msClear
  TestLogLine "CRASHED: Test_MacroSpeedup_Interleaved - " & err.Description
  err.Clear
  Resume Finally
End Sub


' ============================
' Extra Micro-Tests
' ============================

' Test E: Idempotent pushes on a single item (3x) with stepwise pops
Private Sub Test_MacroSpeedup_IdempotentPush()
  Dim TmpBaselineScreenUpdating As Boolean
  Dim TmpBaselineCalculation    As XlCalculation
  Dim TmpBaselineEnableEvents   As Boolean
  Dim TmpBaselineDisplayAlerts  As Boolean
  Dim TmpBaselineCursor         As XlMousePointer

  On Error GoTo OnError

  SnapshotState TmpBaselineScreenUpdating, TmpBaselineCalculation, TmpBaselineEnableEvents, _
                TmpBaselineDisplayAlerts, TmpBaselineCursor

  ' Use DisplayAlerts for clarity
  MacroSpeedup msPush, msDisplayAlerts    ' push 1
  MacroSpeedup msPush, msDisplayAlerts    ' push 2
  MacroSpeedup msPush, msDisplayAlerts    ' push 3

  ' While stacked, value is forced False
  PPResult "Idem.Push3.AlertsFast", "False", BStr(Application.DisplayAlerts)

  ' Pop once -> still False (prior layer had False)
  MacroSpeedup msPop, msDisplayAlerts
  PPResult "Idem.Pop1.AlertsStillFast", "False", BStr(Application.DisplayAlerts)

  ' Pop twice -> still False (prior layer had False)
  MacroSpeedup msPop, msDisplayAlerts
  PPResult "Idem.Pop2.AlertsStillFast", "False", BStr(Application.DisplayAlerts)

  ' Pop third -> baseline restored
  MacroSpeedup msPop, msDisplayAlerts
  PPResult "Idem.Pop3.AlertsBaseline", BStr(TmpBaselineDisplayAlerts), BStr(Application.DisplayAlerts)

  ' Safety: Clear should now be a no-op on Alerts
  MacroSpeedup msClear
  PPResult "Idem.Clear.Baseline.Alerts", BStr(TmpBaselineDisplayAlerts), BStr(Application.DisplayAlerts)
  PPResult "Idem.Clear.Baseline.ScreenUpdating", BStr(TmpBaselineScreenUpdating), BStr(Application.ScreenUpdating)
  PPResult "Idem.Clear.Baseline.Calculation", CalcStr(TmpBaselineCalculation), CalcStr(Application.Calculation)
  PPResult "Idem.Clear.Baseline.EnableEvents", BStr(TmpBaselineEnableEvents), BStr(Application.EnableEvents)
  PPResult "Idem.Clear.Baseline.Cursor", CursorStr(TmpBaselineCursor), CursorStr(Application.Cursor)

Finally:
  Exit Sub

OnError:
  MacroSpeedup msClear
  TestLogLine "CRASHED: Test_MacroSpeedup_IdempotentPush - " & err.Description
  err.Clear
  Resume Finally
End Sub

' Test F: Mid-scope override -- Events pushed False, temporarily set True, pop -> expect False, then Clear -> baseline
Private Sub Test_MacroSpeedup_MidScopeOverride()
  Dim TmpBaselineScreenUpdating As Boolean
  Dim TmpBaselineCalculation    As XlCalculation
  Dim TmpBaselineEnableEvents   As Boolean
  Dim TmpBaselineDisplayAlerts  As Boolean
  Dim TmpBaselineCursor         As XlMousePointer

  On Error GoTo OnError

  SnapshotState TmpBaselineScreenUpdating, TmpBaselineCalculation, TmpBaselineEnableEvents, _
                TmpBaselineDisplayAlerts, TmpBaselineCursor

  MacroSpeedup msInit                    ' all fast (Events=False)
  MacroSpeedup msPush, msEnableEvents    ' push another layer for Events

  ' Temporary override inside routine
  Application.EnableEvents = True
  PPResult "Override.Events.TempTrue", "True", BStr(Application.EnableEvents)

  ' Pop should restore to fast (False), not baseline
  MacroSpeedup msPop, msEnableEvents
  PPResult "Override.Events.RestoredToFast", "False", BStr(Application.EnableEvents)

  ' Finalize
  MacroSpeedup msClear
  PPResult "Override.Clear.ScreenUpdating", BStr(TmpBaselineScreenUpdating), BStr(Application.ScreenUpdating)
  PPResult "Override.Clear.Calculation", CalcStr(TmpBaselineCalculation), CalcStr(Application.Calculation)
  PPResult "Override.Clear.EnableEvents", BStr(TmpBaselineEnableEvents), BStr(Application.EnableEvents)
  PPResult "Override.Clear.DisplayAlerts", BStr(TmpBaselineDisplayAlerts), BStr(Application.DisplayAlerts)
  PPResult "Override.Clear.Cursor", CursorStr(TmpBaselineCursor), CursorStr(Application.Cursor)

Finally:
  Exit Sub

OnError:
  MacroSpeedup msClear
  TestLogLine "CRASHED: Test_MacroSpeedup_MidScopeOverride - " & err.Description
  err.Clear
  Resume Finally
End Sub

' Test G: Error-path finalizer simulate error after pushes, ensure Clear restores baseline in handler
Private Sub Test_MacroSpeedup_ErrorPathFinalizer()
  Dim TmpBaselineScreenUpdating As Boolean
  Dim TmpBaselineCalculation    As XlCalculation
  Dim TmpBaselineEnableEvents   As Boolean
  Dim TmpBaselineDisplayAlerts  As Boolean
  Dim TmpBaselineCursor         As XlMousePointer

  SnapshotState TmpBaselineScreenUpdating, TmpBaselineCalculation, TmpBaselineEnableEvents, _
                TmpBaselineDisplayAlerts, TmpBaselineCursor

  On Error GoTo OnError

  ' Push a mix of items, leave some unpopped
  MacroSpeedup msPush, msScreenUpdating
  MacroSpeedup msPush, msCalculation
  MacroSpeedup msPush, msDisplayAlerts
  MacroSpeedup msPush, msCursor

  ' Simulate an error mid-routine
  err.Raise 12345, "Test_MacroSpeedup_ErrorPathFinalizer", "Simulated error for test"

  ' (unreached)
  TestLogLine "Test_MacroSpeedup_ErrorPathFinalizer | did not raise as expected | - | FAIL"
  GoTo Finally

OnError:
  ' Finalizer path ensure everything is unwound
  MacroSpeedup msClear

  ' Asserts: baseline restored
  PPResult "ErrorPath.Restore.ScreenUpdating", BStr(TmpBaselineScreenUpdating), BStr(Application.ScreenUpdating)
  PPResult "ErrorPath.Restore.Calculation", CalcStr(TmpBaselineCalculation), CalcStr(Application.Calculation)
  PPResult "ErrorPath.Restore.EnableEvents", BStr(TmpBaselineEnableEvents), BStr(Application.EnableEvents)
  PPResult "ErrorPath.Restore.DisplayAlerts", BStr(TmpBaselineDisplayAlerts), BStr(Application.DisplayAlerts)
  PPResult "ErrorPath.Restore.Cursor", CursorStr(TmpBaselineCursor), CursorStr(Application.Cursor)

Finally:
  err.Clear
End Sub





