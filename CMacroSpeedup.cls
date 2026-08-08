VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMacroSpeedup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Copyright (c) 2026 Northbound Group
' Contact: stephendeetz@northboundgroup.com
' SPDX-License-Identifier: MIT
' Class Module: CMacroSpeedup
Option Explicit

' Stacks (maps to provided CStack class)
Private StkScreenUpdating As CStack    ' Boolean
Private StkCalculation    As CStack    ' Long (XlCalculation)
Private StkEnableEvents   As CStack    ' Boolean
Private StkDisplayAlerts  As CStack    ' Boolean
Private StkCursor         As CStack    ' Long (XlMousePointer)

Private Sub Class_Initialize()
  Set StkScreenUpdating = New CStack
  Set StkCalculation = New CStack
  Set StkEnableEvents = New CStack
  Set StkDisplayAlerts = New CStack
  Set StkCursor = New CStack
End Sub

' ===== Public Surface =====
' Purpose: Clear all item stacks, discarding any pending push/pop state.
Public Sub Init()
  StkScreenUpdating.Clear
  StkCalculation.Clear
  StkEnableEvents.Clear
  StkDisplayAlerts.Clear
  StkCursor.Clear
End Sub

' Purpose: Push and apply the speedup setting for every tracked item.
Public Sub PushAll()
  PushScreenUpdating
  PushCalculation
  PushEnableEvents
  PushDisplayAlerts
  PushCursorWait
End Sub

' Purpose: Pop and restore the saved setting for every tracked item.
Public Sub PopAll()
  PopCursor
  PopDisplayAlerts
  PopEnableEvents
  PopCalculation
  PopScreenUpdating
End Sub

' Purpose: Pop every pending entry off all item stacks, restoring original settings.
Public Sub ClearAll()
  Do While StkCursor.Count > 0:         PopCursor:         Loop
  Do While StkDisplayAlerts.Count > 0:  PopDisplayAlerts:  Loop
  Do While StkEnableEvents.Count > 0:   PopEnableEvents:   Loop
  Do While StkCalculation.Count > 0:    PopCalculation:    Loop
  Do While StkScreenUpdating.Count > 0: PopScreenUpdating: Loop
End Sub

' Purpose: Push and apply the speedup setting for a single tracked item.
Public Sub PushItem(ByVal AItem As msMacroItems)
  Select Case AItem
    Case msScreenUpdating: PushScreenUpdating
    Case msCalculation:    PushCalculation
    Case msEnableEvents:   PushEnableEvents
    Case msDisplayAlerts:  PushDisplayAlerts
    Case msCursor:         PushCursorWait
  End Select
End Sub

' Purpose: Pop and restore the saved setting for a single tracked item.
Public Sub PopItem(ByVal AItem As msMacroItems)
  Select Case AItem
    Case msScreenUpdating: PopScreenUpdating
    Case msCalculation:    PopCalculation
    Case msEnableEvents:   PopEnableEvents
    Case msDisplayAlerts:  PopDisplayAlerts
    Case msCursor:         PopCursor
  End Select
End Sub

' ===== Push/Pop per item (idempotent sets) =====
Private Sub PushScreenUpdating()
  Dim TmpPrev As Boolean
  TmpPrev = Application.ScreenUpdating
  StkScreenUpdating.Push TmpPrev
  If TmpPrev <> False Then Application.ScreenUpdating = False
End Sub

Private Sub PopScreenUpdating()
  Dim TmpPrev As Variant
  TmpPrev = StkScreenUpdating.Pop
  If Not IsNull(TmpPrev) Then
    If Application.ScreenUpdating <> CBool(TmpPrev) Then Application.ScreenUpdating = CBool(TmpPrev)
  End If
End Sub

Private Sub PushCalculation()
  Dim TmpPrev As Long
  TmpPrev = Application.Calculation
  StkCalculation.Push TmpPrev
  If TmpPrev <> xlCalculationManual Then Application.Calculation = xlCalculationManual
End Sub

Private Sub PopCalculation()
  Dim TmpPrev As Variant
  TmpPrev = StkCalculation.Pop
  If Not IsNull(TmpPrev) Then
    If Application.Calculation <> CLng(TmpPrev) Then Application.Calculation = CLng(TmpPrev)
  End If
End Sub

Private Sub PushEnableEvents()
  Dim TmpPrev As Boolean
  TmpPrev = Application.EnableEvents
  StkEnableEvents.Push TmpPrev
  If TmpPrev <> False Then Application.EnableEvents = False
End Sub

Private Sub PopEnableEvents()
  Dim TmpPrev As Variant
  TmpPrev = StkEnableEvents.Pop
  If Not IsNull(TmpPrev) Then
    If Application.EnableEvents <> CBool(TmpPrev) Then Application.EnableEvents = CBool(TmpPrev)
  End If
End Sub

Private Sub PushDisplayAlerts()
  Dim TmpPrev As Boolean
  TmpPrev = Application.DisplayAlerts
  StkDisplayAlerts.Push TmpPrev
  If TmpPrev <> False Then Application.DisplayAlerts = False
End Sub

Private Sub PopDisplayAlerts()
  Dim TmpPrev As Variant
  TmpPrev = StkDisplayAlerts.Pop
  If Not IsNull(TmpPrev) Then
    If Application.DisplayAlerts <> CBool(TmpPrev) Then Application.DisplayAlerts = CBool(TmpPrev)
  End If
End Sub

Private Sub PushCursorWait()
  Dim TmpPrev As Long
  TmpPrev = Application.Cursor
  StkCursor.Push TmpPrev
  If TmpPrev <> xlWait Then Application.Cursor = xlWait
End Sub

Private Sub PopCursor()
  Dim TmpPrev As Variant
  TmpPrev = StkCursor.Pop
  If Not IsNull(TmpPrev) Then
    If Application.Cursor <> CLng(TmpPrev) Then Application.Cursor = CLng(TmpPrev)
  End If
End Sub


