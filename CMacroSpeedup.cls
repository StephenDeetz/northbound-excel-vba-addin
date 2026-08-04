VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMacroSpeedup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
Public Sub Init()
  StkScreenUpdating.Clear
  StkCalculation.Clear
  StkEnableEvents.Clear
  StkDisplayAlerts.Clear
  StkCursor.Clear
End Sub

Public Sub PushAll()
  PushScreenUpdating
  PushCalculation
  PushEnableEvents
  PushDisplayAlerts
  PushCursorWait
End Sub

Public Sub PopAll()
  PopCursor
  PopDisplayAlerts
  PopEnableEvents
  PopCalculation
  PopScreenUpdating
End Sub

Public Sub ClearAll()
  Do While StkCursor.Count > 0:         PopCursor:         Loop
  Do While StkDisplayAlerts.Count > 0:  PopDisplayAlerts:  Loop
  Do While StkEnableEvents.Count > 0:   PopEnableEvents:   Loop
  Do While StkCalculation.Count > 0:    PopCalculation:    Loop
  Do While StkScreenUpdating.Count > 0: PopScreenUpdating: Loop
End Sub

Public Sub PushItem(ByVal AItem As msMacroItems)
  Select Case AItem
    Case msScreenUpdating: PushScreenUpdating
    Case msCalculation:    PushCalculation
    Case msEnableEvents:   PushEnableEvents
    Case msDisplayAlerts:  PushDisplayAlerts
    Case msCursor:         PushCursorWait
  End Select
End Sub

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
  Dim Prev As Boolean
  Prev = Application.ScreenUpdating
  StkScreenUpdating.Push Prev
  If Prev <> False Then Application.ScreenUpdating = False
End Sub

Private Sub PopScreenUpdating()
  Dim Prev As Variant
  Prev = StkScreenUpdating.Pop
  If Not IsNull(Prev) Then
    If Application.ScreenUpdating <> CBool(Prev) Then Application.ScreenUpdating = CBool(Prev)
  End If
End Sub

Private Sub PushCalculation()
  Dim Prev As Long
  Prev = Application.Calculation
  StkCalculation.Push Prev
  If Prev <> xlCalculationManual Then Application.Calculation = xlCalculationManual
End Sub

Private Sub PopCalculation()
  Dim Prev As Variant
  Prev = StkCalculation.Pop
  If Not IsNull(Prev) Then
    If Application.Calculation <> CLng(Prev) Then Application.Calculation = CLng(Prev)
  End If
End Sub

Private Sub PushEnableEvents()
  Dim Prev As Boolean
  Prev = Application.EnableEvents
  StkEnableEvents.Push Prev
  If Prev <> False Then Application.EnableEvents = False
End Sub

Private Sub PopEnableEvents()
  Dim Prev As Variant
  Prev = StkEnableEvents.Pop
  If Not IsNull(Prev) Then
    If Application.EnableEvents <> CBool(Prev) Then Application.EnableEvents = CBool(Prev)
  End If
End Sub

Private Sub PushDisplayAlerts()
  Dim Prev As Boolean
  Prev = Application.DisplayAlerts
  StkDisplayAlerts.Push Prev
  If Prev <> False Then Application.DisplayAlerts = False
End Sub

Private Sub PopDisplayAlerts()
  Dim Prev As Variant
  Prev = StkDisplayAlerts.Pop
  If Not IsNull(Prev) Then
    If Application.DisplayAlerts <> CBool(Prev) Then Application.DisplayAlerts = CBool(Prev)
  End If
End Sub

Private Sub PushCursorWait()
  Dim Prev As Long
  Prev = Application.Cursor
  StkCursor.Push Prev
  If Prev <> xlWait Then Application.Cursor = xlWait
End Sub

Private Sub PopCursor()
  Dim Prev As Variant
  Prev = StkCursor.Pop
  If Not IsNull(Prev) Then
    If Application.Cursor <> CLng(Prev) Then Application.Cursor = CLng(Prev)
  End If
End Sub


