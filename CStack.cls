VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ============================
' Class: CStack
' Purpose: Standard LIFO stack with safe TryPop / TryPeek
'
'
' CStack members:
' Coll()               : Lazy accessor; initializes and returns backing Collection.
' Class_Terminate      : Releases backing Collection on object destruction.
' Push(AValue)         : Pushes a value onto the stack.
' Pop()                : Pops and returns top value; returns Null if empty.
' Peek()               : Returns top value without removing; returns Null if empty.
' Count()              : Returns number of items in stack.
' Clear()              : Empties stack by releasing Collection.
' IsEmpty()            : Returns True if stack has no items.
' TryPop(AValue,AErrStr): Safe pop; returns True on success, else False + error string.
' TryPeek(AValue,AErrStr): Safe peek; returns True on success, else False + error string.
'
' ============================

Option Explicit

Private CollObj As Collection


' ===== Guarded Accessor =====
Private Function Coll() As Collection
  If CollObj Is Nothing Then
    Set CollObj = New Collection
  End If
  Set Coll = CollObj
End Function


' ===== Cleanup =====
Private Sub Class_Terminate()
  Set CollObj = Nothing
End Sub


' ===== Core Methods =====
' Purpose: Push a value onto the top of the stack.
Public Sub Push(ByVal AValue As Variant)
  Coll.Add AValue
End Sub

' Purpose: Pop and return the top value; returns Null if the stack is empty.
Public Function Pop() As Variant
  Pop = Null
  
  If CollObj Is Nothing Then Exit Function
  If CollObj.Count = 0 Then Exit Function
  
  Pop = CollObj(CollObj.Count)
  CollObj.Remove CollObj.Count
  
  If CollObj.Count = 0 Then
    Set CollObj = Nothing
  End If
End Function

' Purpose: Return the top value without removing it; returns Null if the stack is empty.
Public Function Peek() As Variant
  Peek = Null
  
  If CollObj Is Nothing Then Exit Function
  If CollObj.Count = 0 Then Exit Function
  
  Peek = CollObj(CollObj.Count)
End Function

' Purpose: Return the number of items currently on the stack.
Public Function Count() As Long
  Count = 0
  If CollObj Is Nothing Then Exit Function
  Count = CollObj.Count
End Function

' Purpose: Empty the stack by releasing the backing Collection.
Public Sub Clear()
  If Not CollObj Is Nothing Then
    Set CollObj = Nothing
  End If
End Sub

' Purpose: Return True if the stack has no items.
Public Function IsEmpty() As Boolean
  IsEmpty = False
  If CollObj Is Nothing Then
    IsEmpty = True
  Else
    IsEmpty = (CollObj.Count = 0)
  End If
End Function


' ===== Safe Methods =====
' Purpose: Safely pop the top value; returns True on success, else False with an error string.
Public Function TryPop(ByRef AValue As Variant, _
                       ByRef AErrStr As String) As Boolean
  
  ' Reset outputs
  AErrStr = vbNullString
  AValue = Null

  If CollObj Is Nothing Then
    AErrStr = "Stack is empty."
    Exit Function
  End If

  If CollObj.Count = 0 Then
    AErrStr = "Stack is empty."
    Exit Function
  End If

  AValue = CollObj(CollObj.Count)
  CollObj.Remove CollObj.Count
  TryPop = True
End Function

' Purpose: Safely peek the top value; returns True on success, else False with an error string.
Public Function TryPeek(ByRef AValue As Variant, _
                        ByRef AErrStr As String) As Boolean
  
  ' Reset outputs
  AErrStr = vbNullString
  AValue = Null

  If CollObj Is Nothing Then
    AErrStr = "Stack is empty."
    Exit Function
  End If

  If CollObj.Count = 0 Then
    AErrStr = "Stack is empty."
    Exit Function
  End If

  AValue = CollObj(CollObj.Count)
  TryPeek = True
End Function





