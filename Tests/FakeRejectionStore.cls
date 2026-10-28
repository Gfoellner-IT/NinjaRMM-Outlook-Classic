VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FakeRejectionStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' FakeRejectionStore v0.1.0
'
' Test double for IRejectionStore: keeps the keys in memory. FailOnAdd simulates a broken file.
Implements IRejectionStore
Option Explicit

Public Keys As Collection
Public FailOnAdd As Boolean

Private Sub Class_Initialize()
    Set Keys = New Collection
End Sub

Private Function IRejectionStore_LoadKeys() As Collection
    Dim copy As Collection
    Dim key As Variant

    Set copy = New Collection
    For Each key In Keys
        copy.Add key
    Next key
    Set IRejectionStore_LoadKeys = copy
End Function

Private Sub IRejectionStore_AddKey(ByVal key As String)
    If FailOnAdd Then Err.Raise 76, "FakeRejectionStore", "path not found (simulated)"
    Keys.Add key
End Sub
