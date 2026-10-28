VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IRejectionStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' IRejectionStore v0.1.0
'
' Where the message keys of rejected proposals are remembered, so that a rejected mail is not
' proposed again (also after Outlook restarts). The real implementation is FileRejectionStore,
' unit tests use FakeRejectionStore.
Option Explicit

' All remembered keys (Collection of String).
Public Function LoadKeys() As Collection
End Function

' Remembers one more key. May raise an error; MatchRejections swallows it.
Public Sub AddKey(ByVal key As String)
End Sub
