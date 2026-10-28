VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITokenFetcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ITokenFetcher v0.1.0
'
' Seam between NinjaAuth and the network, so the token logic can be unit tested
' with a fake. The real implementation is HttpTokenFetcher.
Option Explicit

' Sends a form-encoded POST to url and returns the response body.
' httpStatus receives the HTTP status code, 0 if no response was received at all
' (no network, DNS failure, timeout).
Public Function PostForm(ByVal url As String, ByVal body As String, ByRef httpStatus As Long) As String
End Function
