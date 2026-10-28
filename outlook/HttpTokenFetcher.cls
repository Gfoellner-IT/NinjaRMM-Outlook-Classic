VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpTokenFetcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' HttpTokenFetcher v0.1.0
'
' ITokenFetcher on top of MSXML2.ServerXMLHTTP.6.0 (late bound, no reference needed).
' Uses the system proxy settings of the machine and explicit timeouts, so a missing
' network never blocks Outlook for long.
Implements ITokenFetcher
Option Explicit

' Timeouts in milliseconds: resolve, connect, send, receive.
Private Const TIMEOUT_RESOLVE As Long = 5000
Private Const TIMEOUT_CONNECT As Long = 5000
Private Const TIMEOUT_SEND As Long = 10000
Private Const TIMEOUT_RECEIVE As Long = 15000

Private Function ITokenFetcher_PostForm(ByVal url As String, ByVal body As String, ByRef httpStatus As Long) As String
    Dim http As Object

    On Error GoTo Failed

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts TIMEOUT_RESOLVE, TIMEOUT_CONNECT, TIMEOUT_SEND, TIMEOUT_RECEIVE
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.setRequestHeader "Accept", "application/json"
    http.send body

    httpStatus = CLng(http.Status)
    ITokenFetcher_PostForm = CStr(http.responseText)
    Exit Function

Failed:
    httpStatus = 0
    ITokenFetcher_PostForm = ""
End Function
