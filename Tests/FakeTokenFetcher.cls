VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FakeTokenFetcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' FakeTokenFetcher v0.1.0
'
' Test double for ITokenFetcher: returns a canned response and records the request.
Implements ITokenFetcher
Option Explicit

Public ResponseBody As String
Public ResponseStatus As Long
Public CallCount As Long
Public LastUrl As String
Public LastBody As String

Private Function ITokenFetcher_PostForm(ByVal url As String, ByVal body As String, ByRef httpStatus As Long) As String
    CallCount = CallCount + 1
    LastUrl = url
    LastBody = body
    httpStatus = ResponseStatus
    ITokenFetcher_PostForm = ResponseBody
End Function
