VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "INinjaTransport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' INinjaTransport v0.1.0
'
' Seam between NinjaApi and the network, so the API logic can be unit tested with
' a fake. The real implementation is NinjaTransport.
Option Explicit

' Sends an authenticated GET request to the NinjaOne API.
' path is relative to the API base URL, e.g. "ticketing/ticket/123".
' Returns the response body. httpStatus receives the HTTP status code; it is 0 if no
' response was received (offline, timeout) or no token could be obtained, in which
' case errorText describes the problem (it never contains secrets).
Public Function HttpGet(ByVal path As String, ByRef httpStatus As Long, ByRef errorText As String) As String
End Function

' Sends an authenticated POST request with a JSON body (UTF-8). Used for the read-only
' board queries ("ticketing/trigger/board/{id}/run"); nothing is changed in NinjaOne.
' Status and error handling as for HttpGet.
Public Function HttpPostJson(ByVal path As String, ByVal jsonBody As String, ByRef httpStatus As Long, ByRef errorText As String) As String
End Function
