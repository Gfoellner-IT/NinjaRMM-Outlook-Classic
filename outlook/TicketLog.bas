Attribute VB_Name = "TicketLog"
' TicketLog v0.1.0
'
' Minimal logging for the ticket automation. Rules run unattended, so problems are
' logged instead of shown in dialogs. Every line goes to the Immediate window, to an
' in-memory list (for tests) and to %TEMP%\ninja_ticket_log.txt (cut at about 512 KB).
' Never log secrets.

Option Explicit

Private Const LOG_FILE_NAME As String = "ninja_ticket_log.txt"
Private Const MAX_FILE_BYTES As Long = 524288
Private Const MAX_MEMORY_LINES As Long = 500

Private pLines As Collection
Private pFileLoggingDisabled As Boolean

Public Sub Info(ByVal message As String)
    WriteLine "INFO  ", message
End Sub

Public Sub Warn(ByVal message As String)
    WriteLine "WARN  ", message
End Sub

Public Sub Fail(ByVal message As String)
    WriteLine "ERROR ", message
End Sub

' All lines logged since the last Clear, one per line (newest last).
Public Function Text() As String
    Dim item As Variant
    Dim result As String

    If pLines Is Nothing Then Exit Function
    For Each item In pLines
        result = result & CStr(item) & vbCrLf
    Next item
    Text = result
End Function

Public Sub Clear()
    Set pLines = New Collection
End Sub

' Unit tests switch the file off.
Public Sub DisableFileLogging()
    pFileLoggingDisabled = True
End Sub

' True after DisableFileLogging (unit tests). TicketAudit then writes no audit file either.
Public Function IsFileLoggingDisabled() As Boolean
    IsFileLoggingDisabled = pFileLoggingDisabled
End Function

Private Sub WriteLine(ByVal level As String, ByVal message As String)
    Dim line As String

    line = Format$(Now, "yyyy-mm-dd hh:nn:ss") & " " & level & message

    If pLines Is Nothing Then Set pLines = New Collection
    If pLines.Count >= MAX_MEMORY_LINES Then pLines.Remove 1
    pLines.Add line

    Debug.Print line
    If Not pFileLoggingDisabled Then AppendToFile line
End Sub

Private Sub AppendToFile(ByVal line As String)
    Dim fso As Object
    Dim stream As Object
    Dim path As String

    On Error GoTo Done
    path = Environ$("TEMP") & "\" & LOG_FILE_NAME
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(path) Then
        If fso.GetFile(path).Size > MAX_FILE_BYTES Then fso.DeleteFile path, True
    End If
    Set stream = fso.OpenTextFile(path, 8, True)   ' 8 = ForAppending
    stream.WriteLine line
    stream.Close
Done:
End Sub
