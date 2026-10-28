Attribute VB_Name = "MatchRejections"
' MatchRejections v0.1.0
'
' Remembers which mails the user rejected as a proposal (by message key), so they are not proposed
' again. The keys are kept in memory and written through IRejectionStore (file) so they survive an
' Outlook restart. A mail without a message key cannot be remembered. While file logging is disabled
' (unit tests) no file is used unless a store is injected with UseStore. No Outlook types.

Option Explicit

Private pStore As IRejectionStore
Private pKeys As Object
Private pLoaded As Boolean

Public Function IsRejected(ByVal messageKey As String) As Boolean
    If Len(messageKey) = 0 Then Exit Function
    EnsureLoaded
    IsRejected = pKeys.Exists(messageKey)
End Function

Public Sub Reject(ByVal messageKey As String)
    Dim store As IRejectionStore

    If Len(messageKey) = 0 Then Exit Sub
    EnsureLoaded
    If pKeys.Exists(messageKey) Then Exit Sub
    pKeys.Add messageKey, True

    On Error Resume Next
    Set store = ActiveStore()
    If Not store Is Nothing Then store.AddKey messageKey
    If Err.Number <> 0 Then TicketLog.Warn "Rejection could not be saved (" & Err.Description & "); it only holds until Outlook is closed"
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------
'  Test seams
' ---------------------------------------------------------------------------

Public Sub UseStore(ByVal store As IRejectionStore)
    Set pStore = store
    pLoaded = False
    Set pKeys = Nothing
End Sub

Public Sub ResetState()
    Set pStore = Nothing
    pLoaded = False
    Set pKeys = Nothing
End Sub

' ---------------------------------------------------------------------------
'  Implementation
' ---------------------------------------------------------------------------

Private Function ActiveStore() As IRejectionStore
    If Not pStore Is Nothing Then
        Set ActiveStore = pStore
        Exit Function
    End If
    If TicketLog.IsFileLoggingDisabled() Then Exit Function

    Set pStore = New FileRejectionStore
    Set ActiveStore = pStore
End Function

Private Sub EnsureLoaded()
    Dim store As IRejectionStore
    Dim keys As Collection
    Dim key As Variant

    If pLoaded Then Exit Sub
    pLoaded = True
    Set pKeys = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    Set store = ActiveStore()
    If store Is Nothing Then Exit Sub
    Set keys = store.LoadKeys()
    If Err.Number <> 0 Then
        TicketLog.Warn "Rejections could not be read (" & Err.Description & ")"
        Exit Sub
    End If
    On Error GoTo 0

    For Each key In keys
        If Not pKeys.Exists(CStr(key)) Then pKeys.Add CStr(key), True
    Next key
End Sub
