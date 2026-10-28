Attribute VB_Name = "OutlookTestGuard"
' OutlookTestGuard v0.4.0
'
' Safety net for the integration tests: they may only touch the local test data file
' ("NinjaTest", a .pst that is not connected to any mail account), never the real mailbox.
' A store is accepted only if ALL of these are true:
'   - display name is "NinjaTest"
'   - it is a data file store that is not an Exchange store
'   - its file name ends with ".pst"
' The tests get their working folder only through FindTestTicketsFolder.
' LastProblem tells why no test folder was found (listing of all stores and the checks).

Option Explicit

Public Const TEST_STORE_NAME As String = "NinjaTest"
Public Const TEST_INBOX_NAME As String = "Posteingang"

Private pProblem As String

Public Property Get LastProblem() As String
    LastProblem = pProblem
End Property

Public Function IsTestStore(ByVal store As Outlook.Store) As Boolean
    IsTestStore = (Len(RejectReason(store)) = 0)
End Function

' "" if the store is the test store, otherwise the first failed check.
Public Function RejectReason(ByVal store As Outlook.Store) As String
    On Error GoTo Failed

    If store Is Nothing Then
        RejectReason = "no store"
        Exit Function
    End If
    If store.DisplayName <> TEST_STORE_NAME Then
        RejectReason = "display name is '" & store.DisplayName & "'"
        Exit Function
    End If
    If Not store.IsDataFileStore Then
        RejectReason = "not a data file store"
        Exit Function
    End If
    If store.ExchangeStoreType <> olNotExchange Then
        RejectReason = "ExchangeStoreType is " & CStr(store.ExchangeStoreType)
        Exit Function
    End If
    If LCase$(Right$(store.FilePath, 4)) <> ".pst" Then
        RejectReason = "file path is '" & store.FilePath & "'"
        Exit Function
    End If
    Exit Function

Failed:
    RejectReason = "error " & Err.Number & ": " & Err.Description
End Function

' The folder "Tickets" below the folder "Posteingang" of the test store, Nothing if the test store is
' missing, not safe, or has no such folders (see LastProblem). A new .pst has no registered default
' inbox (GetDefaultFolder fails), so the folders are looked up by name below the store root.
Public Function FindTestTicketsFolder() As Outlook.Folder
    Dim stores As Outlook.Stores
    Dim i As Long
    Dim reason As String
    Dim listing As String
    Dim inbox As Outlook.Folder
    Dim tickets As Outlook.Folder
    Dim candidate As Outlook.Store
    Dim problem As String

    pProblem = ""

    On Error GoTo Failed
    Set stores = Application.Session.Stores
    For i = 1 To stores.Count
        ' A store whose data file is unreachable (e.g. a network share while offline) raises an
        ' error already when it is fetched: skip it and go on.
        Set candidate = Nothing
        problem = ""
        On Error Resume Next
        Set candidate = stores.Item(i)
        If Err.Number <> 0 Then problem = Err.Description
        Err.Clear
        On Error GoTo Failed

        If candidate Is Nothing Then
            listing = listing & " [" & i & "] unreadable: " & problem & ";"
        Else
            reason = RejectReason(candidate)
            If Len(reason) = 0 Then
                Set inbox = ChildFolder(candidate.GetRootFolder, TEST_INBOX_NAME)
                If inbox Is Nothing Then
                    pProblem = "The test store has no folder '" & TEST_INBOX_NAME & "'."
                    Exit Function
                End If
                Set tickets = ChildFolder(inbox, TicketConfig.FOLDER_TICKETS)
                If tickets Is Nothing Then
                    pProblem = "The test store has no folder '" & TEST_INBOX_NAME & "\" & TicketConfig.FOLDER_TICKETS & "'."
                    Exit Function
                End If
                Set FindTestTicketsFolder = tickets
                Exit Function
            End If
            listing = listing & " [" & i & "] " & StoreName(candidate) & ": " & reason & ";"
        End If
    Next i

    pProblem = "No store accepted. Stores:" & listing
    Exit Function

Failed:
    pProblem = "Error " & Err.Number & " (" & Err.Description & ") after checking: " & listing
End Function

Private Function StoreName(ByVal store As Outlook.Store) As String
    On Error Resume Next
    StoreName = "'" & store.DisplayName & "'"
End Function

' Direct child folder by name (case-insensitive), Nothing if there is none.
Public Function ChildFolder(ByVal parent As Outlook.Folder, ByVal folderName As String) As Outlook.Folder
    Dim child As Outlook.Folder

    If parent Is Nothing Then Exit Function
    For Each child In parent.Folders
        If LCase$(child.Name) = LCase$(folderName) Then
            Set ChildFolder = child
            Exit Function
        End If
    Next child
End Function

' Deleted items folder of the test store (the registered default if present, else by name).
Public Function DeletedItemsOf(ByVal store As Outlook.Store) As Outlook.Folder
    On Error Resume Next
    Set DeletedItemsOf = store.GetDefaultFolder(olFolderDeletedItems)
    On Error GoTo 0
    If DeletedItemsOf Is Nothing Then Set DeletedItemsOf = ChildFolder(store.GetRootFolder, "Gel" & ChrW$(246) & "schte Elemente")
End Function
