VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OutlookTicketFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' OutlookTicketFolder v0.1.0
'
' Thin adapter: ITicketFolder on top of an Outlook.Folder. The decision logic is in
' TicketParsing/TicketService; this class only talks to Outlook.
Implements ITicketFolder
Option Explicit

Private pFolder As Outlook.Folder

Public Sub Init(ByVal folder As Outlook.Folder)
    Set pFolder = folder
End Sub

Public Property Get Raw() As Outlook.Folder
    Set Raw = pFolder
End Property

Private Property Get ITicketFolder_Name() As String
    ITicketFolder_Name = pFolder.Name
End Property

Private Property Let ITicketFolder_Name(ByVal value As String)
    pFolder.Name = value
End Property

Private Function ITicketFolder_FindClosedStatusMail(ByRef receivedTime As Date) As Boolean
    Dim i As Long
    Dim item As Object
    Dim found As Boolean

    For i = pFolder.Items.Count To 1 Step -1
        Set item = pFolder.Items(i)
        If item.Class = olMail Then
            If TicketParsing.IsClosedMail(item.Body) Then
                If Not found Or item.ReceivedTime > receivedTime Then receivedTime = item.ReceivedTime
                found = True
            End If
        End If
    Next i

    ITicketFolder_FindClosedStatusMail = found
End Function

Private Function ITicketFolder_RecentMails(ByVal since As Date) As Collection
    Dim result As Collection
    Dim items As Outlook.Items
    Dim item As Object
    Dim adapter As OutlookIncomingMail

    Set result = New Collection
    Set ITicketFolder_RecentMails = result

    Set items = pFolder.Items.Restrict("[ReceivedTime] >= '" & Format$(since, "ddddd h:nn AMPM") & "'")
    For Each item In items
        If item.Class = olMail Then
            Set adapter = New OutlookIncomingMail
            adapter.Init item
            result.Add adapter
        End If
    Next item
End Function
