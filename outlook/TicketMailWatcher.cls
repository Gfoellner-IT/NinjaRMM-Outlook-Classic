VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TicketMailWatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' TicketMailWatcher v0.1.0
'
' Sorts ticket mails into ticket folders as soon as they arrive in a folder (event Items.ItemAdd),
' without the rule action "Run a script" (which Outlook has disabled since 2017 unless the
' registry value EnableUnsafeClientMailRules is set).
'
' Limits of ItemAdd (Microsoft Learn): the event does not run when a large number of items
' arrive at once. Therefore a run of RunEmailRule at Outlook start (see
' Ticketautomatismus.StartMailWatcher) and from time to time sweeps up anything missed.
Option Explicit

Private WithEvents pItems As Outlook.Items
Private pInbox As Outlook.Folder
Private pStore As ITicketStore
Private pIndex As ITicketIndex
Private pPrompt As IMatchPrompt
Private pDelayMs As Long

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' Starts watching the folder. delayMs is the time to wait before a new mail is processed, so
' that other rules can move it first (0 in tests).
' index (optional): if given, mails of other senders are matched to existing tickets
' (TicketService.MatchForeignMail; acts according to the safety levels, see MatchLevels).
' prompt (optional): the dialogs of the confirmation level; after every new mail the hint question
' "review the new proposals now?" is offered when there are proposals (MatchReview.OfferReview).
Public Sub StartWatching(ByVal inbox As Outlook.Folder, ByVal store As ITicketStore, ByVal delayMs As Long, _
                         Optional ByVal index As ITicketIndex = Nothing, Optional ByVal prompt As IMatchPrompt = Nothing)
    Set pInbox = inbox
    Set pStore = store
    Set pIndex = index
    Set pPrompt = prompt
    pDelayMs = delayMs
    Set pItems = inbox.Items
End Sub

Public Sub StopWatching()
    Set pItems = Nothing
    Set pStore = Nothing
    Set pIndex = Nothing
    Set pPrompt = Nothing
    Set pInbox = Nothing
End Sub

Public Property Get IsWatching() As Boolean
    IsWatching = Not pItems Is Nothing
End Property

Private Sub pItems_ItemAdd(ByVal item As Object)
    Dim adapter As OutlookIncomingMail

    On Error GoTo Failed
    If pStore Is Nothing Then Exit Sub
    If Not TypeOf item Is Outlook.MailItem Then Exit Sub

    If pDelayMs > 0 Then Sleep pDelayMs

    ' A rule may have moved the mail in the meantime: then it is none of our business.
    If item.Parent.EntryID <> pInbox.EntryID Then Exit Sub

    Set adapter = New OutlookIncomingMail
    adapter.Init item
    If Not TicketService.ProcessMail(adapter, pStore) Then
        If Not pIndex Is Nothing Then TicketService.MatchForeignMail adapter, pIndex, pStore, True
    End If
    ' New proposals (also from the start-up sweep, which never asks) are offered for review here.
    If Not pPrompt Is Nothing Then MatchReview.OfferReview pPrompt, pStore
    Exit Sub

Failed:
    TicketLog.Fail "TicketMailWatcher: " & Err.Description
End Sub
