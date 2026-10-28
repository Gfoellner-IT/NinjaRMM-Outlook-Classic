VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ReviewForm 
   Caption         =   "ReviewForm"
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11760
   OleObjectBlob   =   "ReviewForm.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "ReviewForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Review dialog (variant B of M8): a list of the proposals with check boxes. The user ticks mails and
' accepts or rejects them in one go; what is left when the form is closed stays a proposal.
' Controls (names matter): lblInfo (Label), lstProposals (ListBox, ListStyle = option, MultiSelect = multi),
' btnSelectAll, btnAccept, btnReject, btnClose (CommandButtons).
' All logic is in the Public procedures, so unit tests can drive the form without showing it.
' Source of this code: forms\ReviewForm.code.txt; tools\New-ReviewForm.ps1 builds forms\ReviewForm.frm from it.

Private pCount As Long
Private pDecisions() As Long     ' decision per proposal (index 1..pCount), 0 = still open
Private pRows() As Long          ' proposal index of every row still in the list (index 1..pRowCount)
Private pRowCount As Long

' Fills the list with the proposals (Collection of MatchProposal, oldest first).
Public Sub Setup(ByVal proposals As Collection)
    Dim i As Long

    pCount = proposals.Count
    ReDim pDecisions(1 To IIf(pCount > 0, pCount, 1))
    ReDim pRows(1 To IIf(pCount > 0, pCount, 1))
    pRowCount = pCount

    Me.Caption = MatchPromptText.DialogTitle()
    lblInfo.Caption = MatchPromptText.FormInfoText()
    btnSelectAll.Caption = MatchPromptText.FormSelectAllCaption()
    btnAccept.Caption = MatchPromptText.FormAcceptCaption()
    btnReject.Caption = MatchPromptText.FormRejectCaption()
    btnClose.Caption = MatchPromptText.FormCloseCaption()

    lstProposals.Clear
    For i = 1 To pCount
        pDecisions(i) = DecisionSkip
        pRows(i) = i
        lstProposals.AddItem MatchPromptText.ListLine(proposals(i), i, pCount)
    Next i
End Sub

' Ticks every row.
Public Sub SelectAll()
    Dim i As Long

    For i = 0 To lstProposals.ListCount - 1
        lstProposals.Selected(i) = True
    Next i
End Sub

' Number of ticked rows.
Public Function SelectedCount() As Long
    Dim i As Long

    For i = 0 To lstProposals.ListCount - 1
        If lstProposals.Selected(i) Then SelectedCount = SelectedCount + 1
    Next i
End Function

' Gives every ticked row this decision (DecisionAccept or DecisionReject) and removes it from the list.
' When no row is left the form closes itself.
Public Sub ApplySelected(ByVal decision As Long)
    Dim listIndex As Long
    Dim k As Long

    For listIndex = lstProposals.ListCount - 1 To 0 Step -1
        If lstProposals.Selected(listIndex) Then
            pDecisions(pRows(listIndex + 1)) = decision
            For k = listIndex + 1 To pRowCount - 1
                pRows(k) = pRows(k + 1)
            Next k
            pRowCount = pRowCount - 1
            lstProposals.RemoveItem listIndex
        End If
    Next listIndex

    If lstProposals.ListCount = 0 Then Me.Hide
End Sub

' Number of rows that were not decided yet.
Public Function OpenCount() As Long
    OpenCount = pRowCount
End Function

' The decisions in the order of the proposals (Collection of Long, 0 = skip).
Public Function Result() As Collection
    Dim list As Collection
    Dim i As Long

    Set list = New Collection
    For i = 1 To pCount
        list.Add pDecisions(i)
    Next i
    Set Result = list
End Function

Private Sub btnSelectAll_Click()
    SelectAll
End Sub

Private Sub btnAccept_Click()
    ApplySelected DecisionAccept
End Sub

Private Sub btnReject_Click()
    ApplySelected DecisionReject
End Sub

Private Sub btnClose_Click()
    Me.Hide
End Sub

' The close button of the title bar only hides the form, so the answers can still be read.
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        Me.Hide
    End If
End Sub

