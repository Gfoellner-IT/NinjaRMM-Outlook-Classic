VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FakeMatchPrompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' FakeMatchPrompt v0.1.0
'
' Test double for IMatchPrompt: answers the hint question with AnswerAskNow and returns the
' decisions set with SetDecisions (or Nothing when Cancel is set). It records what it was shown.
Implements IMatchPrompt
Option Explicit

Public AnswerAskNow As Boolean
Public Cancel As Boolean
Public AskCount As Long
Public ReviewCount As Long
Public LastPendingCount As Long
Public Shown As Collection

' If set, the fake calls MatchReview.OfferReview from inside Review (a mail event that arrives
' while the dialog is open).
Public ReentryStore As ITicketStore

Private pDecisions As Collection

Private Sub Class_Initialize()
    Set pDecisions = New Collection
End Sub

Public Sub SetDecisions(ParamArray values() As Variant)
    Dim i As Long

    Set pDecisions = New Collection
    For i = LBound(values) To UBound(values)
        pDecisions.Add CLng(values(i))
    Next i
End Sub

Private Function IMatchPrompt_AskReviewNow(ByVal pendingCount As Long) As Boolean
    AskCount = AskCount + 1
    LastPendingCount = pendingCount
    IMatchPrompt_AskReviewNow = AnswerAskNow
End Function

Private Function IMatchPrompt_Review(ByVal proposals As Collection) As Collection
    ReviewCount = ReviewCount + 1
    Set Shown = proposals

    If Not ReentryStore Is Nothing Then MatchReview.OfferReview Me, ReentryStore

    If Cancel Then Exit Function
    Set IMatchPrompt_Review = pDecisions
End Function
