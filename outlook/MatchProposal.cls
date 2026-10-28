VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MatchProposal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' MatchProposal v0.1.0
'
' One suggested assignment of a mail to a ticket, waiting for the user's decision (level 2).
Option Explicit

Public Mail As IIncomingMail
Public TicketId As Long
Public TicketSummary As String
Public Rule As MatchRule
Public CreatedAt As Date
