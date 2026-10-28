VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MatchTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' MatchTicket v0.1.0
'
' Minimal ticket facts the TicketMatcher needs (value object, no behaviour).
Option Explicit

Public Id As Long
Public Summary As String
Public IsOpen As Boolean
