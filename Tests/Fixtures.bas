Attribute VB_Name = "Fixtures"
' GENERATED FILE - do not edit. Source: Tests\fixtures\*.json
' Regenerate with: tools\New-FixturesModule.ps1
'
' Synthetic sample responses of the NinjaOne API used by the unit tests.

Option Explicit

Public Function BoardRunEmpty() As String
    Dim s As String
    s = s & "{""data"":[],""metadata"":{""lastCursorId"":0,""filters"":null}}"
    BoardRunEmpty = s
End Function

Public Function BoardRunPage1() As String
    Dim s As String
    s = s & "{""data"":[{""id"":301,""summary"":""Example printer problem in the warehouse"",""status"":{""statusId"":2000,""displayName"":""Offen""},""organization"":""Example Org"",""requester"
    s = s & """:""Example User"",""source"":""EMAIL"",""tags"":[]},{""id"":302,""summary"":""AW: TICKET (B\u00fcro - AT) / Example subject with prefix"",""status"":{""statusId"":1000,""displayN"
    s = s & "ame"":""Neu""},""organization"":""Example Org"",""requester"":""Example User"",""source"":""HELP_REQUEST"",""tags"":[]}],""metadata"":{""lastCursorId"":2,""filters"":null}}"
    BoardRunPage1 = s
End Function

Public Function BoardRunPage2() As String
    Dim s As String
    s = s & "{""data"":[{""id"":303,""summary"":""CSM | [TicketNr. 900000101] - Example vendor ticket"",""status"":{""statusId"":5000,""displayName"":""Aufgeloest""},""organization"":""Example"
    s = s & " Org"",""requester"":""Example User"",""source"":""TECHNICIAN"",""tags"":[""Example""]}],""metadata"":{""lastCursorId"":3,""filters"":null}}"
    BoardRunPage2 = s
End Function

Public Function BoardRunVendor() As String
    Dim s As String
    s = s & "{""data"":[{""id"":401,""summary"":""CSM | [TicketNr. 900000101] - Example vendor ticket"",""status"":{""statusId"":2000,""displayName"":""Offen""},""organization"":""Example Org"""
    s = s & ",""requester"":""Example User"",""source"":""TECHNICIAN"",""tags"":[""Example""]},{""id"":402,""summary"":""Older ticket for [TicketNr. 900000101]"",""status"":{""statusId"":6000,""di"
    s = s & "splayName"":""Geschlossen""},""organization"":""Example Org"",""requester"":""Example User"",""source"":""EMAIL"",""tags"":[]},{""id"":403,""summary"":""Unrelated ticket the filter s"
    s = s & "hould not have returned"",""status"":{""statusId"":2000,""displayName"":""Offen""},""organization"":""Example Org"",""requester"":""Example User"",""source"":""EMAIL"",""tags"":[]}],"""
    s = s & "metadata"":{""lastCursorId"":3,""filters"":null}}"
    BoardRunVendor = s
End Function

Public Function LogClosedAutomation() As String
    Dim s As String
    s = s & "[{""id"":3,""appUserContactUid"":null,""appUserContactId"":null,""appUserContactType"":null,""type"":""SAVE"",""body"":"""",""publicEntry"":false,""system"":true,""createTime"":17882"
    s = s & "70906.209405,""changeDiff"":{""status"":{""new"":""Geschlossen"",""old"":""Offen""}},""automation"":{""id"":1000,""name"":""Close Tickets Trigger""}},{""id"":2,""appUserContactUid"":nu"
    s = s & "ll,""appUserContactId"":null,""appUserContactType"":null,""type"":""SAVE"",""body"":"""",""publicEntry"":false,""system"":true,""createTime"":1787665335.253304,""changeDiff"":{""sta"
    s = s & "tus"":{""new"":""Offen"",""old"":""Neu""}},""automation"":{""id"":1006,""name"":""Set ticket status to open""}},{""id"":1,""appUserContactUid"":""00000000-0000-0000-0000-000000000001"
    s = s & """,""appUserContactId"":1,""appUserContactType"":""END_USER"",""type"":""SAVE"",""body"":"""",""publicEntry"":true,""system"":false,""createTime"":1787660000.000000,""changeDiff"":{""s"
    s = s & "tatus"":{""new"":""Neu"",""old"":null}},""automation"":null}]"
    LogClosedAutomation = s
End Function

Public Function LogClosedEnglish() As String
    Dim s As String
    s = s & "[{""id"":2,""type"":""SAVE"",""createTime"":1791295288.695002,""changeDiff"":{""status"":{""new"":""Closed"",""old"":""Open""}},""automation"":null},{""id"":1,""type"":""SAVE"",""createTime"
    s = s & """:1791294817.345961,""changeDiff"":{""status"":{""new"":""New"",""old"":null}},""automation"":null}]"
    LogClosedEnglish = s
End Function

Public Function LogClosedManual() As String
    Dim s As String
    s = s & "[{""id"":4,""appUserContactUid"":""00000000-0000-0000-0000-0000000000a1"",""appUserContactId"":100,""appUserContactType"":""TECHNICIAN"",""type"":""SAVE"",""body"":"""",""publicEntr"
    s = s & "y"":false,""system"":false,""createTime"":1791295288.695002,""changeDiff"":{""status"":{""new"":""Geschlossen"",""old"":""Neu""}},""automation"":null},{""id"":3,""appUserContactUid"":"
    s = s & "null,""appUserContactId"":null,""appUserContactType"":null,""type"":""SAVE"",""body"":"""",""publicEntry"":false,""system"":true,""createTime"":1791295288.680079,""changeDiff"":{""r"
    s = s & "elated"":[{""new"":[20110],""old"":[]}]},""automation"":null},{""id"":2,""appUserContactUid"":""00000000-0000-0000-0000-0000000000a1"",""appUserContactId"":100,""appUserContact"
    s = s & "Type"":""TECHNICIAN"",""type"":""SAVE"",""body"":"""",""publicEntry"":false,""system"":false,""createTime"":1791295226.616701,""changeDiff"":{""tags"":{""new"":[""ExampleTag""],""old"":[]"
    s = s & "},""assignedAppUserId"":{""new"":{""name"":""Technician A""},""old"":null}},""automation"":null},{""id"":1,""appUserContactUid"":""00000000-0000-0000-0000-000000000001"",""appUser"
    s = s & "ContactId"":1,""appUserContactType"":""END_USER"",""type"":""SAVE"",""body"":"""",""publicEntry"":true,""system"":false,""createTime"":1791294817.345961,""changeDiff"":{""type"":{""new"
    s = s & """:""PROBLEM"",""old"":null},""status"":{""new"":""Neu"",""old"":null},""summary"":{""new"":""Example"",""old"":null}},""automation"":null}]"
    LogClosedManual = s
End Function

Public Function LogDeleted() As String
    Dim s As String
    s = s & "[{""id"":2,""type"":""DELETE"",""createTime"":1791500000.750000,""changeDiff"":{},""automation"":null},{""id"":1,""type"":""SAVE"",""createTime"":1791100000.000000,""changeDiff"":{""s"
    s = s & "tatus"":{""new"":""Neu"",""old"":null}},""automation"":null}]"
    LogDeleted = s
End Function

Public Function LogOpenOnly() As String
    Dim s As String
    s = s & "[{""id"":2,""type"":""SAVE"",""createTime"":1791295000.000000,""changeDiff"":{""status"":{""new"":""Offen"",""old"":""Neu""}},""automation"":null},{""id"":1,""type"":""SAVE"",""createTime"":"
    s = s & "1791294817.345961,""changeDiff"":{""status"":{""new"":""Neu"",""old"":null}},""automation"":{""id"":1006,""name"":""Set ticket status to open""}}]"
    LogOpenOnly = s
End Function

Public Function LogReopened() As String
    Dim s As String
    s = s & "[{""id"":5,""type"":""SAVE"",""createTime"":1791400000.500000,""changeDiff"":{""status"":{""new"":""Geschlossen"",""old"":""Offen""}},""automation"":null},{""id"":4,""type"":""SAVE"",""crea"
    s = s & "teTime"":1791300000.250000,""changeDiff"":{""status"":{""new"":""Offen"",""old"":""Geschlossen""}},""automation"":null},{""id"":3,""type"":""SAVE"",""createTime"":1791200000.125000,""c"
    s = s & "hangeDiff"":{""status"":{""new"":""Geschlossen"",""old"":""Neu""}},""automation"":null},{""id"":1,""type"":""SAVE"",""createTime"":1791100000.000000,""changeDiff"":{""status"":{""new"":""N"
    s = s & "eu"",""old"":null}},""automation"":null}]"
    LogReopened = s
End Function

Public Function TicketClosed() As String
    Dim s As String
    s = s & "{""id"":20102,""version"":5,""nodeId"":null,""clientId"":1,""locationId"":null,""assignedAppUserId"":100,""requesterUid"":""00000000-0000-0000-0000-000000000001"",""subject"":""Ex"
    s = s & "ample"",""status"":{""name"":""CLOSED"",""displayName"":""Geschlossen"",""parentId"":null,""statusId"":6000},""type"":""PROBLEM"",""priority"":""NONE"",""severity"":""MAJOR"",""ticketFormI"
    s = s & "d"":1,""source"":""END_USER"",""tags"":[""ExampleTag""],""ccList"":{""uids"":null,""emails"":[]},""createTime"":1791294817.345961,""deleted"":false,""attributeValues"":[],""additiona"
    s = s & "lAssignedTechnicianIds"":[101,102],""followupTime"":null}"
    TicketClosed = s
End Function

Public Function TicketDeleted() As String
    Dim s As String
    s = s & "{""id"":12000,""version"":2,""clientId"":1,""subject"":""Example deleted ticket"",""status"":{""name"":""OPEN"",""displayName"":""Offen"",""parentId"":null,""statusId"":2000},""type"":""P"
    s = s & "ROBLEM"",""createTime"":1791000000.000000,""deleted"":true}"
    TicketDeleted = s
End Function

Public Function TicketNotFound() As String
    Dim s As String
    s = s & "{""resultCode"":""TICKET_NOT_FOUND"",""errorMessage"":""Ticket not found""}"
    TicketNotFound = s
End Function

Public Function TicketOpen() As String
    Dim s As String
    s = s & "{""id"":20101,""version"":3,""nodeId"":null,""clientId"":1,""locationId"":null,""assignedAppUserId"":null,""requesterUid"":""00000000-0000-0000-0000-000000000001"",""subject"":""E"
    s = s & "xample open ticket"",""status"":{""name"":""OPEN"",""displayName"":""Offen"",""parentId"":null,""statusId"":2000},""type"":""PROBLEM"",""priority"":""NONE"",""severity"":""MAJOR"",""ticket"
    s = s & "FormId"":1,""source"":""END_USER"",""tags"":[],""ccList"":{""uids"":null,""emails"":[]},""createTime"":1787665335.253304,""deleted"":false,""attributeValues"":[],""additionalAssign"
    s = s & "edTechnicianIds"":[],""followupTime"":null}"
    TicketOpen = s
End Function

Public Function TokenError() As String
    Dim s As String
    s = s & "{""error"":""invalid_client"",""error_description"":""Client authentication failed""}"
    TokenError = s
End Function

Public Function TokenResponse() As String
    Dim s As String
    s = s & "{""access_token"":""FAKE-ACCESS-TOKEN-FOR-TESTS"",""expires_in"":3600,""scope"":""monitoring"",""token_type"":""Bearer""}"
    TokenResponse = s
End Function
