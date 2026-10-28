Attribute VB_Name = "JsonLite"
' JsonLite v0.1.0
'
' Small JSON parser without references (late-bound Scripting.Dictionary only).
'
' Access: Parse/TryParse return the parsed value, GetString/TryGetNumber/GetBoolean/GetNode read by dotted path.
'
' Result types:
'   object  -> Scripting.Dictionary (case-sensitive keys, TypeName = "Dictionary")
'   array   -> Collection (1-based)
'   string  -> String
'   number  -> Long if it is an integer that fits, otherwise Double
'   true/false -> Boolean
'   null    -> Null
'
' Numbers are converted with Val, so the system decimal separator never matters.
' Parse raises ERR_JSON_INVALID for invalid input; TryParse returns False instead.
' The parser keeps its state in module variables and is not re-entrant.

Option Explicit

Public Const ERR_JSON_INVALID As Long = vbObjectError + 11060

Private pText As String
Private pPos As Long
Private pLen As Long

' ---------------------------------------------------------------------------
'  Public API
' ---------------------------------------------------------------------------

' Parses json and returns the value. Raises ERR_JSON_INVALID for invalid input.
Public Function Parse(ByVal json As String) As Variant
    Dim result As Variant

    ParseTo json, result
    If IsObject(result) Then Set Parse = result Else Parse = result
End Function

' Parses json into result. The caller must pass a fresh (Empty) Variant:
' assigning to a Variant that still holds an object would hit that object's default property.
Public Sub ParseTo(ByVal json As String, ByRef result As Variant)
    Dim parsed As Variant

    pText = json
    pPos = 1
    pLen = Len(json)

    On Error GoTo Failed
    ParseValue parsed
    SkipWhitespace
    If pPos <= pLen Then Fail "unexpected text after the JSON value"
    On Error GoTo 0

    pText = ""
    If IsObject(parsed) Then Set result = parsed Else result = parsed
    Exit Sub

Failed:
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    pText = ""
    Err.Raise errNumber, errSource, errDescription
End Sub

' Like ParseTo, but returns False instead of raising. result stays Empty on failure.
Public Function TryParse(ByVal json As String, ByRef result As Variant) As Boolean
    On Error GoTo Failed
    ParseTo json, result
    TryParse = True
    Exit Function
Failed:
    TryParse = False
End Function

' ---------------------------------------------------------------------------
'  Typed access by dotted path, e.g. GetString(root, "status.displayName")
'  Missing keys, null and values of another type give the default / False / Nothing.
' ---------------------------------------------------------------------------

Public Function GetString(ByVal root As Variant, ByVal path As String, Optional ByVal defaultValue As String = "") As String
    Dim v As Variant

    GetString = defaultValue
    If Not Walk(root, path, v) Then Exit Function
    If VarType(v) = vbString Then GetString = CStr(v)
End Function

Public Function TryGetNumber(ByVal root As Variant, ByVal path As String, ByRef value As Double) As Boolean
    Dim v As Variant

    value = 0
    If Not Walk(root, path, v) Then Exit Function
    Select Case VarType(v)
        Case vbInteger, vbLong, vbDouble, vbSingle
            value = CDbl(v)
            TryGetNumber = True
    End Select
End Function

Public Function GetBoolean(ByVal root As Variant, ByVal path As String, Optional ByVal defaultValue As Boolean = False) As Boolean
    Dim v As Variant

    GetBoolean = defaultValue
    If Not Walk(root, path, v) Then Exit Function
    If VarType(v) = vbBoolean Then GetBoolean = CBool(v)
End Function

' Dictionary or Collection at path, Nothing otherwise.
Public Function GetNode(ByVal root As Variant, ByVal path As String) As Object
    Dim v As Variant

    If Not Walk(root, path, v) Then Exit Function
    If IsObject(v) Then Set GetNode = v
End Function

' True if the key exists, even when its value is null.
Public Function HasPath(ByVal root As Variant, ByVal path As String) As Boolean
    Dim v As Variant

    HasPath = Walk(root, path, v)
End Function

' Walks the dotted path through nested objects. result must be a fresh Variant.
Private Function Walk(ByVal root As Variant, ByVal path As String, ByRef result As Variant) As Boolean
    Dim parts() As String
    Dim current As Object
    Dim i As Long

    If Not IsObject(root) Then Exit Function
    Set current = root
    parts = Split(path, ".")

    For i = LBound(parts) To UBound(parts)
        If TypeName(current) <> "Dictionary" Then Exit Function
        If Not current.Exists(parts(i)) Then Exit Function

        If i = UBound(parts) Then
            If IsObject(current.Item(parts(i))) Then
                Set result = current.Item(parts(i))
            Else
                result = current.Item(parts(i))
            End If
            Walk = True
            Exit Function
        End If

        If Not IsObject(current.Item(parts(i))) Then Exit Function
        Set current = current.Item(parts(i))
    Next i
End Function

' ---------------------------------------------------------------------------
'  Parser
' ---------------------------------------------------------------------------

Private Sub ParseValue(ByRef result As Variant)
    Dim ch As String

    SkipWhitespace
    If pPos > pLen Then Fail "unexpected end of input"
    ch = Mid$(pText, pPos, 1)

    Select Case ch
        Case "{"
            ParseObject result
        Case "["
            ParseArray result
        Case """"
            result = ParseString()
        Case "t"
            ExpectLiteral "true"
            result = True
        Case "f"
            ExpectLiteral "false"
            result = False
        Case "n"
            ExpectLiteral "null"
            result = Null
        Case Else
            result = ParseNumber()
    End Select
End Sub

Private Sub ParseObject(ByRef result As Variant)
    Dim dict As Object

    Set dict = CreateObject("Scripting.Dictionary")
    pPos = pPos + 1
    SkipWhitespace

    If Mid$(pText, pPos, 1) = "}" Then
        pPos = pPos + 1
        Set result = dict
        Exit Sub
    End If

    Do
        ParseMember dict
        SkipWhitespace
        Select Case Mid$(pText, pPos, 1)
            Case ","
                pPos = pPos + 1
            Case "}"
                pPos = pPos + 1
                Exit Do
            Case Else
                Fail "',' or '}' expected"
        End Select
    Loop

    Set result = dict
End Sub

' Separate procedure so that "value" is a fresh, empty Variant for every member
' (assigning a scalar to a Variant that still holds an object would hit its default property).
Private Sub ParseMember(ByVal dict As Object)
    Dim key As String
    Dim value As Variant

    SkipWhitespace
    If Mid$(pText, pPos, 1) <> """" Then Fail "object key expected"
    key = ParseString()

    SkipWhitespace
    If Mid$(pText, pPos, 1) <> ":" Then Fail "':' expected"
    pPos = pPos + 1

    ParseValue value
    If IsObject(value) Then
        Set dict.Item(key) = value
    Else
        dict.Item(key) = value
    End If
End Sub

Private Sub ParseArray(ByRef result As Variant)
    Dim items As Collection

    Set items = New Collection
    pPos = pPos + 1
    SkipWhitespace

    If Mid$(pText, pPos, 1) = "]" Then
        pPos = pPos + 1
        Set result = items
        Exit Sub
    End If

    Do
        ParseElement items
        SkipWhitespace
        Select Case Mid$(pText, pPos, 1)
            Case ","
                pPos = pPos + 1
            Case "]"
                pPos = pPos + 1
                Exit Do
            Case Else
                Fail "',' or ']' expected"
        End Select
    Loop

    Set result = items
End Sub

Private Sub ParseElement(ByVal items As Collection)
    Dim value As Variant

    ParseValue value
    items.Add value
End Sub

Private Function ParseString() As String
    Dim quotePos As Long
    Dim segment As String
    Dim backslash As Long
    Dim escapePos As Long
    Dim escapeChar As String
    Dim hex4 As String
    Dim result As String

    pPos = pPos + 1   ' opening quote

    Do
        quotePos = InStr(pPos, pText, """")
        If quotePos = 0 Then Fail "unterminated string"

        segment = Mid$(pText, pPos, quotePos - pPos)
        backslash = InStr(segment, "\")

        If backslash = 0 Then
            result = result & segment
            pPos = quotePos + 1
            Exit Do
        End If

        result = result & Left$(segment, backslash - 1)
        escapePos = pPos + backslash - 1
        escapeChar = Mid$(pText, escapePos + 1, 1)

        Select Case escapeChar
            Case """", "\", "/"
                result = result & escapeChar
                pPos = escapePos + 2
            Case "b"
                result = result & vbBack
                pPos = escapePos + 2
            Case "f"
                result = result & Chr$(12)
                pPos = escapePos + 2
            Case "n"
                result = result & vbLf
                pPos = escapePos + 2
            Case "r"
                result = result & vbCr
                pPos = escapePos + 2
            Case "t"
                result = result & vbTab
                pPos = escapePos + 2
            Case "u"
                hex4 = Mid$(pText, escapePos + 2, 4)
                If Len(hex4) < 4 Then Fail "invalid \u escape"
                If Not (hex4 Like "[0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f]") Then Fail "invalid \u escape"
                result = result & ChrW$(CLng(Val("&H" & hex4 & "&")))
                pPos = escapePos + 6
            Case Else
                pPos = escapePos
                Fail "invalid escape sequence"
        End Select
    Loop

    ParseString = result
End Function

Private Function ParseNumber() As Variant
    Dim startPos As Long
    Dim token As String
    Dim ch As String
    Dim value As Double

    startPos = pPos
    Do While pPos <= pLen
        ch = Mid$(pText, pPos, 1)
        If InStr("0123456789+-.eE", ch) = 0 Then Exit Do
        pPos = pPos + 1
    Loop

    token = Mid$(pText, startPos, pPos - startPos)
    If Len(token) = 0 Then Fail "value expected"
    If Not (Left$(token, 1) Like "[0-9-]") Then Fail "invalid number"

    value = Val(token)
    If InStr(token, ".") = 0 And InStr(LCase$(token), "e") = 0 And Abs(value) < 2147483647# Then
        ParseNumber = CLng(value)
    Else
        ParseNumber = value
    End If
End Function

Private Sub ExpectLiteral(ByVal literal As String)
    If Mid$(pText, pPos, Len(literal)) <> literal Then Fail "invalid literal"
    pPos = pPos + Len(literal)
End Sub

Private Sub SkipWhitespace()
    Dim ch As String

    Do While pPos <= pLen
        ch = Mid$(pText, pPos, 1)
        If ch <> " " And ch <> vbTab And ch <> vbLf And ch <> vbCr Then Exit Do
        pPos = pPos + 1
    Loop
End Sub

Private Sub Fail(ByVal message As String)
    Err.Raise ERR_JSON_INVALID, "JsonLite.Parse", "Invalid JSON at position " & pPos & ": " & message
End Sub
