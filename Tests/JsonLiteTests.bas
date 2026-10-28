Attribute VB_Name = "JsonLiteTests"
' JsonLiteTests v0.1.0
'
' Unit tests for the JSON parser.

Option Explicit

Private Function ParseError(ByVal json As String) As Long
    Dim result As Variant

    On Error Resume Next
    JsonLite.ParseTo json, result
    ParseError = Err.Number
    Err.Clear
End Function

Public Sub Test_Json_Scalars()
    TestAssert.AssertEqual "abc", JsonLite.Parse("""abc"""), "string"
    TestAssert.AssertEqual "", JsonLite.Parse(""""""), "empty string"
    TestAssert.AssertEqual 42, JsonLite.Parse("42"), "integer"
    TestAssert.AssertEqual -7, JsonLite.Parse("-7"), "negative integer"
    TestAssert.AssertTrue JsonLite.Parse("true"), "true"
    TestAssert.AssertFalse JsonLite.Parse("false"), "false"
    TestAssert.AssertTrue IsNull(JsonLite.Parse("null")), "null"
    TestAssert.AssertTrue IsNull(JsonLite.Parse("  null  ")), "whitespace around the value"
End Sub

Public Sub Test_Json_NumberTypes()
    Dim d As Double

    TestAssert.AssertEqual "Long", TypeName(JsonLite.Parse("6000")), "integer fits a Long"
    TestAssert.AssertEqual "Double", TypeName(JsonLite.Parse("1.5")), "fraction is a Double"
    TestAssert.AssertEqual "Double", TypeName(JsonLite.Parse("3000000000")), "too big for a Long"
    TestAssert.AssertEqual "Double", TypeName(JsonLite.Parse("1e3")), "exponent is a Double"
    d = JsonLite.Parse("1791295288.695002")
    TestAssert.AssertTrue Abs(d - 1791295288.695002) < 0.000001, "decimal point independent of the system locale"
    TestAssert.AssertEqual 1000, JsonLite.Parse("1e3")
    TestAssert.AssertEqual 0.25, JsonLite.Parse("2.5E-1")
End Sub

Public Sub Test_Json_StringEscapes()
    TestAssert.AssertEqual "a""b", JsonLite.Parse("""a\""b"""), "escaped quote"
    TestAssert.AssertEqual "a\b", JsonLite.Parse("""a\\b"""), "escaped backslash"
    TestAssert.AssertEqual "a/b", JsonLite.Parse("""a\/b"""), "escaped slash"
    TestAssert.AssertEqual "a" & vbLf & "b" & vbTab & "c", JsonLite.Parse("""a\nb\tc"""), "newline and tab"
    TestAssert.AssertEqual "x" & ChrW$(252) & "y", JsonLite.Parse("""x\" & "u00fcy"""), "unicode escape"
    TestAssert.AssertEqual ChrW$(8364), JsonLite.Parse("""\" & "u20AC"""), "unicode escape (upper case hex)"
    TestAssert.AssertEqual "a\", JsonLite.Parse("""a\\"""), "string ending with an escaped backslash"
    TestAssert.AssertEqual "end""", JsonLite.Parse("""end\"""""), "string ending with an escaped quote"
End Sub

Public Sub Test_Json_NestedStructures()
    Dim root As Object

    Set root = JsonLite.Parse("{""status"":{""name"":""CLOSED"",""statusId"":6000},""tags"":[""a"",""b""],""deleted"":false,""x"":null}")

    TestAssert.AssertEqual "Dictionary", TypeName(root)
    TestAssert.AssertEqual "CLOSED", root("status")("name")
    TestAssert.AssertEqual 6000, root("status")("statusId")
    TestAssert.AssertEqual "Collection", TypeName(root("tags"))
    TestAssert.AssertEqual 2, root("tags").Count
    TestAssert.AssertEqual "b", root("tags")(2)
    TestAssert.AssertFalse root("deleted")
    TestAssert.AssertTrue IsNull(root("x"))
    TestAssert.AssertTrue root.Exists("x"), "null member exists"
    TestAssert.AssertFalse root.Exists("y"), "unknown member"
End Sub

Public Sub Test_Json_EmptyContainersAndWhitespace()
    TestAssert.AssertEqual 0, JsonLite.Parse("{}").Count
    TestAssert.AssertEqual 0, JsonLite.Parse("[]").Count
    TestAssert.AssertEqual 2, JsonLite.Parse(" [ 1 ,"" "" ] ").Count
    TestAssert.AssertEqual 1, JsonLite.Parse(vbCrLf & "{" & vbCrLf & vbTab & """a"" : 1" & vbCrLf & "}" & vbCrLf).Count
End Sub

Public Sub Test_Json_ArrayOfObjects()
    Dim list As Object

    Set list = JsonLite.Parse("[{""id"":1},{""id"":2},{""id"":3}]")
    TestAssert.AssertEqual 3, list.Count
    TestAssert.AssertEqual 3, list(3)("id")
End Sub

Public Sub Test_Json_KeysAreCaseSensitiveAndLastDuplicateWins()
    Dim root As Object

    Set root = JsonLite.Parse("{""a"":1,""A"":2,""a"":3}")
    TestAssert.AssertEqual 2, root.Count
    TestAssert.AssertEqual 3, root("a"), "last duplicate wins"
    TestAssert.AssertEqual 2, root("A")
End Sub

Public Sub Test_Json_InvalidInput()
    TestAssert.AssertError JsonLite.ERR_JSON_INVALID, ParseError(""), "empty"
    TestAssert.AssertError JsonLite.ERR_JSON_INVALID, ParseError("   "), "blank"
    TestAssert.AssertError JsonLite.ERR_JSON_INVALID, ParseError("{"), "open object"
    TestAssert.AssertError JsonLite.ERR_JSON_INVALID, ParseError("[1,]"), "trailing comma in array"
    TestAssert.AssertError JsonLite.ERR_JSON_INVALID, ParseError("{""a"":}"), "missing value"
    TestAssert.AssertError JsonLite.ERR_JSON_INVALID, ParseError("{""a"" 1}"), "missing colon"
    TestAssert.AssertError JsonLite.ERR_JSON_INVALID, ParseError("{a:1}"), "unquoted key"
    TestAssert.AssertError JsonLite.ERR_JSON_INVALID, ParseError("tru"), "short literal"
    TestAssert.AssertError JsonLite.ERR_JSON_INVALID, ParseError("""abc"), "unterminated string"
    TestAssert.AssertError JsonLite.ERR_JSON_INVALID, ParseError("""a\qb"""), "invalid escape"
    TestAssert.AssertError JsonLite.ERR_JSON_INVALID, ParseError("""\u12"""), "short unicode escape"
    TestAssert.AssertError JsonLite.ERR_JSON_INVALID, ParseError("1 2"), "text after the value"
    TestAssert.AssertError JsonLite.ERR_JSON_INVALID, ParseError("not json"), "garbage"
    TestAssert.AssertError JsonLite.ERR_JSON_INVALID, ParseError("<html></html>"), "html"
End Sub

Public Sub Test_Json_TryParse()
    Dim result As Variant
    Dim failed As Variant

    TestAssert.AssertTrue JsonLite.TryParse("{""a"":1}", result)
    TestAssert.AssertEqual 1, result("a")
    TestAssert.AssertFalse JsonLite.TryParse("{""a"":", failed)
    TestAssert.AssertTrue IsEmpty(failed), "result stays Empty on failure"
End Sub

Public Sub Test_Json_ParserCanBeReused()
    ' After a failed parse the next parse must start clean.
    TestAssert.AssertError JsonLite.ERR_JSON_INVALID, ParseError("{""a"":")
    TestAssert.AssertEqual 5, JsonLite.Parse("5")
End Sub

Public Sub Test_Json_PathAccess()
    Dim root As Variant
    Dim d As Double

    JsonLite.ParseTo "{""status"":{""name"":""CLOSED"",""statusId"":6000},""deleted"":true,""n"":null,""s"":""x""}", root

    TestAssert.AssertEqual "CLOSED", JsonLite.GetString(root, "status.name")
    TestAssert.AssertEqual "dflt", JsonLite.GetString(root, "status.missing", "dflt"), "missing key"
    TestAssert.AssertEqual "dflt", JsonLite.GetString(root, "n", "dflt"), "null gives the default"
    TestAssert.AssertEqual "dflt", JsonLite.GetString(root, "status.statusId", "dflt"), "number is not a string"
    TestAssert.AssertEqual "dflt", JsonLite.GetString(root, "s.deeper", "dflt"), "path through a scalar"

    TestAssert.AssertTrue JsonLite.TryGetNumber(root, "status.statusId", d)
    TestAssert.AssertEqual 6000, d
    TestAssert.AssertFalse JsonLite.TryGetNumber(root, "status.name", d), "string is not a number"
    TestAssert.AssertFalse JsonLite.TryGetNumber(root, "n", d), "null is not a number"

    TestAssert.AssertTrue JsonLite.GetBoolean(root, "deleted", False)
    TestAssert.AssertFalse JsonLite.GetBoolean(root, "missing", False)
    TestAssert.AssertTrue JsonLite.GetBoolean(root, "missing", True), "default"

    TestAssert.AssertTrue JsonLite.HasPath(root, "n"), "null member exists"
    TestAssert.AssertFalse JsonLite.HasPath(root, "status.nope")
    TestAssert.AssertTrue Not JsonLite.GetNode(root, "status") Is Nothing, "object node"
    TestAssert.AssertTrue JsonLite.GetNode(root, "s") Is Nothing, "scalar is not a node"
End Sub

Public Sub Test_Json_LargeDocument()
    Dim parts() As String
    Dim i As Long
    Dim list As Object

    ReDim parts(1 To 3000)
    For i = 1 To 3000
        parts(i) = "{""id"":" & i & ",""name"":""item " & i & """,""tags"":[""a"",""b""],""note"":""line\nbreak""}"
    Next i

    Set list = JsonLite.Parse("[" & Join(parts, ",") & "]")
    TestAssert.AssertEqual 3000, list.Count
    TestAssert.AssertEqual "item 3000", list(3000)("name")
End Sub
