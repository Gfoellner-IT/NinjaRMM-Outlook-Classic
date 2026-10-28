VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileRejectionStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' FileRejectionStore v0.1.0
'
' IRejectionStore in the file rejected_matches.txt next to the audit log
' (%LOCALAPPDATA%\NinjaRMM-Outlook-Classic\audit): one line per rejected mail, "key yyyy-mm-dd"
' (ASCII). Only lines that start with an 8 digit hex key count; everything else is ignored.
' BaseFolder exists for unit tests.
Implements IRejectionStore
Option Explicit

Public BaseFolder As String

Private Const FILE_NAME As String = "rejected_matches.txt"

Public Function Folder() As String
    If Len(BaseFolder) > 0 Then
        Folder = BaseFolder
    Else
        Folder = Environ$("LOCALAPPDATA") & "\NinjaRMM-Outlook-Classic\audit"
    End If
End Function

Private Function IRejectionStore_LoadKeys() As Collection
    Dim result As Collection
    Dim fso As Object
    Dim stream As Object
    Dim line As String
    Dim key As String

    Set result = New Collection
    Set IRejectionStore_LoadKeys = result

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(Folder() & "\" & FILE_NAME) Then Exit Function

    Set stream = fso.OpenTextFile(Folder() & "\" & FILE_NAME, 1, False, 0)
    Do While Not stream.AtEndOfStream
        line = Trim$(stream.ReadLine)
        key = LCase$(Left$(line, 8))
        If IsKey(key) Then result.Add key
    Loop
    stream.Close
End Function

Private Sub IRejectionStore_AddKey(ByVal key As String)
    Dim fso As Object
    Dim stream As Object

    key = LCase$(Trim$(key))
    If Not IsKey(key) Then Exit Sub

    Set fso = CreateObject("Scripting.FileSystemObject")
    EnsureFolder fso, Folder()
    Set stream = fso.OpenTextFile(Folder() & "\" & FILE_NAME, 8, True, 0)   ' ForAppending, create, ASCII
    stream.WriteLine key & " " & Format$(Now, "yyyy-mm-dd")
    stream.Close
End Sub

Private Function IsKey(ByVal text As String) As Boolean
    IsKey = (Len(text) = 8) And (text Like "[0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f]")
End Function

Private Sub EnsureFolder(ByVal fso As Object, ByVal path As String)
    Dim parent As String

    If fso.FolderExists(path) Then Exit Sub
    parent = fso.GetParentFolderName(path)
    If Len(parent) > 0 Then EnsureFolder fso, parent
    fso.CreateFolder path
End Sub
