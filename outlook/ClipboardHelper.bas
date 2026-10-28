Attribute VB_Name = "ClipboardHelper"
' ClipboardHelper v0.1.0
'
' Puts text on the Windows clipboard without a VBA reference: first through the Forms 2.0 DataObject
' (late bound, keeps umlauts), otherwise through clip.exe (uses the ANSI code page). Not covered by
' the unit tests (it needs the desktop); checked by hand in Outlook.

Option Explicit

' True if the text is on the clipboard.
Public Function CopyText(ByVal text As String) As Boolean
    Dim data As Object
    Dim shell As Object
    Dim proc As Object

    On Error GoTo TryClip
    Set data = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    data.SetText text
    data.PutInClipboard
    CopyText = True
    Exit Function

TryClip:
    On Error GoTo Failed
    Set shell = CreateObject("WScript.Shell")
    Set proc = shell.Exec("cmd /c clip")
    proc.StdIn.Write text
    proc.StdIn.Close
    CopyText = True
    Exit Function

Failed:
    CopyText = False
End Function
