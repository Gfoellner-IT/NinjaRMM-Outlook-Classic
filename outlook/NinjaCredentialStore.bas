Attribute VB_Name = "NinjaCredentialStore"
' NinjaCredentialStore v0.1.0
'
' Reads generic credentials from the Windows Credential Manager (advapi32).
' The user name holds the client id, the password holds the client secret.
' Write and delete exist for the unit tests (round trip through the real API)
' and for scripted setup; production code only reads.
'
' Secrets must never be logged or shown in messages.

Option Explicit

Private Const CRED_TYPE_GENERIC As Long = 1
Private Const CRED_PERSIST_SESSION As Long = 1

' CREDENTIALW. VBA aligns the members like C does, so LenB gives sizeof (80 on 64 bit).
Private Type CREDENTIAL
    Flags As Long
    CredType As Long
    TargetName As LongPtr
    Comment As LongPtr
    LastWrittenLow As Long
    LastWrittenHigh As Long
    CredentialBlobSize As Long
    CredentialBlob As LongPtr
    Persist As Long
    AttributeCount As Long
    Attributes As LongPtr
    TargetAlias As LongPtr
    UserName As LongPtr
End Type

Private Declare PtrSafe Function CredReadW Lib "advapi32" (ByVal TargetName As LongPtr, ByVal CredType As Long, _
    ByVal Flags As Long, ByRef Credential As LongPtr) As Long
Private Declare PtrSafe Function CredWriteW Lib "advapi32" (ByRef Credential As CREDENTIAL, ByVal Flags As Long) As Long
Private Declare PtrSafe Function CredDeleteW Lib "advapi32" (ByVal TargetName As LongPtr, ByVal CredType As Long, _
    ByVal Flags As Long) As Long
Private Declare PtrSafe Sub CredFree Lib "advapi32" (ByVal Buffer As LongPtr)
Private Declare PtrSafe Sub CopyToCredential Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As CREDENTIAL, _
    ByVal Source As LongPtr, ByVal Length As LongPtr)
Private Declare PtrSafe Sub CopyMemoryRaw Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, _
    ByVal Source As LongPtr, ByVal Length As LongPtr)
Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long

' ---------------------------------------------------------------------------
'  Public API
' ---------------------------------------------------------------------------

' Reads the generic credential stored under target.
' Returns False if there is no such entry (userName and secret are then empty).
Public Function ReadCredential(ByVal target As String, ByRef userName As String, ByRef secret As String) As Boolean
    Dim pCredential As LongPtr
    Dim cred As CREDENTIAL

    userName = ""
    secret = ""

    If CredReadW(StrPtr(target), CRED_TYPE_GENERIC, 0, pCredential) = 0 Then Exit Function

    On Error GoTo Cleanup
    CopyToCredential cred, pCredential, LenB(cred)
    userName = PointerToString(cred.UserName)
    secret = BlobToString(cred.CredentialBlob, cred.CredentialBlobSize)
    ReadCredential = True

Cleanup:
    CredFree pCredential
End Function

' Writes (or replaces) a session-only generic credential. For tests and setup scripts.
Public Function WriteCredential(ByVal target As String, ByVal userName As String, ByVal secret As String) As Boolean
    Dim cred As CREDENTIAL
    Dim blob() As Byte

    If Len(secret) = 0 Then Exit Function
    blob = secret   ' UTF-16LE bytes, the same encoding the Credential Manager UI uses

    cred.CredType = CRED_TYPE_GENERIC
    cred.TargetName = StrPtr(target)
    cred.CredentialBlobSize = UBound(blob) + 1
    cred.CredentialBlob = VarPtr(blob(0))
    cred.Persist = CRED_PERSIST_SESSION
    cred.UserName = StrPtr(userName)

    WriteCredential = (CredWriteW(cred, 0) <> 0)
End Function

Public Function DeleteCredential(ByVal target As String) As Boolean
    DeleteCredential = (CredDeleteW(StrPtr(target), CRED_TYPE_GENERIC, 0) <> 0)
End Function

' ---------------------------------------------------------------------------
'  Helpers
' ---------------------------------------------------------------------------

Private Function PointerToString(ByVal pText As LongPtr) As String
    Dim length As Long
    Dim result As String

    If pText = 0 Then Exit Function
    length = lstrlenW(pText)
    If length = 0 Then Exit Function

    result = Space$(length)
    CopyMemoryRaw StrPtr(result), pText, length * 2
    PointerToString = result
End Function

Private Function BlobToString(ByVal pBlob As LongPtr, ByVal sizeInBytes As Long) As String
    Dim result As String

    If pBlob = 0 Or sizeInBytes < 2 Then Exit Function

    result = Space$(sizeInBytes \ 2)
    CopyMemoryRaw StrPtr(result), pBlob, (sizeInBytes \ 2) * 2
    BlobToString = result
End Function
