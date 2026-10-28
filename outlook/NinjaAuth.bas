Attribute VB_Name = "NinjaAuth"
' NinjaAuth v0.2.0
'
' OAuth2 client-credentials authentication for the NinjaOne API.
' No browser, no listener, no refresh token: the client id and secret are read from
' the Windows Credential Manager and exchanged for a short-lived access token that is
' cached in memory until shortly before it expires.
'
' Credential Manager entry (generic credential):
'   Internet or network address: NINJA_CREDENTIAL_TARGET
'   User name:                   client id of the "API Services" application
'   Password:                    client secret
'
' Errors raised by GetAccessToken (never contain the secret):
'   ERR_NO_CREDENTIAL    no usable credential entry
'   ERR_TOKEN_REQUEST    HTTP error or no response (see LastHttpStatus, 0 = no response)
'   ERR_TOKEN_RESPONSE   200 response that is not a valid token response

Option Explicit

Public Const NINJA_TOKEN_URL As String = "https://eu.ninjarmm.com/ws/oauth/token"
Public Const NINJA_SCOPE As String = "monitoring"

Public Const ERR_NO_CREDENTIAL As Long = vbObjectError + 11050
Public Const ERR_TOKEN_REQUEST As Long = vbObjectError + 11051
Public Const ERR_TOKEN_RESPONSE As Long = vbObjectError + 11052

Private pFetcher As ITokenFetcher
Private pCredentialTarget As String
Private pToken As String
Private pExpiresAt As Date
Private pLastHttpStatus As Long

Public Function NINJA_API_BASE_URL() As String
    NINJA_API_BASE_URL = LocalSettings.ApiBaseUrl()
End Function
Public Function NINJA_CREDENTIAL_TARGET() As String
    NINJA_CREDENTIAL_TARGET = LocalSettings.Value("CredentialTarget", "NinjaRMM-Outlook-Classic/API")
End Function

' ---------------------------------------------------------------------------
'  Public API
' ---------------------------------------------------------------------------

' Returns a valid access token, requesting a new one when necessary.
Public Function GetAccessToken() As String
    If Not NinjaTokenLogic.IsTokenValid(pToken, pExpiresAt, Now) Then
        RequestNewToken
    End If
    GetAccessToken = pToken
End Function

' Forgets the cached token (e.g. after a 401 from the API).
Public Sub InvalidateToken()
    pToken = ""
    pExpiresAt = 0
End Sub

' HTTP status of the last token request (0 = no response received).
Public Property Get LastHttpStatus() As Long
    LastHttpStatus = pLastHttpStatus
End Property

' True if the error means "could not reach the token endpoint" (offline, timeout,
' server error) as opposed to a configuration or credential problem.
Public Function IsServiceUnavailable(ByVal errNumber As Long) As Boolean
    IsServiceUnavailable = (errNumber = ERR_TOKEN_REQUEST) And _
        (pLastHttpStatus = 0 Or pLastHttpStatus >= 500)
End Function

' ---------------------------------------------------------------------------
'  Test seams
' ---------------------------------------------------------------------------

Public Sub UseFetcher(ByVal fetcher As ITokenFetcher)
    Set pFetcher = fetcher
End Sub

Public Sub UseCredentialTarget(ByVal target As String)
    pCredentialTarget = target
End Sub

' Back to production defaults and empty cache.
Public Sub ResetState()
    Set pFetcher = Nothing
    pCredentialTarget = ""
    pLastHttpStatus = 0
    InvalidateToken
End Sub

' ---------------------------------------------------------------------------
'  Implementation
' ---------------------------------------------------------------------------

Private Sub RequestNewToken()
    Dim target As String
    Dim clientId As String
    Dim clientSecret As String
    Dim body As String
    Dim response As String
    Dim status As Long
    Dim accessToken As String
    Dim expiresIn As Long
    Dim grantedScope As String
    Dim detail As String

    InvalidateToken

    target = NINJA_CREDENTIAL_TARGET
    If Len(pCredentialTarget) > 0 Then target = pCredentialTarget

    If Not NinjaCredentialStore.ReadCredential(target, clientId, clientSecret) Or _
       Len(clientId) = 0 Or Len(clientSecret) = 0 Then
        Err.Raise ERR_NO_CREDENTIAL, "NinjaAuth.GetAccessToken", _
            "No client id/secret found in the Windows Credential Manager entry '" & target & "'."
    End If

    If pFetcher Is Nothing Then Set pFetcher = New HttpTokenFetcher

    body = NinjaTokenLogic.BuildTokenRequestBody(clientId, clientSecret, NINJA_SCOPE)
    response = pFetcher.PostForm(NINJA_TOKEN_URL, body, status)
    pLastHttpStatus = status

    If status <> 200 Then
        If status = 0 Then
            detail = "no response from the token endpoint"
        Else
            detail = "HTTP " & status
            If Len(NinjaTokenLogic.ParseErrorResponse(response)) > 0 Then
                detail = detail & " - " & NinjaTokenLogic.ParseErrorResponse(response)
            End If
        End If
        Err.Raise ERR_TOKEN_REQUEST, "NinjaAuth.GetAccessToken", "Token request failed: " & detail
    End If

    If Not NinjaTokenLogic.ParseTokenResponse(response, accessToken, expiresIn, grantedScope) Then
        Err.Raise ERR_TOKEN_RESPONSE, "NinjaAuth.GetAccessToken", "The token response could not be read."
    End If

    pToken = accessToken
    pExpiresAt = NinjaTokenLogic.ComputeExpiry(Now, expiresIn)
End Sub
