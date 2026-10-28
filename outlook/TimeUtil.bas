Attribute VB_Name = "TimeUtil"
' TimeUtil v0.1.0
'
' UTC to local time with the Windows time zone rules (including daylight saving
' time of the given date). Used so that tickets closed shortly before midnight end up
' in the right archive month.

Option Explicit

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Declare PtrSafe Function SystemTimeToTzSpecificLocalTime Lib "kernel32" ( _
    ByVal lpTimeZoneInformation As LongPtr, ByRef lpUniversalTime As SYSTEMTIME, _
    ByRef lpLocalTime As SYSTEMTIME) As Long

' Converts a UTC date to local time. 0 stays 0; if the conversion fails the input is returned.
Public Function UtcToLocal(ByVal utcTime As Date) As Date
    Dim utcParts As SYSTEMTIME
    Dim localParts As SYSTEMTIME

    If utcTime = 0 Then Exit Function
    UtcToLocal = utcTime

    utcParts.wYear = Year(utcTime)
    utcParts.wMonth = Month(utcTime)
    utcParts.wDay = Day(utcTime)
    utcParts.wHour = Hour(utcTime)
    utcParts.wMinute = Minute(utcTime)
    utcParts.wSecond = Second(utcTime)

    If SystemTimeToTzSpecificLocalTime(0, utcParts, localParts) = 0 Then Exit Function

    UtcToLocal = DateSerial(localParts.wYear, localParts.wMonth, localParts.wDay) + _
                 TimeSerial(localParts.wHour, localParts.wMinute, localParts.wSecond)
End Function
