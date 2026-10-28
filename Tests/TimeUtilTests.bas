Attribute VB_Name = "TimeUtilTests"
' TimeUtilTests v0.1.0
'
' The exact offset depends on the time zone of the machine, so the tests check
' properties that hold everywhere.

Option Explicit

Public Sub Test_UtcToLocal_ZeroStaysZero()
    TestAssert.AssertEqual 0#, CDbl(TimeUtil.UtcToLocal(0))
End Sub

Public Sub Test_UtcToLocal_OffsetIsPlausible()
    Dim utc As Date
    Dim hours As Double

    utc = DateSerial(2026, 7, 1) + TimeSerial(12, 0, 0)
    hours = (TimeUtil.UtcToLocal(utc) - utc) * 24
    TestAssert.AssertTrue Abs(hours) <= 14, "summer offset within +-14 hours"

    utc = DateSerial(2026, 1, 15) + TimeSerial(12, 0, 0)
    hours = (TimeUtil.UtcToLocal(utc) - utc) * 24
    TestAssert.AssertTrue Abs(hours) <= 14, "winter offset within +-14 hours"
End Sub

Public Sub Test_UtcToLocal_KeepsMinutesAndSeconds()
    Dim utc As Date
    Dim localTime As Date

    utc = DateSerial(2026, 10, 6) + TimeSerial(14, 1, 28)
    localTime = TimeUtil.UtcToLocal(utc)
    TestAssert.AssertEqual 1, Minute(localTime), "minute"
    TestAssert.AssertEqual 28, Second(localTime), "second"
End Sub

Public Sub Test_UtcToLocal_IsMonotonic()
    Dim a As Date
    Dim b As Date

    a = DateSerial(2026, 3, 29) + TimeSerial(0, 30, 0)   ' around the spring clock change in Europe
    b = a + TimeSerial(6, 0, 0)
    TestAssert.AssertTrue TimeUtil.UtcToLocal(b) > TimeUtil.UtcToLocal(a)
End Sub
