Attribute VB_Name = "modWait"

Public Const WAIT_SHORT_SECONDS As Double = 10
Public Const WAIT_DEFAULT_SECONDS As Double = 60
Public Const WAIT_LONG_SECONDS As Double = 180

Public Function HasTimedOut(ByVal startTime As Double, ByVal timeoutSeconds As Double) As Boolean

    Dim elapsed As Double

    elapsed = Timer - startTime
    If elapsed < 0 Then elapsed = elapsed + 86400

    HasTimedOut = (elapsed >= timeoutSeconds)

End Function

Public Function ReportTimeout(ByVal context As String) As Boolean

    On Error Resume Next
    gCtx.timeout = True
    On Error GoTo 0

    If context <> "" Then
        MsgBox "TimeOut: " & context, vbCritical, "TimeOut"
    End If

    ReportTimeout = False

End Function

Public Function WaitForIEReady(ByVal ie As Object, ByVal timeoutSeconds As Double, ByVal context As String) As Boolean

    Dim startTime As Double

    startTime = Timer

    Do While ie.Busy Or ie.ReadyState <> 4
        DoEvents
        If HasTimedOut(startTime, timeoutSeconds) Then
            WaitForIEReady = ReportTimeout(context)
            Exit Function
        End If
    Loop

    WaitForIEReady = True

End Function

Public Function WaitForNonEmpty(ByRef value As Variant, ByVal timeoutSeconds As Double, ByVal context As String) As Boolean

    Dim startTime As Double

    startTime = Timer

    Do While Trim(CStr(value)) = ""
        DoEvents
        If HasTimedOut(startTime, timeoutSeconds) Then
            WaitForNonEmpty = ReportTimeout(context)
            Exit Function
        End If
    Loop

    WaitForNonEmpty = True

End Function

Public Function WaitForProcess(ByVal proc As Object, ByVal timeoutSeconds As Double, ByVal context As String) As Boolean

    Dim startTime As Double

    startTime = Timer

    Do While proc.Status = 0
        DoEvents
        If HasTimedOut(startTime, timeoutSeconds) Then
            WaitForProcess = ReportTimeout(context)
            Exit Function
        End If
    Loop

    WaitForProcess = True

End Function
