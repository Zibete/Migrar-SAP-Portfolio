Attribute VB_Name = "modTestAssertions"

Option Explicit

Public Sub AssertEqual(ByVal name As String, ByVal expected As Variant, ByVal actual As Variant)
    Dim message As String

    If expected <> actual Then
        message = "FAIL: " & name & " expected=" & expected & " actual=" & actual
        Debug.Print message
        TestLogLine message
        Err.Raise vbObjectError + 511, "AssertEqual", Mid$(message, 7)
    End If
End Sub

Public Sub AssertTrue(ByVal name As String, ByVal condition As Boolean)
    Dim message As String

    If Not condition Then
        message = "FAIL: " & name & " expected=True actual=False"
        Debug.Print message
        TestLogLine message
        Err.Raise vbObjectError + 512, "AssertTrue", Mid$(message, 7)
    End If
End Sub

Public Sub AssertContains(ByVal name As String, ByVal value As String, ByVal fragment As String)
    Dim message As String

    If InStr(1, value, fragment) = 0 Then
        message = "FAIL: " & name & " expected_contains=" & fragment & " actual=" & value
        Debug.Print message
        TestLogLine message
        Err.Raise vbObjectError + 513, "AssertContains", Mid$(message, 7)
    End If
End Sub

Public Sub AssertNotContains(ByVal name As String, ByVal value As String, ByVal fragment As String)
    Dim message As String

    If InStr(1, value, fragment) > 0 Then
        message = "FAIL: " & name & " expected_not_contains=" & fragment & " actual=" & value
        Debug.Print message
        TestLogLine message
        Err.Raise vbObjectError + 514, "AssertNotContains", Mid$(message, 7)
    End If
End Sub
