Attribute VB_Name = "modSafeOps"

Public Sub SafeDeleteQuery(ByVal queryName As String)

    On Error GoTo CleanExit
    ActiveWorkbook.Queries.Item(queryName).Delete

CleanExit:
    On Error GoTo 0

End Sub

Public Sub SafeDeleteSheet(ByVal sheetName As String)

    Dim ws As Worksheet

    On Error GoTo CleanExit
    Set ws = ThisWorkbook.Sheets(sheetName)
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If

CleanExit:
    On Error GoTo 0

End Sub
