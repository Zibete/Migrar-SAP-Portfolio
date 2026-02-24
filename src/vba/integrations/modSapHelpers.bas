Attribute VB_Name = "modSapHelpers"

Public Function SapTryFindById(ByVal session As Object, ByVal idPath As String) As Object

    On Error GoTo CleanFail
    Set SapTryFindById = session.findById(idPath, False)
    Exit Function

CleanFail:
    Set SapTryFindById = Nothing
    On Error GoTo 0

End Function

Public Function SapTrySetText(ByVal session As Object, ByVal idPath As String, ByVal value As String) As Boolean

    On Error GoTo CleanFail
    session.findById(idPath).Text = value
    SapTrySetText = True
    Exit Function

CleanFail:
    SapTrySetText = False
    On Error GoTo 0

End Function
