Attribute VB_Name = "modState"

Public gCtx As AppContext

Public Sub ResetContext()
    Set gCtx = New AppContext
End Sub

Public Sub EnsureContext()
    If gCtx Is Nothing Then
        Set gCtx = New AppContext
    End If
End Sub

Public Function ResolveContext(Optional ctx As AppContext) As AppContext
    If ctx Is Nothing Then
        EnsureContext
        Set ResolveContext = gCtx
    Else
        Set ResolveContext = ctx
    End If
End Function
