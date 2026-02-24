Attribute VB_Name = "modSorting"

Sub ordenar(num_columna, order)

    gCtx.ControlarCambios = False
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Set columna = gCtx.tblDatos.ListColumns(num_columna - 2)
        
    gCtx.tblDatos.DataBodyRange.AutoFilter Field:=gCtx.tblDatos.ListColumns("Referencia").index, Criteria1:="<>"
        
    columna.DataBodyRange.Sort Key1:=columna.DataBodyRange.Cells(1), Order1:=order, Header:=xlYes
                    
    gCtx.ControlarCambios = True
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub

