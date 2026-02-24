Attribute VB_Name = "modReferences"

Sub largoyletraRef()

    gCtx.largoReferencia = ""
    gCtx.letra = ""

    For Each celda In gCtx.rngRemitoRef.DataBodyRange
        If Hoja2.Cells(celda.Row, gCtx.rngReferencia.DataBodyRange.Column) = "" Then Exit For
        If Len(celda.Value) = 13 Or Len(celda.Value) = 14 Then
            sumaLongitudes = sumaLongitudes + Len(celda.Value)
            contador = contador + 1
        End If
        If Not IsEmpty(celda.Value) Then
            If InStr(1, celda.Value, "A") > 0 Then CountA = CountA + 1
            If InStr(1, celda.Value, "R") > 0 Then countR = countR + 1
            If InStr(1, celda.Value, "C") > 0 Then countC = countC + 1
        End If
    Next celda
    
    If sumaLongitudes <> "" Then gCtx.largoReferencia = Round(sumaLongitudes / contador, 0)
    If CountA > countR And CountA > countC Then gCtx.letra = "A"
    If countR > CountA And countR > countC Then gCtx.letra = "R"
    If countC > CountA And countC > countR Then gCtx.letra = "C"
    

End Sub
