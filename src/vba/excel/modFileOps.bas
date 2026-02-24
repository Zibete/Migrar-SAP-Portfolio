Attribute VB_Name = "modFileOps"

Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As LongPtr, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

Sub AbrirPDF()

    Dim rutaBase As String

    rutaBase = GetRutaCarpeta()

    For indice = 1 To Selection.Rows.Count
    
        i = Selection.Cells(indice, 1).Row
        
        If Not Hoja2.Rows(i).EntireRow.Hidden Then
        
            nombre = Hoja2.Cells(i, gCtx.rngNombreArchivo.Range.Column).Value
        
            ShellExecute 0, "open", rutaBase & nombre, vbNullString, vbNullString, 1
        
            REF = Hoja2.Cells(i, gCtx.rngReferencia.Range.Column).Value
        
            archivoReferencia = rutaBase & REF & "-Hoja 1.pdf"
        
            If Dir(archivoReferencia) <> "" Then

                ShellExecute 0, "open", archivoReferencia, vbNullString, vbNullString, 1

            End If
        
        End If

    Next indice
    
End Sub

Sub eliminarPDF()

    If gCtx.SELECTION_USER = 1 Then
        respuesta = MsgBox("Se eliminará el archivo: """ & Hoja2.Cells(Selection.Row, gCtx.rngNombreArchivo.Range.Column).Value & """." _
        & vbCrLf & vbCrLf & "Esta acción no puede deshacerse." & vbCrLf & vbCrLf & " ¿Desea continuar?" & rutaArchivo, vbYesNo + vbExclamation, "Confirmar eliminación")
    Else
        respuesta = MsgBox("Se eliminarán los " & gCtx.SELECTION_USER & " archivos seleccionados." _
        & vbCrLf & vbCrLf & "Esta acción no puede deshacerse." & vbCrLf & vbCrLf & " ¿Desea continuar?" & rutaArchivo, vbYesNo + vbExclamation, "Confirmar eliminación")
    End If

    If respuesta = vbYes Then
    
        rutaBase = GetRutaCarpeta()
    
        For indice = 1 To Selection.Rows.Count
    
            i = Selection.Cells(indice, 1).Row
        
            If Not Hoja2.Rows(i).EntireRow.Hidden Then
        
                nombre = Hoja2.Cells(i, gCtx.rngNombreArchivo.Range.Column).Value
        
                If Dir(rutaBase & nombre) <> "" Then
            
                    Kill rutaBase & nombre
                
                    If Dir(rutaBase & nombre) = "" Then
                        SetRowStatus i, ESTADO_ELIMINADO, ""
                        Hoja2.Cells(i, gCtx.rngRemitoRef.Range.Column).Value = ""
                    End If
                
                Else
                    MsgBox "El archivo: """ & nombre & """" & "no existe en la ruta """ & rutaBase & """", vbCritical
                End If
            
            End If

        Next indice
    
    End If
    
End Sub
Sub RenombrarSeleccion()

    For indice = 1 To Selection.Rows.Count
    
        i = Selection.Cells(indice, 1).Row
        
        If Not Hoja2.Rows(i).EntireRow.Hidden Then
        
            countRenombrar = countRenombrar + 1
        
            With ProgressBar
                .Show vbModeless
                .Lbl1.Caption = "Renombrando archivos. Progreso..." & " (" & Format(countRenombrar / gCtx.SELECTION_USER, "0%") & ")"
                .Lbl2.Caption = "Renombrando archivo " & countRenombrar & " de " & gCtx.SELECTION_USER & " (" & Format(countRenombrar / gCtx.SELECTION_USER, "0%") & ")"
                .pb1.Max = gCtx.SELECTION_USER
                .pb2.Max = gCtx.SELECTION_USER
                .pb1.Value = countRenombrar
                .pb2.Value = countRenombrar
            End With

            Call Renombrar(i)

        End If
        
    Next indice
    
    Unload ProgressBar
     
End Sub
Sub Renombrar(i)

    'On Error GoTo finProced

    gCtx.nuevoNombre = ""
    
    gCtx.NombreArchivo = Hoja2.Cells(i, gCtx.rngNombreArchivo.Range.Column).Value

    site = CStr(Hoja2.Cells(i, gCtx.rngSite.Range.Column))
    Referencia = Hoja2.Cells(i, gCtx.rngReferencia.Range.Column)
    If Len(Referencia) < 13 Then Referencia = ""
    fechaBase = Hoja2.Cells(i, gCtx.rngFechaBase.Range.Column)
    estadoPago = Hoja2.Cells(i, gCtx.rngEstadoDelPago.Range.Column)
    tipoDoc = Hoja2.Cells(i, gCtx.rngTipoDoc.Range.Column)

    hasRetailWeb = (Hoja2.Cells(i, gCtx.rngRetailWeb_SB.Range.Column) <> "")
    gCtx.nuevoNombre = CoreBuildNombreBase(site, tipoDoc, Referencia, fechaBase, hasRetailWeb, estadoPago)
        
    nombreArchivoNuevo = gCtx.nuevoNombre & ".pdf"
    
    
    If GetEliminarDuplicados() = FLAG_SI Then
    
        If InStr(1, nombreArchivoNuevo, "INS") > 0 Then
        
            Kill gCtx.rutaCarpeta & gCtx.NombreArchivo
            For j = gCtx.rngSubtotalFactura.Range.Column To gCtx.tblDatos.Range.Columns.Count + 2
                Hoja2.Cells(i, j) = ""
            Next j
            SetRowStatus i, ESTADO_ELIMINADO, "Eliminado: Son insumos"

        End If

    End If
            
    If nombreArchivoNuevo <> gCtx.NombreArchivo Then
    
        If Dir(gCtx.rutaCarpeta & nombreArchivoNuevo) <> "" Then 'Ya existe
            If GetEliminarDuplicados() = FLAG_SI Then 'And Len(Referencia) = largoReferencia And InStr(Referencia, letra) > 0 Then
                Kill gCtx.rutaCarpeta & gCtx.NombreArchivo
                For j = gCtx.rngSubtotalFactura.Range.Column To gCtx.tblDatos.Range.Columns.Count + 2
                    Hoja2.Cells(i, j) = ""
                Next j
                SetRowStatus i, ESTADO_ELIMINADO, "Eliminado: Ya existe un archivo con ese nombre"
            Else
                Do
                    countName = countName + 1
                    nombreArchivoNuevo = gCtx.nuevoNombre & "-" & countName & ".pdf"
                Loop While Dir(gCtx.rutaCarpeta & nombreArchivoNuevo) <> ""
            End If
        End If
        
        If Dir(gCtx.rutaCarpeta & gCtx.NombreArchivo) <> "" Then
            Name gCtx.rutaCarpeta & gCtx.NombreArchivo As gCtx.rutaCarpeta & nombreArchivoNuevo
            Hoja2.Cells(i, gCtx.rngNombreArchivo.Range.Column).Value = nombreArchivoNuevo
        End If
    
    End If
    
finProced:
    
    Hoja2.Columns(Hoja2.Range("tblDatos[[#Headers],[Nombre Archivo]]").Column).AutoFit
    
End Sub

Function ContarPaginasPDF(ByVal rutaArchivo As String) As Long

    Dim tempQuery As WorkbookQuery
    Dim tempSheet As Worksheet
    Dim tbl As ListObject
    Dim idColumn As ListColumn

    On Error GoTo ErrorHandler ' Habilitar manejo de errores
    pqFormula = "let Source = Pdf.Tables(File.Contents(""" & rutaArchivo & """), [Implementation=""1.3""]) in Source"

    SafeDeleteQuery "Temp_PDF_Info"
    SafeDeleteSheet "Temp_PDF_Info_Sheet"

    Set tempQuery = ThisWorkbook.Queries.Add(Name:="Temp_PDF_Info", Formula:=pqFormula)

    Set tempSheet = ThisWorkbook.Worksheets.Add
    tempSheet.Name = "Temp_PDF_Info_Sheet"

    With tempSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""Temp_PDF_Info"";Extended Properties=""""", _
        Destination:=tempSheet.Range("A1")).QueryTable
        .CommandText = "SELECT * FROM [Temp_PDF_Info]"
        .Refresh BackgroundQuery:=False
    End With

'    MsgBox rutaArchivo

    Set tbl = tempSheet.ListObjects(1)

    On Error Resume Next
    Set idColumn = tbl.ListColumns("Id")
    If idColumn Is Nothing Then
       If tbl.ListColumns.Count > 0 Then Set idColumn = tbl.ListColumns(1)
    End If
    On Error GoTo ErrorHandler

    maxPageNum = 0
    If Not idColumn Is Nothing Then
        For r = 1 To tbl.ListRows.Count
             cellValue = CStr(tbl.DataBodyRange.Cells(r, idColumn.index).Value)
            If UCase(Left(cellValue, 4)) = "PAGE" Then
                On Error Resume Next ' Ignorar si no se puede convertir a número
                pageNum = CLng(Mid(cellValue, 5))
                If Err.Number = 0 Then
                    If pageNum > maxPageNum Then
                        maxPageNum = pageNum
                    End If
                End If
                Err.Clear ' Limpiar error si ocurrió en CLng
                On Error GoTo ErrorHandler
            End If
        Next r
    End If

    SafeDeleteQuery "Temp_PDF_Info"
    SafeDeleteSheet "Temp_PDF_Info_Sheet"

    ContarPaginasPDF = maxPageNum
    Exit Function

ErrorHandler:
    SafeDeleteQuery "Temp_PDF_Info"
    SafeDeleteSheet "Temp_PDF_Info_Sheet"
    ContarPaginasPDF = -1 ' Indicar error (o podrías usar 0)
    MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "en ContarPaginasPDF", vbCritical
End Function




