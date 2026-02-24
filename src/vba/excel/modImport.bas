Attribute VB_Name = "modImport"

Private Const IMPORT_BACKUP_SHEET As String = "sheetImportBackup"

Sub Importar_0_SeleccionarDocumentos()
    Importar_1_SeleccionarDocumentos (GetRutaCarpeta())
End Sub
Sub Importar_0_Option()
    
    Importar_1_SeleccionarDocumentos (RUTA_IMPORTAR)

End Sub

Private Function BeginImportTransaction(ByVal ctx As AppContext, ByRef backupRows As Long, ByRef backupCols As Long) As Worksheet

    Dim backupSheet As Worksheet

    Set ctx = ResolveContext(ctx)

    backupRows = ctx.tblDatos.DataBodyRange.Rows.Count
    backupCols = ctx.tblDatos.DataBodyRange.Columns.Count

    SafeDeleteSheet IMPORT_BACKUP_SHEET

    Set backupSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    backupSheet.Name = IMPORT_BACKUP_SHEET
    backupSheet.Visible = xlSheetVeryHidden

    ctx.tblDatos.DataBodyRange.Copy
    backupSheet.Range("A1").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    Set BeginImportTransaction = backupSheet

End Function

Private Sub CommitImportTransaction(ByVal backupSheet As Worksheet)

    If Not backupSheet Is Nothing Then
        SafeDeleteSheet backupSheet.Name
    End If

End Sub

Private Sub RollbackImportTransaction(ByVal backupSheet As Worksheet, ByVal ctx As AppContext, ByVal backupRows As Long, ByVal backupCols As Long)

    If backupSheet Is Nothing Then Exit Sub

    Set ctx = ResolveContext(ctx)

    ctx.tblDatos.DataBodyRange.ClearContents
    backupSheet.Range("A1").Resize(backupRows, backupCols).Copy
    ctx.tblDatos.DataBodyRange.Cells(1, 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    SafeDeleteSheet backupSheet.Name

End Sub



Sub Importar_1_SeleccionarDocumentos(Ruta, Optional ctx As AppContext)
    
    Set ctx = ResolveContext(ctx)
    inicio = Timer

    countImpago = ContarImpagos(ctx)
    
    If Not ConfirmarImpagos(countImpago) Then Exit Sub
    
inicioProced:

    PrepararImportacion ctx
    ConfigurarOrigenDatos ctx

    If Ruta = RUTA_IMPORTAR Then
    
        formImportar.Show
        
    Else
        
        Set ArchivosSeleccionados = Application.FileDialog(msoFileDialogOpen)
        ConfigurarDialogoArchivos ArchivosSeleccionados, Ruta

        If ArchivosSeleccionados.Show = -1 Then
            ProcesarArchivosSeleccionados ArchivosSeleccionados, Ruta, ctx
        End If
    
    End If
    
    FinalizarImportacion inicio, ctx
    
End Sub

Private Function ContarImpagos(Optional ctx As AppContext) As Long

    Set ctx = ResolveContext(ctx)
    countImpago = 0
    
    For Each fila In ctx.tblDatos.ListRows
        If Hoja2.Cells(fila.Range.Row, ctx.rngEstado.Range.Column) = ESTADO_CONTABILIZADO Then
            If Hoja2.Cells(fila.Range.Row, ctx.rngPagado.Range.Column) = "NO" And Hoja2.Cells(fila.Range.Row, ctx.rngCompensacion.Range.Column) <> "" Then
                countImpago = countImpago + 1
            End If
        End If
    Next fila

    ContarImpagos = countImpago

End Function

Private Function ConfirmarImpagos(countImpago) As Boolean

    If countImpago = 0 Then
        ConfirmarImpagos = True
        Exit Function
    End If

    ConfirmarImpagos = (MsgBox("Hay " & countImpago & " registro(s) contabilizado(s) que no fueron pagados en RetailWeb." _
    & vbCrLf & vbCrLf & " Ã‚Â¿Desea continuar?" & rutaArchivo, vbYesNo + vbExclamation, "Confirmar acciÃƒÂ³n") _
    <> vbNo)

End Function

Private Sub PrepararImportacion(Optional ctx As AppContext)

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    UnprotectHoja2Safe
    
    asignaciones
    Set ctx = ResolveContext(ctx)
End Sub

Private Sub ConfigurarOrigenDatos(Optional ctx As AppContext)

    Set ctx = ResolveContext(ctx)
    If GetOrigenDatos() = ORIGEN_DATOS_SB Then
        If GetMantenerDatos() = FLAG_SI Then
            Set libroMigrar = ThisWorkbook
            Set sheetRetailWeb = libroMigrar.Sheets("sheetRetailWeb")
            MENSAJE = MsgBox("La herramienta utilizarÃƒÂ¡ datos almacenados en lugar de descargarlos nuevamente. " & vbCrLf & vbCrLf & "ÃƒÅ¡ltima actualizaciÃƒÂ³n: " & sheetRetailWeb.Range("A1"), vbOKOnly, "RETAILWEB")
        Else
            ctx.reporteSB = True
        End If
    End If

End Sub

Private Sub ConfigurarDialogoArchivos(ArchivosSeleccionados, Ruta)

    With ArchivosSeleccionados
        .AllowMultiSelect = True
        .Title = "Seleccionar archivos PDF para procesar"
        .InitialFileName = Ruta
        .Filters.Clear
        .Filters.Add "Archivos PDF", "*.pdf"
    End With

End Sub

Private Sub ProcesarArchivosSeleccionados(ArchivosSeleccionados, Ruta, Optional ctx As AppContext)

    Dim backupSheet As Worksheet
    Dim backupRows As Long
    Dim backupCols As Long

    Set ctx = ResolveContext(ctx)
    Set backupSheet = BeginImportTransaction(ctx, backupRows, backupCols)
    On Error GoTo ImportError
    capt1 = "Leyendo documentos PDF"

    ctx.ControlarCambios = False

    tareas2 = ArchivosSeleccionados.SelectedItems.Count
    tareasSB = 4
    tareasFin = 2
    tareas1 = tareas2 + tareasSB + tareas2 + tareasFin

    'Abrir
    With ProgressBar
        .Show vbModeless
        .Lbl1.Caption = capt1 & " (0%)"
        .Lbl2.Caption = "Preparando todo..." & " (0%)"
        .pb1.Max = tareas1
    End With

    ctx.tblDatos.AutoFilter.ShowAllData
    ctx.tblDatos.DataBodyRange.ClearContents

    Dim FSO As Scripting.FileSystemObject
    Set FSO = New Scripting.FileSystemObject

    SetConfigValue "Vend", ""
    SetConfigValue "nombreProveedor", ""

    y = ctx.tblDatos.Range.Row + 1

    pb2Value = 0

    For Each rutaArchivo In ArchivosSeleccionados.SelectedItems

        ctx.NombreArchivo = FSO.GetFileName(rutaArchivo)
        Hoja2.Cells(y, ctx.rngNombreArchivo.Range.Column).Value = ctx.NombreArchivo

        ctx.rutaCarpeta = FSO.GetParentFolderName(rutaArchivo) & "\"
        SetRutaCarpeta ctx.rutaCarpeta
        
        pb2Value = pb2Value + 1
        
        ' 1-Leyendo PDF... (count)
        With ProgressBar
            .pb2.Max = tareas2
            .pb1.Value = y - 8
            .pb2.Value = pb2Value
            .Lbl1.Caption = capt1 & " (" & Format(.pb1.Value / tareas1, "0%") & ")"
            .Lbl2.Caption = "Leyendo PDF: " & .pb2.Value & " de " & tareas2 & " (" & Format(.pb2.Value / tareas2, "0%") & ")"
        End With

        Call Importar_2_ProcesarPDF(y, rutaArchivo, ctx)

        If Hoja2.Cells(y, ctx.rngReferencia.Range.Column) <> "" Then
            y = y + 1
        Else
            ctx.tblDatos.DataBodyRange.Rows(y - ctx.tblDatos.Range.Row).ClearContents
        End If

    Next rutaArchivo

    ' 2-Buscando datos de RetailWeb (3)
    capt2 = "Buscando datos de RetailWeb..."
    With ProgressBar
        .pb2.Max = tareasSB
        .pb1.Value = ProgressBar.pb1.Value + 1
        .pb2.Value = 0
        .Lbl1.Caption = capt1 & " (" & Format(ProgressBar.pb1.Value / tareas1, "0%") & ")"
        .Lbl2.Caption = capt2
    End With

    For Each Conn In ThisWorkbook.Connections
        Conn.Delete
    Next Conn

    SafeDeleteQuery "Tabla_PDF"
    SafeDeleteSheet "DatosPDF"

    'Cubo
    If GetOrigenDatos() = ORIGEN_DATOS_CUBO Then Call Cubo(capt1, capt2)
    'CuboSB
    If GetOrigenDatos() = ORIGEN_DATOS_SB Then Call CuboSB(capt1, capt2)

    Call Importar_3_SB_to_MIGRAR(y, Ruta, ctx)
    Call Importar_5_Finalizar(y, Ruta, ctx)

    CommitImportTransaction backupSheet
    Exit Sub

ImportError:
    RollbackImportTransaction backupSheet, ctx, backupRows, backupCols
    ctx.ControlarCambios = True
    MENSAJE = MsgBox("Error durante la importaciÃ³n: " & Err.Description, vbCritical, "ImportaciÃ³n")

End Sub

Private Sub FinalizarImportacion(inicio, Optional ctx As AppContext)

    Set ctx = ResolveContext(ctx)
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1

    ActiveSheet.Shapes("flecha").Visible = False

    ProtectHoja2ForUi

    Unload ProgressBar

    fin = Timer
    duracion = fin - inicio

    minutos = Int(duracion \ 60)
    segundos = duracion Mod 60

    Debug.Print "Tiempo de ejecuciÃƒÂ³n: " & minutos & " minutos " & Format(segundos, "0.00") & " segundos"

    Application.Windows(ThisWorkbook.Name).DisplayHeadings = False

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.CutCopyMode = False

    Application.Windows(ThisWorkbook.Name).DisplayWorkbookTabs = False

End Sub

Sub Importar_2_MIGRAR_to_DB(y, REF, Optional ctx As AppContext)

    Set ctx = ResolveContext(ctx)
    Set DB_Encontrado = ctx.rngReferencia_DB.DataBodyRange.Find(What:=REF, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)

    If Not DB_Encontrado Is Nothing Then
        Set fila = ctx.tblDataBase.ListRows(DB_Encontrado.Row - ctx.tblDataBase.DataBodyRange.Row + 1)
    Else
        Set fila = ctx.tblDataBase.ListRows.Add
    End If
    
''''''''''''''' Provisorio
    fila.Range(1, ctx.rngVendor_DB.Range.Column) = GetVendorFilter()
''''''''''''''' Provisorio

    If Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column) <> "" Then fila.Range(1, ctx.rngTipoDoc_DB.Range.Column) = Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column)
    If Hoja2.Cells(y, ctx.rngRetailWeb_SB.Range.Column) <> "" Then fila.Range(1, ctx.rngRetailWeb_DB.Range.Column) = Hoja2.Cells(y, ctx.rngRetailWeb_SB.Range.Column)
    fila.Range(1, ctx.rngReferencia_DB.Range.Column) = REF
    If Hoja2.Cells(y, ctx.rngSite.Range.Column) <> "" Then fila.Range(1, ctx.rngSite_DB.Range.Column) = Hoja2.Cells(y, ctx.rngSite.Range.Column)
    If Hoja2.Cells(y, ctx.rngFechaDeFactura.Range.Column) <> "" Then fila.Range(1, ctx.rngFecha_DB.Range.Column) = Hoja2.Cells(y, ctx.rngFechaDeFactura.Range.Column)
    If Hoja2.Cells(y, ctx.rngTotalBrutoFactura.Range.Column) <> "" Then fila.Range(1, ctx.rngTotal_DB.Range.Column) = Hoja2.Cells(y, ctx.rngTotalBrutoFactura.Range.Column)
    If Hoja2.Cells(y, ctx.rngSubtotalFactura.Range.Column) <> "" Then fila.Range(1, ctx.rngSubtotal_DB.Range.Column) = Hoja2.Cells(y, ctx.rngSubtotalFactura.Range.Column)
    If Hoja2.Cells(y, ctx.rngII.Range.Column) <> "" Then fila.Range(1, ctx.rngII_DB.Range.Column) = Hoja2.Cells(y, ctx.rngII.Range.Column)
    If Hoja2.Cells(y, ctx.rngIVA.Range.Column) <> "" Then fila.Range(1, ctx.rngIVA_DB.Range.Column) = Hoja2.Cells(y, ctx.rngIVA.Range.Column)
            
    For Each encabezado In ctx.tblDatos.HeaderRowRange
        If (Left(encabezado.Value, 4) = "IIBB" Or Left(encabezado.Value, 4) = "Perc") Then
            If Hoja2.Cells(y, encabezado.Column) <> "" Then
                COD_PERC = Right(encabezado, 4)
                VAL_PERC = Hoja2.Cells(y, encabezado.Column)
                If fila.Range(1, ctx.rngPerc1_DB.Range.Column) = "" Then
                    fila.Range(1, ctx.rngPerc1_DB.Range.Column) = VAL_PERC & COD_PERC
                ElseIf fila.Range(1, ctx.rngPerc2_DB.Range.Column) = "" Then
                    fila.Range(1, ctx.rngPerc2_DB.Range.Column) = VAL_PERC & COD_PERC
                ElseIf fila.Range(1, ctx.rngPerc3_DB.Range.Column) = "" Then
                    fila.Range(1, ctx.rngPerc3_DB.Range.Column) = VAL_PERC & COD_PERC
                ElseIf fila.Range(1, ctx.rngPerc4_DB.Range.Column) = "" Then
                    fila.Range(1, ctx.rngPerc4_DB.Range.Column) = VAL_PERC & COD_PERC
                End If
            End If
            If fila.Range(1, ctx.rngPerc4_DB.Range.Column) <> "" Then Exit For
        End If
    Next encabezado
    
    If Hoja2.Cells(y, ctx.rngCAE.Range.Column) <> "" Then fila.Range(1, ctx.rngCAE_DB.Range.Column) = Hoja2.Cells(y, ctx.rngCAE.Range.Column)
    If Hoja2.Cells(y, ctx.rngVTOCAE.Range.Column) <> "" Then fila.Range(1, ctx.rngVTOCAE_DB.Range.Column) = Hoja2.Cells(y, ctx.rngVTOCAE.Range.Column)
    If Hoja2.Cells(y, ctx.rngFechaBase.Range.Column) <> "" Then fila.Range(1, ctx.rngFechaBase_DB.Range.Column) = Hoja2.Cells(y, ctx.rngFechaBase.Range.Column)
    If Hoja2.Cells(y, ctx.rngEstadoDelPago.Range.Column) <> "" Then fila.Range(1, ctx.rngEstado_DB.Range.Column) = Hoja2.Cells(y, ctx.rngEstadoDelPago.Range.Column)
    If Hoja2.Cells(y, ctx.rngComentarios_User.Range.Column) <> "" Then fila.Range(1, ctx.rngComentarios_DB.Range.Column) = Hoja2.Cells(y, ctx.rngComentarios_User.Range.Column)
    If Hoja2.Cells(y, ctx.rngNombreArchivo.Range.Column) <> "" Then fila.Range(1, ctx.rngRefPDF_DB.Range.Column) = Hoja2.Cells(y, ctx.rngNombreArchivo.Range.Column)

End Sub

Sub Importar_3_SB_to_MIGRAR(y, Ruta, Optional ctx As AppContext)

    '''''''RW to MIGRAR
    Set ctx = ResolveContext(ctx)
    Set libroMigrar = ThisWorkbook
    On Error Resume Next
    Set sheetRetailWeb = libroMigrar.Sheets("sheetRetailWeb")
    Set tblCuboSB = sheetRetailWeb.ListObjects("tblCuboSB")
    On Error GoTo 0

    If Not sheetRetailWeb Is Nothing Then

        Set rngVendor_tblCuboSB = tblCuboSB.ListColumns("Proveedor")
        Set rngSite_tblCuboSB = tblCuboSB.ListColumns("Negocio")
        Set rngValorizado_SB_tblCuboSB = tblCuboSB.ListColumns("Costo Valorizado")
        Set rngComentarios_tblCuboSB = tblCuboSB.ListColumns("Observaciones del Pago")
        Set rngRetailWeb_SB_tblCuboSB = tblCuboSB.ListColumns("RetailWeb #")
        Set rngEstadoDelPago_tblCuboSB = tblCuboSB.ListColumns("Estado del Pago")
        Set rngTieneScan_SB_tblCuboSB = tblCuboSB.ListColumns("Tiene Factura Scaneada")
        Set rngReferencia_tblCuboSB = tblCuboSB.ListColumns("Referencia Ext.")
        Set rngFechaNegocio_tblCuboSB = tblCuboSB.ListColumns("Fecha de Negocio")
        Set rngTotalBrutoFactura_tblCuboSB = tblCuboSB.ListColumns("Total Bruto Factura")
        Set rngFechaDeFactura_tblCuboSB = tblCuboSB.ListColumns("Fecha de Factura")
        Set rngMontoNetoFactura_tblCuboSB = tblCuboSB.ListColumns("Monto Neto Factura")
        Set rngPagado_tblCuboSB = tblCuboSB.ListColumns("Fecha de Pago")
        Set rngAnulado_tblCuboSB = tblCuboSB.ListColumns("Anula/Anulado")
        
        If Ruta <> RUTA_IMPORTAR Then
            ReDim ctx.vendors(0)
            ctx.vendors(0) = GetVendorFilter()
        End If
   
        If IsEmpty(ctx.vendors) Then
            ctx.vendors(0) = GetVendorFilter()
        End If

        If ctx.vendors(0) <> "" Then

            For Each vndr In ctx.vendors
            
                If y > ctx.ultimaFila Then Exit For
                
                For Each dato In rngVendor_tblCuboSB.DataBodyRange
                
                    If y > ctx.ultimaFila Then Exit For
                
                    pagado = sheetRetailWeb.Cells(dato.Row, rngPagado_tblCuboSB.Range.Column)
                    anulado = sheetRetailWeb.Cells(dato.Row, rngAnulado_tblCuboSB.Range.Column)

                    If dato = vndr * 1 And pagado = "" And anulado = "No" Then
                        
                        retailWeb# = sheetRetailWeb.Cells(dato.Row, rngRetailWeb_SB_tblCuboSB.Range.Column)
                        
                        Dim celda As Range
                        Dim duplicado As Boolean
                        duplicado = False
                        
                        For Each celda In ctx.rngRetailWeb_SB.DataBodyRange
                            If Trim(CStr(celda.Value)) = Trim(CStr(retailWeb#)) Then
                                duplicado = True
                                Exit For
                            End If
                        Next celda

                        If Not duplicado Then
                        'If UBound(vendors) > LBound(vendors) Then
                            Hoja2.Cells(y, ctx.rngVendorProveedor_SB.Range.Column) = sheetRetailWeb.Cells(dato.Row, tblCuboSB.ListColumns("Proveedor").Range.Column)
                            Hoja2.Cells(y, ctx.rngNombreProveedor_SB.Range.Column) = sheetRetailWeb.Cells(dato.Row, tblCuboSB.ListColumns("Proveedor3").Range.Column)
                        'End If
                        
                        Hoja2.Cells(y, ctx.rngPagado.Range.Column) = "NO"
        
                        Hoja2.Cells(y, ctx.rngRetailWeb_SB.Range.Column) = retailWeb#
                        
                        Hoja2.Cells(y, ctx.rngSite_SB.Range.Column) = sheetRetailWeb.Cells(dato.Row, rngSite_tblCuboSB.Range.Column)
                        Hoja2.Cells(y, ctx.rngSite.Range.Column) = sheetRetailWeb.Cells(dato.Row, rngSite_tblCuboSB.Range.Column)
    
                        Call asignarCORS(y, Hoja2.Cells(y, ctx.rngSite_SB.Range.Column))
        
                        Hoja2.Cells(y, ctx.rngReferencia.Range.Column) = CStr(Hoja2.Cells(y, ctx.rngRetailWeb_SB.Range.Column))
                        Hoja2.Cells(y, ctx.rngRemitoRef.Range.Column) = UCase(sheetRetailWeb.Cells(dato.Row, rngReferencia_tblCuboSB.Range.Column))
                                            
                        Hoja2.Cells(y, ctx.rngFechaDoc_SB.Range.Column) = sheetRetailWeb.Cells(dato.Row, rngFechaDeFactura_tblCuboSB.Range.Column)
                        
                        Hoja2.Cells(y, ctx.rngTotalBruto_SB.Range.Column) = Replace(sheetRetailWeb.Cells(dato.Row, rngTotalBrutoFactura_tblCuboSB.Range.Column), "-", "") * 1
                        
                        Hoja2.Cells(y, ctx.rngValorizado_SB.Range.Column) = Replace(sheetRetailWeb.Cells(dato.Row, rngValorizado_SB_tblCuboSB.Range.Column), "-", "") * 1
                        
                        Hoja2.Cells(y, ctx.rngSubtotal_SB.Range.Column) = Replace(sheetRetailWeb.Cells(dato.Row, rngMontoNetoFactura_tblCuboSB.Range.Column), "-", "") * 1
                        
                        fechaNeg_SB = Format(sheetRetailWeb.Cells(dato.Row, rngFechaNegocio_tblCuboSB.Range.Column), "dd/mm/yyyy")
                        If fechaNeg_SB <> "" And fechaNeg_SB <> 0 And IsDate(fechaNeg_SB) Then
                            fechaNeg_SB = CDate(fechaNeg_SB)
                            Hoja2.Cells(y, ctx.rngFechaNeg_SB.Range.Column) = fechaNeg_SB
                        End If
                
                        Hoja2.Cells(y, ctx.rngTieneScan_SB.Range.Column) = UCase(sheetRetailWeb.Cells(dato.Row, rngTieneScan_SB_tblCuboSB.Range.Column))
                        Hoja2.Cells(y, ctx.rngEstadoDelPago_SB.Range.Column) = sheetRetailWeb.Cells(dato.Row, rngEstadoDelPago_tblCuboSB.Range.Column)
                        Hoja2.Cells(y, ctx.rngObservacionesDelPago_SB.Range.Column) = sheetRetailWeb.Cells(dato.Row, rngComentarios_tblCuboSB.Range.Column)
                        
                        ESFCE = False
                        Set rngProveedor = ctx.rngVendor_Prov.DataBodyRange.Find(What:=Hoja2.Cells(y, ctx.rngVendorProveedor_SB.Range.Column), LookAt:=xlWhole)
                        ESFCE = CoreIsFCE( _
                            Hoja3.Cells(rngProveedor.Row, ctx.rngEsPyme_Prov.Range.Column) = "SI", _
                            Hoja2.Cells(y, ctx.rngTotalBruto_SB.Range.Column), _
                            ctx.montoFCE _
                        )


                        Hoja2.Cells(y, ctx.rngRemitoRef.Range.Column).NumberFormat = "General"
                        tipoDocSB = CoreTipoDocFromRetailWeb( _
                            Hoja2.Cells(y, ctx.rngRetailWeb_SB.Range.Column), _
                            CStr(Hoja2.Cells(y, ctx.rngRemitoRef.Range.Column)), _
                            ESFCE _
                        )
                        If tipoDocSB <> "" Then Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column) = tipoDocSB
                        
                        y = y + 1

                        End If
                    End If
                Next dato
            Next vndr
        End If
    End If
    
    Importar_4_DB_to_MIGRAR
    
End Sub

Sub Importar_4_DB_to_MIGRAR(Optional ctx As AppContext)

    Set ctx = ResolveContext(ctx)
        If Not ctx.rngRetailWeb_DB.DataBodyRange Is Nothing Then
            For Each RetailWeb_DB In ctx.rngRetailWeb_DB.DataBodyRange
                For Each retailWeb_SB In ctx.rngRetailWeb_SB.DataBodyRange
                    If retailWeb_SB <> "" And RetailWeb_DB <> "" And retailWeb_SB = RetailWeb_DB Then
                        If Hoja2.Cells(retailWeb_SB.Row, ctx.rngNombreArchivo.Range.Column) = "" Then

                            Hoja2.Cells(retailWeb_SB.Row, ctx.rngReferencia.Range.Column) = ctx.sheetDataBase.Cells(RetailWeb_DB.Row, ctx.rngReferencia_DB.Range.Column)
                            Hoja2.Cells(retailWeb_SB.Row, ctx.rngFechaDeFactura.Range.Column) = ctx.sheetDataBase.Cells(RetailWeb_DB.Row, ctx.rngFecha_DB.Range.Column)
                            Hoja2.Cells(retailWeb_SB.Row, ctx.rngTotalBrutoFactura.Range.Column) = ctx.sheetDataBase.Cells(RetailWeb_DB.Row, ctx.rngTotal_DB.Range.Column)
                            Hoja2.Cells(retailWeb_SB.Row, ctx.rngSubtotalFactura.Range.Column) = ctx.sheetDataBase.Cells(RetailWeb_DB.Row, ctx.rngSubtotal_DB.Range.Column)
                            Hoja2.Cells(retailWeb_SB.Row, ctx.rngIVA.Range.Column) = ctx.sheetDataBase.Cells(RetailWeb_DB.Row, ctx.rngIVA_DB.Range.Column)
                            Hoja2.Cells(retailWeb_SB.Row, ctx.rngII.Range.Column) = ctx.sheetDataBase.Cells(RetailWeb_DB.Row, ctx.rngII_DB.Range.Column)
                            
                            'Percepciones:
                            txtPerc1 = ctx.sheetDataBase.Cells(RetailWeb_DB.Row, ctx.rngPerc1_DB.Range.Column)
                            txtPerc2 = ctx.sheetDataBase.Cells(RetailWeb_DB.Row, ctx.rngPerc2_DB.Range.Column)
                            txtPerc3 = ctx.sheetDataBase.Cells(RetailWeb_DB.Row, ctx.rngPerc3_DB.Range.Column)
    
                            If txtPerc1 <> "" Then PERC1 = Left(txtPerc1, Len(txtPerc1) - 4)
                            If txtPerc1 <> "" Then codPerc1 = Right(txtPerc1, 4)
                            
                            If txtPerc2 <> "" Then PERC2 = Left(txtPerc2, Len(txtPerc2) - 4)
                            If txtPerc2 <> "" Then codPerc2 = Right(txtPerc2, 4)
                                                        
                            If txtPerc3 <> "" Then PERC3 = Left(txtPerc3, Len(txtPerc3) - 4)
                            If txtPerc3 <> "" Then codPerc3 = Right(txtPerc3, 4)
                            
                            For Each encabezado In ctx.tblDatos.HeaderRowRange
                                If Right(encabezado, 4) = codPerc1 Then
                                    Hoja2.Cells(retailWeb_SB.Row, encabezado.Column) = PERC1 * 1
                                End If
                                If Right(encabezado, 4) = codPerc2 Then
                                    Hoja2.Cells(retailWeb_SB.Row, encabezado.Column) = PERC2 * 1
                                End If
                                If Right(encabezado, 4) = codPerc3 Then
                                    Hoja2.Cells(retailWeb_SB.Row, encabezado.Column) = PERC3 * 1
                                End If
                            Next encabezado

                            Hoja2.Cells(retailWeb_SB.Row, ctx.rngCAE.Range.Column) = ctx.sheetDataBase.Cells(RetailWeb_DB.Row, ctx.rngCAE_DB.Range.Column)
                            Hoja2.Cells(retailWeb_SB.Row, ctx.rngVTOCAE.Range.Column) = ctx.sheetDataBase.Cells(RetailWeb_DB.Row, ctx.rngVTOCAE_DB.Range.Column)
                            Hoja2.Cells(retailWeb_SB.Row, ctx.rngFechaBase.Range.Column) = ctx.sheetDataBase.Cells(RetailWeb_DB.Row, ctx.rngFechaBase_DB.Range.Column)
                            Hoja2.Cells(retailWeb_SB.Row, ctx.rngEstadoDelPago.Range.Column) = ctx.sheetDataBase.Cells(RetailWeb_DB.Row, ctx.rngEstado_DB.Range.Column)
                            Hoja2.Cells(retailWeb_SB.Row, ctx.rngComentarios_User.Range.Column) = ctx.sheetDataBase.Cells(RetailWeb_DB.Row, ctx.rngComentarios_DB.Range.Column)
                            ctx.rngComentarios_User.Range.Columns.AutoFit
                            Hoja2.Cells(retailWeb_SB.Row, ctx.rngNombreArchivo.Range.Column) = "Completado por usuario (No hay PDF)"
                        
                        End If
                    End If
                Next retailWeb_SB
            Next RetailWeb_DB
        End If
        
    

End Sub

Sub Importar_4_DB_to_MIGRAR_PDF(y, DB_Encontrado, Optional ctx As AppContext)

    Set ctx = ResolveContext(ctx)
    Set REF_Encontrada = ctx.rngReferencia.DataBodyRange.Find( _
    What:=ctx.sheetDataBase.Cells(DB_Encontrado.Row, ctx.rngReferencia_DB.Range.Column), LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    
    If REF_Encontrada Is Nothing Then

        site = ctx.sheetDataBase.Cells(DB_Encontrado.Row, ctx.rngSite_DB.Range.Column)
    
        If site <> "" Then Call asignarCORS(y, site)
                
        Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column) = ctx.sheetDataBase.Cells(DB_Encontrado.Row, ctx.rngTipoDoc_DB.Range.Column)
    ''''''''''''''' Provisorio
        SetConfigValue "Vend", ctx.sheetDataBase.Cells(DB_Encontrado.Row, ctx.rngVendor_DB.Range.Column)
        Set rngProveedor = ctx.rngVendor_Prov.DataBodyRange.Find(What:=GetVendorFilter(), LookAt:=xlWhole)
        SetConfigValue "nombreProveedor", Hoja3.Cells(rngProveedor.Row, ctx.rngNombre_Prov.Range.Column)

    ''''''''''''''' Provisorio
        Hoja2.Cells(y, ctx.rngReferencia.Range.Column) = ctx.sheetDataBase.Cells(DB_Encontrado.Row, ctx.rngReferencia_DB.Range.Column)
        Hoja2.Cells(y, ctx.rngRemitoRef.Range.Column) = ctx.sheetDataBase.Cells(DB_Encontrado.Row, ctx.rngReferencia_DB.Range.Column)
        Hoja2.Cells(y, ctx.rngFechaDeFactura.Range.Column) = ctx.sheetDataBase.Cells(DB_Encontrado.Row, ctx.rngFecha_DB.Range.Column)
        Hoja2.Cells(y, ctx.rngTotalBrutoFactura.Range.Column) = ctx.sheetDataBase.Cells(DB_Encontrado.Row, ctx.rngTotal_DB.Range.Column)
        Hoja2.Cells(y, ctx.rngSubtotalFactura.Range.Column) = ctx.sheetDataBase.Cells(DB_Encontrado.Row, ctx.rngSubtotal_DB.Range.Column)
        Hoja2.Cells(y, ctx.rngIVA.Range.Column) = ctx.sheetDataBase.Cells(DB_Encontrado.Row, ctx.rngIVA_DB.Range.Column)
        Hoja2.Cells(y, ctx.rngII.Range.Column) = ctx.sheetDataBase.Cells(DB_Encontrado.Row, ctx.rngII_DB.Range.Column)
                            
        'Percepciones:
        txtPerc1 = ctx.sheetDataBase.Cells(DB_Encontrado.Row, ctx.rngPerc1_DB.Range.Column)
        txtPerc2 = ctx.sheetDataBase.Cells(DB_Encontrado.Row, ctx.rngPerc2_DB.Range.Column)
        txtPerc3 = ctx.sheetDataBase.Cells(DB_Encontrado.Row, ctx.rngPerc3_DB.Range.Column)
        txtPerc4 = ctx.sheetDataBase.Cells(DB_Encontrado.Row, ctx.rngPerc4_DB.Range.Column)
    
        If txtPerc1 <> "" Then PERC1 = Left(txtPerc1, Len(txtPerc1) - 4)
        If txtPerc1 <> "" Then codPerc1 = Right(txtPerc1, 4)
        
        If txtPerc2 <> "" Then PERC2 = Left(txtPerc2, Len(txtPerc2) - 4)
        If txtPerc2 <> "" Then codPerc2 = Right(txtPerc2, 4)
                                    
        If txtPerc3 <> "" Then PERC3 = Left(txtPerc3, Len(txtPerc3) - 4)
        If txtPerc3 <> "" Then codPerc3 = Right(txtPerc3, 4)
        
        If txtPerc4 <> "" Then PERC4 = Left(txtPerc4, Len(txtPerc4) - 4)
        If txtPerc4 <> "" Then codPerc4 = Right(txtPerc4, 4)
                            
        For Each encabezado In ctx.tblDatos.HeaderRowRange
            If Right(encabezado, 4) = codPerc1 Then
                Hoja2.Cells(y, encabezado.Column) = PERC1 * 1
            End If
            If Right(encabezado, 4) = codPerc2 Then
                Hoja2.Cells(y, encabezado.Column) = PERC2 * 1
            End If
            If Right(encabezado, 4) = codPerc3 Then
                Hoja2.Cells(y, encabezado.Column) = PERC3 * 1
            End If
            If Right(encabezado, 4) = codPerc4 Then
                Hoja2.Cells(y, encabezado.Column) = PERC4 * 1
            End If
        Next encabezado
    
        Hoja2.Cells(y, ctx.rngCAE.Range.Column) = ctx.sheetDataBase.Cells(DB_Encontrado.Row, ctx.rngCAE_DB.Range.Column)
        Hoja2.Cells(y, ctx.rngVTOCAE.Range.Column) = ctx.sheetDataBase.Cells(DB_Encontrado.Row, ctx.rngVTOCAE_DB.Range.Column)
        Hoja2.Cells(y, ctx.rngFechaBase.Range.Column) = ctx.sheetDataBase.Cells(DB_Encontrado.Row, ctx.rngFechaBase_DB.Range.Column)
        Hoja2.Cells(y, ctx.rngEstadoDelPago.Range.Column) = ctx.sheetDataBase.Cells(DB_Encontrado.Row, ctx.rngEstado_DB.Range.Column)
        Hoja2.Cells(y, ctx.rngComentarios_User.Range.Column) = ctx.sheetDataBase.Cells(DB_Encontrado.Row, ctx.rngComentarios_DB.Range.Column)
        Hoja2.Cells(y, ctx.rngNombreArchivo.Range.Column) = ctx.sheetDataBase.Cells(DB_Encontrado.Row, ctx.rngRefPDF_DB.Range.Column)
        
        ctx.rngComentarios_User.Range.Columns.AutoFit
    
    End If

End Sub

Sub Importar_5_Finalizar(y, Ruta, Optional ctx As AppContext)

    'Finalizar
    Set ctx = ResolveContext(ctx)
    If Ruta <> RUTA_REFRESH Then
        'Formulas
        Call formulaDifSAP
        ctx.rngEstadoCambiado.DataBodyRange.Formula = Hoja2.Cells(1, ctx.rngEstadoCambiado.Range.Column).Formula
        ctx.rngFechaBaseCambiada.DataBodyRange.Formula = Hoja2.Cells(1, ctx.rngEstadoCambiado.Range.Column).Formula
        ctx.rngDifCostos.DataBodyRange.Formula = Hoja2.Cells(1, ctx.rngDifCostos.Range.Column).Formula
        
        ctx.rngTexto.DataBodyRange.FormulaLocal = "=SI.ERROR(BUSCARV([@Sucursal];tblCors[[Sucursal]:[Texto]];2;0);"""")"
        ctx.rngCeBe.DataBodyRange.FormulaLocal = "=SI.ERROR(BUSCARV([@Sucursal];tblCors[[Sucursal]:[CeBe]];3;0);"""")"
        ctx.rngSupl.DataBodyRange.FormulaLocal = "=SI.ERROR(BUSCARV([@Sucursal];tblCors[[Sucursal]:[Supl.]];4;0);"""")"
        ctx.rngNombreSite.DataBodyRange.FormulaLocal = "=SI.ERROR(BUSCARV([@Sucursal];tblCors[[Sucursal]:[Nombre Sucursal]];5;0);"""")"
        ctx.rngZona.DataBodyRange.FormulaLocal = "=SI.ERROR(BUSCARV([@Sucursal];tblCors[[Sucursal]:[Zona]];6;0);"""")"
        ctx.rngAN.DataBodyRange.FormulaLocal = "=SI.ERROR(BUSCARV([@Sucursal];tblCors[[Sucursal]:[AN]];7;0);"""")"
        ctx.rngMails.DataBodyRange.FormulaLocal = "=SI.ERROR(BUSCARV([@Sucursal];tblCors[[Sucursal]:[Mails]];12;0);"""")"

        'Sucursal vacÃ­o
        For Each ssite In ctx.rngSite.DataBodyRange
            Hoja2.Cells(ssite.Row, ctx.rngReferencia.Range.Column).NumberFormat = "General"
            If Hoja2.Cells(ssite.Row, ctx.rngReferencia.Range.Column) = "" Then Exit For
            If ssite = "" Then
                site = Hoja2.Cells(ssite.Row, ctx.rngSite_SB.Range.Column)
                If site <> "" Then
                    For Each fila In ctx.tblCORS.ListRows
                        If fila.Range(ctx.tblCORS.ListColumns("Sucursal").Range.Column) = site Then
                            Hoja2.Cells(ssite.Row, ctx.rngTexto.Range.Column) = fila.Range(ctx.tblCORS.ListColumns("Texto").Range.Column).Value
                            Hoja2.Cells(ssite.Row, ctx.rngCeBe.Range.Column) = fila.Range(ctx.tblCORS.ListColumns("CeBe").Range.Column).Value
                            Hoja2.Cells(ssite.Row, ctx.rngNombreSite.Range.Column) = fila.Range(ctx.tblCORS.ListColumns("Nombre Sucursal").Range.Column).Value
                            Hoja2.Cells(ssite.Row, ctx.rngSupl.Range.Column) = fila.Range(ctx.tblCORS.ListColumns("Supl.").Range.Column).Value
                            Hoja2.Cells(ssite.Row, ctx.rngSite.Range.Column) = fila.Range(ctx.tblCORS.ListColumns("Sucursal").Range.Column).Value
                            Hoja2.Cells(ssite.Row, ctx.rngZona.Range.Column) = fila.Range(ctx.tblCORS.ListColumns("Zona").Range.Column).Value
                            Hoja2.Cells(ssite.Row, ctx.rngAN.Range.Column) = fila.Range(ctx.tblCORS.ListColumns("AN").Range.Column).Value
                            If Hoja2.Cells(ssite.Row, ctx.rngNombreArchivo.Range.Column) = "" Then
                                Hoja2.Cells(ssite.Row, ctx.rngMensajesSap.Range.Column) = "AVISO: Datos de RetailWeb (No hay PDF)"
                            Else
                                Hoja2.Cells(ssite.Row, ctx.rngMensajesSap.Range.Column) = "AVISO: Sucursal extraÃ­da de RetailWeb"
                            End If
                            Exit For
                        End If
                    Next fila
                End If
            End If
        Next ssite
    End If
    
    'Vendor y Proveedor
    If GetVendorFilter() <> "Varios" Then
        For Each vendor_SB In ctx.rngVendorProveedor_SB.DataBodyRange
            If Hoja2.Cells(vendor_SB.Row, ctx.rngReferencia.Range.Column) = "" Then Exit For
            If vendor_SB = "" Then
                Hoja2.Cells(vendor_SB.Row, ctx.rngVendorProveedor_SB.Range.Column) = GetVendorFilter()
                Hoja2.Cells(vendor_SB.Row, ctx.rngNombreProveedor_SB.Range.Column) = GetConfigValue("nombreProveedor", "")
            End If
        Next vendor_SB
    End If
        
    Call largoyletraRef
    
    'Comentarios
    For Each fila In ctx.tblDatos.ListRows
        If Hoja2.Cells(fila.Range.Row, ctx.rngReferencia.Range.Column) = "" Then Exit For
        
        observacionDelPago_SB = Hoja2.Cells(fila.Range.Row, ctx.rngObservacionesDelPago_SB.Range.Column)
        comentarios_User = Hoja2.Cells(fila.Range.Row, ctx.rngComentarios_User.Range.Column)
        
        comAutom = comentarioAutomatico(fila.Range.Row, observacionDelPago_SB, comentarios_User)
        comAutom = Replace(comAutom, "--", "-")
        Hoja2.Cells(fila.Range.Row, ctx.rngComentarios_SB.Range.Column) = comAutom
        
    Next fila
    
    
    
    'Comprobar estados
    For i = ctx.rngReferencia.DataBodyRange.Row To ctx.ultimaFila
        If Hoja2.Cells(i, ctx.rngReferencia.Range.Column) = "" Then Exit For
        If Hoja2.Cells(i, ctx.rngRetailWeb_SB.Range.Column) <> "" Then
            If Hoja2.Cells(i, ctx.rngReferencia.Range.Column) = "" Then Exit For
            If Hoja2.Cells(i, ctx.rngRetailWeb_SB.Range.Column) <> "Error CUBO" Then
                Call ComprobarEstados(i)
            End If
        End If
    Next i

    ctx.endoso = False
    VerificarDatos



    If Ruta <> RUTA_IMPORTAR And Ruta <> RUTA_REFRESH Then
     
        Set ArchivosSeleccionados = Application.FileDialog(msoFileDialogOpen)
        
        tareas2 = ArchivosSeleccionados.SelectedItems.Count
        tareasSB = 4
        tareasFin = 2
        tareas1 = tareas2 + tareasSB + tareas2 + tareasFin
        capt1 = "Leyendo documentos PDF"
    
        
        For Each archivoPDF In ctx.rngNombreArchivo.DataBodyRange
            If archivoPDF = "" Then Exit For
            If Left(archivoPDF, Len("Completado")) <> "Completado" Then
                ' 3.0-Renombrando documentos (count)
                ProgressBar.pb2.Max = tareas2
                ProgressBar.pb1.Value = ProgressBar.pb1.Value + 1
                ProgressBar.pb2.Value = archivoPDF.Row - 8
                ProgressBar.Lbl1.Caption = capt1 & " (" & Format(ProgressBar.pb1.Value / tareas1, "0%") & ")"
                ProgressBar.Lbl2.Caption = "Renombrando documentos: " & ProgressBar.pb2.Value & " de " & tareas2 & " (" & Format(ProgressBar.pb2.Value / tareas2, "0%") & ")"
                
                Renombrar (archivoPDF.Row)
                
            End If
        Next archivoPDF
        
        ' 4-Finalizando... (1)
        ProgressBar.pb2.Max = tareasFin
        ProgressBar.pb2.Value = 1
        ProgressBar.pb1.Value = ProgressBar.pb1.Value + 1
        ProgressBar.Lbl1.Caption = capt1 & " (" & Format(ProgressBar.pb1.Value / tareas1, "0%") & ")"
        ProgressBar.Lbl2.Caption = "Finalizando... (50%)"
        
    End If
        
    'Fecha dd/mm/yyyy
    If Ruta <> RUTA_REFRESH Then
        For Each dato In ctx.rngFechaDeFactura.DataBodyRange
            If ctx.rngReferencia.DataBodyRange.Cells(dato.Row - ctx.rngFechaDeFactura.DataBodyRange.Row + 1) = "" Then Exit For
            If dato <> "" Then
                fechaDoc = Replace(dato.Value, ".", "/")
                If IsDate(CDate(fechaDoc)) Then
                    fechaDoc = CDate(fechaDoc)
                    dato.Value = fechaDoc
                Else
                    dato.Value = "Error"
                End If
            End If
        Next dato
    End If
    
    'SOLO CMQ: Diccionario percepciones faltantes
    If GetVendorFilter() = "<REDACTED_ID_01>" Then
        For Each DIF In ctx.rngDifSap.DataBodyRange
            If Hoja2.Cells(DIF.Row, ctx.rngReferencia.Range.Column) = "" Then Exit For
            If DIF <> 0 And DIF <> "" Then
                extractedRef = Hoja2.Cells(DIF.Row, ctx.rngReferencia.Range.Column)
                If ctx.diccDocumentos.Exists(extractedRef) Then
                     datos = ctx.diccDocumentos(extractedRef)
                     For i = LBound(datos) To UBound(datos) Step 2
                         If datos(i) = "IIBBCABA" And datos(i + 1) <> "" Then Hoja2.Cells(DIF.Row, ctx.rngIIBBCABA.Range.Column) = datos(i + 1) * 1
                         If datos(i) = "IIBBCordoba" And datos(i + 1) <> "" Then Hoja2.Cells(DIF.Row, ctx.rngIIBBCordoba.Range.Column) = datos(i + 1) * 1
                         If datos(i) = "IIBBNeuquen" And datos(i + 1) <> "" Then Hoja2.Cells(DIF.Row, ctx.rngIIBBNeuquen.Range.Column) = datos(i + 1) * 1
                         If datos(i) = "MuniCord" And datos(i + 1) <> "" Then Hoja2.Cells(DIF.Row, ctx.rngMuniCord.Range.Column) = datos(i + 1) * 1
                         If datos(i) = "IIBBCatamarca" And datos(i + 1) <> "" Then Hoja2.Cells(DIF.Row, ctx.rngIIBBCatamarca.Range.Column) = datos(i + 1) * 1
                         If datos(i) = "IIBBEntreRios" And datos(i + 1) <> "" Then Hoja2.Cells(DIF.Row, ctx.rngIIBBEntreRios.Range.Column) = datos(i + 1) * 1
                         If datos(i) = "IIBBSalta" And datos(i + 1) <> "" Then Hoja2.Cells(DIF.Row, ctx.rngIIBBSalta.Range.Column) = datos(i + 1) * 1
                         If datos(i) = "IIBBCorrientes" And datos(i + 1) <> "" Then Hoja2.Cells(DIF.Row, ctx.rngIIBBCorrientes.Range.Column) = datos(i + 1) * 1
                         If datos(i) = "IIBBMendoza" And datos(i + 1) <> "" Then Hoja2.Cells(DIF.Row, ctx.rngIIBBMendoza.Range.Column) = datos(i + 1) * 1
                     Next i
                 End If
            End If
        Next DIF
        
    'SOLO CMQ: CAEA faltante CTO CAEA faltante
        For Each vtoCaeA In ctx.rngVTOCAE.DataBodyRange
            If Hoja2.Cells(vtoCaeA.Row, ctx.rngReferencia.Range.Column) = "" Then Exit For
            If vtoCaeA = "" Then
                For Each CAEA In ctx.rngCAE.DataBodyRange
                    If CAEA <> "" And CAEA = Hoja2.Cells(vtoCaeA.Row, ctx.rngCAE.Range.Column) Then
                        If Hoja2.Cells(CAEA.Row, ctx.rngVTOCAE.Range.Column) <> "" Then
                            Hoja2.Cells(vtoCaeA.Row, ctx.rngVTOCAE.Range.Column) = Hoja2.Cells(CAEA.Row, ctx.rngVTOCAE.Range.Column)
                            Hoja2.Cells(vtoCaeA.Row, ctx.rngMensajesSap.Range.Column) = "AVISO: El Vto. del CAEA fue extraÃ­do del doc. " & Hoja2.Cells(CAEA.Row, ctx.rngReferencia.Range.Column)
                        End If
                    End If
                Next CAEA
            End If
        Next vtoCaeA
    End If
        
    ctx.ControlarCambios = True

    'Diferencia con NC asociada
    If Ruta <> RUTA_REFRESH Then
        For Each tipoDoc In ctx.rngTipoDoc.DataBodyRange
            If Hoja2.Cells(tipoDoc.Row, ctx.rngReferencia.Range.Column) = "" Then Exit For
            If Left(tipoDoc, 2) = "FC" Then
                For Each rtoRef In ctx.rngRemitoRef.DataBodyRange
                    If rtoRef <> "" Then
                        If rtoRef = Hoja2.Cells(tipoDoc.Row, ctx.rngRemitoRef.Range.Column) Then
                            If Left(Hoja2.Cells(rtoRef.Row, ctx.rngTipoDoc.Range.Column), 2) = "NC" Then
                                'Verificar RemitoRef
                                If Hoja2.Cells(rtoRef.Row, ctx.rngRetailWeb_SB.Range.Column) = "" Then
                                    difRetailWeb = Hoja2.Cells(tipoDoc.Row, ctx.rngDifCostos.Range.Column)
                                    SubtotalNC = Hoja2.Cells(rtoRef.Row, ctx.rngSubtotalFactura.Range.Column)
                                    IINC = Hoja2.Cells(rtoRef.Row, ctx.rngII.Range.Column)
                                    Hoja2.Cells(tipoDoc.Row, ctx.rngDifConNC.Range.Column).Value = difRetailWeb - CDbl(SubtotalNC) - CDbl(IINC)
                                Else
                                    Hoja2.Cells(rtoRef.Row, ctx.rngRemitoRef.Range.Column) = Hoja2.Cells(rtoRef.Row, ctx.rngReferencia.Range.Column)
                                End If
                            End If
                        End If
                    End If
                Next rtoRef
            End If
        Next tipoDoc
    End If
    

    If Ruta <> RUTA_IMPORTAR And Ruta <> RUTA_REFRESH Then
        With ProgressBar
            .pb2.Value = 2
            .Lbl2.Caption = "Finalizando... (100%)"
            .pb1.Value = tareas1
            .Lbl1.Caption = capt1 & " (" & Format(ProgressBar.pb1.Value / tareas1, "0%") & ")"
        End With
    End If

    Call Importar_5_AjustesFinales
    
    If Ruta <> RUTA_REFRESH Then
        For Each PercIVA In ctx.rngPercIVA.DataBodyRange
            If Hoja2.Cells(PercIVA.Row, ctx.rngReferencia.Range.Column) = "" Then Exit For
            If PercIVA <> "" Then
                MENSAJE = MsgBox("AVISO: Algunos documentos tienen percepciÃ³n de IVA." & vbLf & "Ver columna ""Perc. IVA J1AP""", vbExclamation, "AVISO")
                Exit For
            End If
        Next PercIVA
    End If
    
    Hoja2.Select

End Sub
Public Sub Importar_5_AjustesFinales(Optional ctx As AppContext)

    'Ajustar columnas fijas visibles
    Set ctx = ResolveContext(ctx)
    rangosAjustar_FIJO = Array( _
        ctx.rngNombreArchivo, ctx.rngSite, ctx.rngNombreSite, ctx.rngFechaDeFactura, ctx.rngReferencia, ctx.rngTotalBrutoFactura, _
        ctx.rngCAE, ctx.rngVTOCAE, ctx.rngDifSap, ctx.rngCompensacion, ctx.rngMensajesSap, _
        ctx.rngTipoDoc, ctx.rngRemitoRef, _
        ctx.rngComentarios_SB)
        
    For Each rango In rangosAjustar_FIJO
        rango.Range.Columns.AutoFit
    Next rango
    
    'Ocultar->verificar->Mostrar
    rangosMostrar_VARIABLE = Array( _
        ctx.rngSubtotalFactura, ctx.rngSubtotalFactura105, _
        ctx.rngII, _
        ctx.rngIVA, ctx.rngIVA105, _
        ctx.rngIIBBBSAS, ctx.rngIIBBCABA, ctx.rngIIBBChubut, ctx.rngIIBBTucuman, ctx.rngIIBBSalta, ctx.rngIIBBNeuquen, _
        ctx.rngIIBBSantaFe, ctx.rngIIBBCatamarca, ctx.rngIIBBChaco, ctx.rngIIBBCordoba, ctx.rngIIBBCorrientes, ctx.rngIIBBEntreRios, _
        ctx.rngIIBBFormosa, ctx.rngIIBBJujuy, ctx.rngIIBBLaPampa, ctx.rngIIBBLaRioja, ctx.rngIIBBMendoza, ctx.rngIIBBMisiones, _
        ctx.rngIIBBRioNegro, ctx.rngIIBBSanJuan, ctx.rngIIBBSantiago, ctx.rngIIBBSanLuis, ctx.rngIIBBSantaCruz, ctx.rngIIBBTierraDelFuego, _
        ctx.rngPercIVA, ctx.rngMuniCord, ctx.rngFechaBase, _
        ctx.rngRetailWeb_SB, ctx.rngSite_SB, ctx.rngFechaDoc_SB, _
        ctx.rngDifCostos, ctx.rngDifConNC, ctx.rngTieneScan_SB)
    
    For Each rango In rangosMostrar_VARIABLE
        rango.Range.EntireColumn.Hidden = True 'Ocultar
        Call MostrarColumna(rango) 'Mostrar?
    Next rango
    
    'Ajustar
    rangosAjustar_VARIABLE = Array( _
        ctx.rngSubtotalFactura, ctx.rngSubtotalFactura105, _
        ctx.rngII, _
        ctx.rngIVA, ctx.rngIVA105, _
        ctx.rngIIBBBSAS, ctx.rngIIBBCABA, ctx.rngIIBBLaRioja, ctx.rngIIBBCatamarca, ctx.rngIIBBCorrientes, ctx.rngIIBBNeuquen, _
        ctx.rngIIBBMendoza, ctx.rngIIBBMisiones, ctx.rngIIBBSalta, ctx.rngIIBBEntreRios, ctx.rngMuniCord, ctx.rngPercIVA, _
        ctx.rngFechaBase)
    
    For Each rango In rangosAjustar_VARIABLE
        If rango.Range.EntireColumn.Hidden = False Then rango.Range.Columns.AutoFit
    Next rango
    
    ctx.rngVendorProveedor_SB.Range.EntireColumn.Hidden = True
    ctx.rngNombreProveedor_SB.Range.EntireColumn.Hidden = True
        
    If GetVendorFilter() = "Varios" Then ctx.rngVendorProveedor_SB.Range.EntireColumn.Hidden = False
    If GetVendorFilter() = "Varios" Then ctx.rngNombreProveedor_SB.Range.EntireColumn.Hidden = False


    'OCULTAR FILAS SIN REFERENCIA
    ctx.tblDatos.DataBodyRange.AutoFilter Field:=ctx.tblDatos.ListColumns("Referencia").index, Criteria1:="<>"
    
End Sub
Sub formulaDifSAP(Optional ctx As AppContext)

    Set ctx = ResolveContext(ctx)
    Const LF As String = vbLf
    Dim formulaLarga As String

    formulaLarga = "=SI.ERROR(" & _
                   "SI(tblDatos[@[Total" & LF & "Bruto]]<>"""";" & _
                   "REDONDEAR(" & _
                     "SUMA(tblDatos[@[Total" & LF & "Bruto]]*1;)-SUMA(" & _
                       "tblDatos[@[Subtotal" & LF & "con IVA 21,00%]]*1;" & _
                       "tblDatos[@[Subtotal" & LF & "con IVA 10,50%]]*1;" & _
                       "tblDatos[@[Impuestos" & LF & "Internos]]*1;" & _
                       "tblDatos[@[IVA" & LF & "21,00%]]*1;" & _
                       "tblDatos[@[IVA" & LF & "10,50%]]*1;" & _
                       "tblDatos[@[IIBB" & LF & "BS. AS." & LF & "J100]]*1;" & _
                       "tblDatos[@[IIBB" & LF & "CABA" & LF & "J101]]*1;" & _
                       "tblDatos[@[IIBB" & LF & "Chubut" & LF & "J102]]*1;" & _
                       "tblDatos[@[IIBB" & LF & "TucumÃ¡n" & LF & "J103]]*1;" & _
                       "tblDatos[@[IIBB" & LF & "Salta" & LF & "J104]]*1;" & _
                       "tblDatos[@[IIBB" & LF & "NeuquÃ©n" & LF & "J105]]*1;" & _
                       "tblDatos[@[IIBB" & LF & "Santa FÃ©" & LF & "J106]]*1;" & _
                       "tblDatos[@[IIBB" & LF & "Catamarca" & LF & "J107]]*1;"


    formulaLarga = formulaLarga & "tblDatos[@[IIBB" & LF & "Chaco" & LF & "J108]]*1;" & _
                                 "tblDatos[@[IIBB" & LF & "CÃ³rdoba" & LF & "J109]]*1;" & _
                                "tblDatos[@[IIBB" & LF & "Corrientes" & LF & "J110]]*1;" & _
                                "tblDatos[@[IIBB" & LF & "Entre RÃ­os" & LF & "J111]]*1;" & _
                                "tblDatos[@[IIBB" & LF & "Formosa" & LF & "J112]]*1;" & _
                                "tblDatos[@[IIBB" & LF & "Jujuy" & LF & "J113]]*1;" & _
                                "tblDatos[@[IIBB" & LF & "La Pampa" & LF & "J114]]*1;" & _
                                "tblDatos[@[IIBB" & LF & "La Rioja" & LF & "J115]]*1;" & _
                                "tblDatos[@[IIBB" & LF & "Mendoza" & LF & "J116]]*1;" & _
                                "tblDatos[@[IIBB" & LF & "Misiones" & LF & "J117]]*1;" & _
                                "tblDatos[@[IIBB" & LF & "Rio Negro" & LF & "J118]]*1;" & _
                                "tblDatos[@[IIBB" & LF & "San Juan" & LF & "J119]]*1;" & _
                                "tblDatos[@[IIBB" & LF & "Santiago del Estero" & LF & "J120]]*1;" & _
                                "tblDatos[@[IIBB" & LF & "San Luis" & LF & "J121]]*1;" & _
                                "tblDatos[@[IIBB" & LF & "Santa Cruz" & LF & "J122]]*1;" & _
                                "tblDatos[@[IIBB" & LF & "Tierra del Fuego" & LF & "J123]]*1;" & _
                                "tblDatos[@[Perc. Munic. CÃ³rdoba" & LF & "MCOR]]*1;" & _
                                "tblDatos[@[Perc. IVA" & LF & "J1AP]]*1" & _
                     ");2" & _
                   ");0" & _
                 ");0)"

    ctx.rngDifSap.DataBodyRange.FormulaLocal = formulaLarga

End Sub

Sub eliminarRegistrosDB(Optional ctx As AppContext)

    Set ctx = ResolveContext(ctx)
    ultimaFilaDB = ctx.tblDataBase.ListRows.Count + ctx.tblDataBase.DataBodyRange.Row - 1
    
    For i = ultimaFilaDB To ctx.tblDataBase.DataBodyRange.Row Step -1
        For j = ctx.ultimaFila To ctx.tblDatos.DataBodyRange.Row Step -1
            If Hoja2.Cells(j, ctx.rngReferencia.Range.Column) <> "" Then
                If Hoja2.Cells(j, ctx.rngRetailWeb_SB.Range.Column) = ctx.sheetDataBase.Cells(i, 1) Then
                    ctx.rngRetailWeb_DB.Rows(i).EntireRow.Delete
                End If
            End If
        Next j
    Next i

End Sub

















