Attribute VB_Name = "modImportPdf"

Sub Importar_2_ProcesarPDF(y, rutaArchivo, Optional ctx As AppContext)
   
    Dim totalPaginas As Long
    Dim tempQuery As WorkbookQuery
    Dim hojaPDF As Worksheet

    Set ctx = ResolveContext(ctx)

    totalPaginas = ContarPaginasPDF(rutaArchivo)

importarDatos:

    If totalPaginas <= 0 Then GoTo fin

    For PageIndex = 1 To totalPaginas

        Set hojaPDF = CargarHojaPDF(rutaArchivo, PageIndex, tempQuery)
        If hojaPDF Is Nothing Then GoTo siguientePagina

        ResolverVendorDesdePDF ctx, hojaPDF, y, vendorNuevo
        If EjecutarParser(ctx, hojaPDF, y, vendorNuevo) Then
            AplicarFechaBase ctx, y
        End If

siguientePagina:
    Next PageIndex

fin:

    LimpiarDatosPDF

End Sub

Private Function CargarHojaPDF(rutaArchivo, pageIndex, ByRef tempQuery As WorkbookQuery) As Worksheet

    Dim PDFTableQuery As String
    Dim NewSheet As Worksheet
    Dim qt As QueryTable
    Dim Page As String

    Page = "Page" & Format(pageIndex, "000")
    PDFTableQuery = CrearConsultaPDF(rutaArchivo, Page)

    LimpiarDatosPDF

    On Error GoTo CleanFail
    Set tempQuery = ThisWorkbook.Queries.Add(Name:="Tabla_PDF", Formula:=PDFTableQuery)

    Set NewSheet = ThisWorkbook.Worksheets.Add
    NewSheet.Name = "DatosPDF"

    Set qt = NewSheet.ListObjects.Add(SourceType:=0, Source:= _
             "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""Tabla_PDF"";Extended Properties=""""", _
             Destination:=NewSheet.Range("A1")).QueryTable

    With qt
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [Tabla_PDF]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .Refresh BackgroundQuery:=False
    End With

    Set CargarHojaPDF = NewSheet
    Exit Function

CleanFail:
    On Error Resume Next
    If Not tempQuery Is Nothing Then tempQuery.Delete
    If Not NewSheet Is Nothing Then NewSheet.Delete
    On Error GoTo 0

End Function

Private Function CrearConsultaPDF(rutaArchivo, Page) As String

    CrearConsultaPDF = "let" & vbCrLf & _
                      "    Origen = Pdf.Tables(File.Contents(""" & rutaArchivo & """), [Implementation=""1.3""])," & vbCrLf & _
                      "    Table001 = Origen{[Id=""" & Page & """]}[Data]" & vbCrLf & _
                      "in" & vbCrLf & _
                      "    Table001"

End Function

Private Sub ResolverVendorDesdePDF(ctx As AppContext, hojaPDF, y, ByRef vendorNuevo)

    Dim vendorEncontrado As String
    Dim vendorNombre As String
    Dim vendorCUIT As String

    Set ctx = ResolveContext(ctx)
    ctx.vendorActual = GetVendorFilter()

    vendorEncontrado = BuscarVendorPorCUIT(ctx, hojaPDF, vendorNombre, vendorCUIT)

    If ctx.vendorActual = "" Then
        If vendorEncontrado <> "" Then
            ctx.vendorActual = vendorEncontrado
            vendorNuevo = ctx.vendorActual
            SetConfigValue "Vend", ctx.vendorActual
            SetConfigValue "nombreProveedor", vendorNombre
            SetConfigValue "CUIT", vendorCUIT
        End If
    Else
        If vendorEncontrado <> "" Then
            vendorNuevo = vendorEncontrado
            If vendorNuevo <> ctx.vendorActual Then
                Hoja2.Cells(y, ctx.rngReferencia.Range.Column).Value = vendorNombre
                Hoja2.Cells(y, ctx.rngTexto.Range.Column).Value = vendorNombre
                SetConfigValue "CUIT", vendorCUIT
            End If
        Else
            If vendorNuevo = "" Then
                Hoja2.Cells(y, ctx.rngReferencia.Range.Column).Value = "CUIT desconocido"
                Hoja2.Cells(y, ctx.rngTexto.Range.Column).Value = "CUIT desconocido"
            End If
        End If
    End If

    AplicarVendorProvisorioA ctx, hojaPDF, y, vendorNuevo
    AplicarVendorProvisorioB ctx, hojaPDF, y, vendorNuevo

End Sub

Private Function BuscarVendorPorCUIT(ctx As AppContext, hojaPDF, ByRef vendorNombre, ByRef vendorCUIT) As String

    Set ctx = ResolveContext(ctx)

    For Each fila In ctx.TblProveedores.ListRows
        CUIT = Hoja3.Cells(fila.Range.Row, ctx.rngCUIT_Prov.Range.Column)
        If CUIT <> "" Then palabrabuscada = Mid(CUIT, 3, Len(CUIT) - 3)
        Set celdaencontrada = hojaPDF.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
        If Not celdaencontrada Is Nothing Then
            BuscarVendorPorCUIT = Hoja3.Cells(fila.Range.Row, ctx.rngVendor_Prov.Range.Column)
            vendorNombre = Hoja3.Cells(fila.Range.Row, ctx.rngNombre_Prov.Range.Column)
            vendorCUIT = Hoja3.Cells(fila.Range.Row, ctx.rngCUIT_Prov.Range.Column)
            Exit Function
        End If
    Next fila

End Function

Private Sub AplicarVendorProvisorioA(ctx As AppContext, hojaPDF, y, ByRef vendorNuevo)

    Set ctx = ResolveContext(ctx)

    If ctx.vendorActual = "" Or ctx.vendorActual = "<REDACTED_ID_03>" Then '<SUPPLIER_A>
        palabrabuscada = " DIAS FECHA DE FACTURA"
        Set celdaencontrada = hojaPDF.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
        If Not celdaencontrada Is Nothing Then
            If ctx.vendorActual = "" Then ctx.vendorActual = "<REDACTED_ID_03>"
            vendorNuevo = "<REDACTED_ID_03>"
            SetConfigValue "Vend", ctx.vendorActual
            SetConfigValue "nombreProveedor", "<SUPPLIER_A>"
        Else
            Hoja2.Cells(y, ctx.rngReferencia.Range.Column).Value = "CUIT desconocido"
            Hoja2.Cells(y, ctx.rngTexto.Range.Column).Value = "CUIT desconocido"
        End If
    End If

End Sub

Private Sub AplicarVendorProvisorioB(ctx As AppContext, hojaPDF, y, ByRef vendorNuevo)

    Set ctx = ResolveContext(ctx)

    If ctx.vendorActual = "" Or ctx.vendorActual = "<REDACTED_ID_14>" Then '<SUPPLIER_B>
        palabrabuscada = "VENDEDOR 001"
        Set celdaencontrada = hojaPDF.UsedRange.Find(What:=palabrabuscada, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
        If Not celdaencontrada Is Nothing Then
            If ctx.vendorActual = "" Then ctx.vendorActual = "<REDACTED_ID_14>"
            vendorNuevo = "<REDACTED_ID_14>"
            SetConfigValue "Vend", ctx.vendorActual
            SetConfigValue "nombreProveedor", "<SUPPLIER_B>"
        Else
            Hoja2.Cells(y, ctx.rngReferencia.Range.Column).Value = "CUIT desconocido"
            Hoja2.Cells(y, ctx.rngTexto.Range.Column).Value = "CUIT desconocido"
        End If
    End If

End Sub

Private Function EjecutarParser(ctx As AppContext, hojaPDF, y, vendorNuevo) As Boolean

    Dim parserName As String

    Set ctx = ResolveContext(ctx)
    EjecutarParser = True

    If ctx.vendorActual = vendorNuevo Then
        If EsVendorMultipaginaEspecial(ctx.vendorActual) Then
            If Not hojaPDF.UsedRange.Find(What:="001 DE 002", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False) Is Nothing Then
                EjecutarParser = False
                Exit Function
            Else
                Call ParseVendor07(hojaPDF, y, ctx)
            End If
        Else
            parserName = ObtenerParser(ctx.vendorActual)
            If parserName <> "" Then Application.Run parserName, hojaPDF, y, ctx
        End If
    End If

End Function

Private Sub AplicarFechaBase(ctx As AppContext, y)

    Set ctx = ResolveContext(ctx)

    If Hoja2.Cells(y, ctx.rngTipoDoc.Range.Column) = "FC-REM" Then
        If InStr(1, ctx.NombreArchivo, "Fecha base", vbTextCompare) > 0 Then
            posicionInicio = InStr(1, ctx.NombreArchivo, "Fecha base", vbTextCompare) + Len("Fecha base")
            caracteresExtraidos = Mid(ctx.NombreArchivo, posicionInicio, 11)
            Hoja2.Cells(y, ctx.rngFechaBase.Range.Column).Value = Replace(caracteresExtraidos, " ", "")
        End If
    End If

End Sub

Private Sub LimpiarDatosPDF()

    SafeDeleteQuery "Tabla_PDF"
    SafeDeleteSheet "DatosPDF"

End Sub

Private Function ObtenerParser(vendorId) As String

    Static registry As Object

    If registry Is Nothing Then
        Set registry = CreateObject("Scripting.Dictionary")
        registry.CompareMode = vbTextCompare
        registry.Add "<REDACTED_ID_04>", "ParseVendor09"
        registry.Add "<REDACTED_ID_05>", "ParseVendor20"
        registry.Add "<REDACTED_ID_03>", "ParseVendor18"
        registry.Add "<REDACTED_ID_14>", "ParseVendor03"
        registry.Add "<REDACTED_ID_09>", "ParseVendor01"
        registry.Add "<REDACTED_ID_08>", "ParseVendor21"
        registry.Add "<REDACTED_ID_10>", "ParseVendor17"
        registry.Add "<REDACTED_ID_20>", "ParseVendor19"
        registry.Add "<REDACTED_ID_11>", "ParseVendor04"
        registry.Add "<REDACTED_ID_12>", "ParseVendor14"
        registry.Add "<REDACTED_ID_17>", "ParseVendor13"
        registry.Add "<REDACTED_ID_21>", "ParseVendor05"
        registry.Add "<REDACTED_ID_01>", "ParseVendor06"
        registry.Add "<REDACTED_ID_22>", "ParseVendor02"
        registry.Add "<REDACTED_ID_15>", "ParseVendor11"
        registry.Add "<REDACTED_ID_13>", "ParseVendor12"
        registry.Add "<REDACTED_ID_19>", "ParseVendor16"
        registry.Add "<REDACTED_ID_18>", "ParseVendor08"
        registry.Add "<REDACTED_ID_16>", "ParseVendor15"
        registry.Add "<REDACTED_ID_06>", "ParseVendor10"
    End If

    If registry.Exists(vendorId) Then
        ObtenerParser = registry(vendorId)
    End If

End Function

Private Function EsVendorMultipaginaEspecial(vendorId) As Boolean

    EsVendorMultipaginaEspecial = (vendorId = "<REDACTED_ID_02>" Or vendorId = "<REDACTED_ID_07>")

End Function

