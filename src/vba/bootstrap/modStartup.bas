Attribute VB_Name = "modStartup"

Sub Auto_Open()

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    UnprotectHoja2Safe
    
'    Hoja3.Range("passwordSB") = ""
    
    LabelTitulo = "** " & Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1) & " **"
    
    Hoja2.Shapes("nombreLibro").TextFrame2.TextRange.Text = LabelTitulo
    
    asignaciones

    Hoja2.EnableOutlining = True
    ProtectHoja2ForUi
    
    
    Application.Windows(ThisWorkbook.Name).DisplayHeadings = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    Application.Windows(ThisWorkbook.Name).DisplayWorkbookTabs = False
    
End Sub
Public Sub asignaciones()

    ResetContext

    LabelTitulo = "** " & Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1) & " **"
    
    Hoja2.Shapes("nombreLibro").TextFrame2.TextRange.Text = LabelTitulo

    gCtx.ControlarCambios = True
    gCtx.timeout = False
    gCtx.reporteSB = False
    
    gCtx.montoToleranciaSAP = Hoja3.Range("montoToleranciaSAP")
    gCtx.montoToleranciaSB = Hoja3.Range("montoToleranciaSB")
    gCtx.montoDOA = Hoja3.Range("montoDOA")
    gCtx.montoFCE = Hoja3.Range("montoFCE")
    
    Set gCtx.sheetDataBase = ThisWorkbook.Sheets("sheetDataBase")
    
    Set gCtx.diccDocumentos = CreateObject("Scripting.Dictionary")

    Set gCtx.tblDatos = Hoja2.ListObjects("tblDatos")
    Set gCtx.TblProveedores = Hoja3.ListObjects("tblProveedores")
    Set gCtx.tblCondPago = Hoja3.ListObjects("tblCondPago")
    Set gCtx.tblCORS = Hoja3.ListObjects("tblCors")
    Set gCtx.tblRet = Hoja3.ListObjects("tblRet")
    Set gCtx.tblPercepciones = Hoja3.ListObjects("tblPercepciones")
    Set gCtx.tblIndicadores = Hoja3.ListObjects("tblIndicadores")
    Set gCtx.tblDataBase = gCtx.sheetDataBase.ListObjects("tblDataBase")
    
    'Percepciones
    Set gCtx.rngTP_Perc = gCtx.tblPercepciones.ListColumns("TP. Perc.")
    Set gCtx.rngDenominacion_Perc = gCtx.tblPercepciones.ListColumns("Denominación Percepción")
    Set gCtx.rngAlicuota_Perc = gCtx.tblPercepciones.ListColumns("Alícuota Percepción")
   
    'Proveedores
    Set gCtx.rngVendor_Prov = gCtx.TblProveedores.ListColumns("Vendor")
    Set gCtx.rngNombre_Prov = gCtx.TblProveedores.ListColumns("Nombre del proveedor")
    Set gCtx.rngAnalista_Prov = gCtx.TblProveedores.ListColumns("Analista")
    Set gCtx.rngDescripcion_Prov = gCtx.TblProveedores.ListColumns("Descripción")
    Set gCtx.rngEsPyme_Prov = gCtx.TblProveedores.ListColumns("¿Es Pyme?")
    Set gCtx.rngCondPago_Prov = gCtx.TblProveedores.ListColumns("Cond. Pago")
    Set gCtx.rngCUIT_Prov = gCtx.TblProveedores.ListColumns("CUIT")
    
    'Cond Pago
    Set gCtx.rngCod_CondPago = gCtx.tblCondPago.ListColumns("Cod. Cond. Pago")
    Set gCtx.rngDescripcion_CondPago = gCtx.tblCondPago.ListColumns("Descripción Cond. Pago")

    'DataBase
    Set gCtx.rngRetailWeb_DB = gCtx.tblDataBase.ListColumns("RetailWeb")
    Set gCtx.rngRefPDF_DB = gCtx.tblDataBase.ListColumns("RefPDF")
    Set gCtx.rngReferencia_DB = gCtx.tblDataBase.ListColumns("Referencia")
    Set gCtx.rngSite_DB = gCtx.tblDataBase.ListColumns("Sucursal")
    Set gCtx.rngTipoDoc_DB = gCtx.tblDataBase.ListColumns("TipoDoc")
    Set gCtx.rngVendor_DB = gCtx.tblDataBase.ListColumns("Vendor")
    Set gCtx.rngFecha_DB = gCtx.tblDataBase.ListColumns("Fecha")
    Set gCtx.rngTotal_DB = gCtx.tblDataBase.ListColumns("Total")
    Set gCtx.rngSubtotal_DB = gCtx.tblDataBase.ListColumns("Subtotal")
    Set gCtx.rngII_DB = gCtx.tblDataBase.ListColumns("II")
    Set gCtx.rngIVA_DB = gCtx.tblDataBase.ListColumns("IVA")
    Set gCtx.rngPerc1_DB = gCtx.tblDataBase.ListColumns("Perc1")
    Set gCtx.rngPerc2_DB = gCtx.tblDataBase.ListColumns("Perc2")
    Set gCtx.rngPerc3_DB = gCtx.tblDataBase.ListColumns("Perc3")
    Set gCtx.rngPerc4_DB = gCtx.tblDataBase.ListColumns("Perc4")
    Set gCtx.rngCAE_DB = gCtx.tblDataBase.ListColumns("CAE")
    Set gCtx.rngVTOCAE_DB = gCtx.tblDataBase.ListColumns("VTOCAE")
    Set gCtx.rngFechaBase_DB = gCtx.tblDataBase.ListColumns("FechaBase")
    Set gCtx.rngEstado_DB = gCtx.tblDataBase.ListColumns("Estado")
    Set gCtx.rngComentarios_DB = gCtx.tblDataBase.ListColumns("Comentarios")

    'RetailWeb
    Set gCtx.rngVendorProveedor_SB = gCtx.tblDatos.ListColumns("Vendor Proveedor")
    Set gCtx.rngNombreProveedor_SB = gCtx.tblDatos.ListColumns("Nombre Proveedor")
    Set gCtx.rngRetailWeb_SB = gCtx.tblDatos.ListColumns("RetailWeb")
    Set gCtx.rngFechaNeg_SB = gCtx.tblDatos.ListColumns("Fecha Negocio" & vbLf & "(RetailWeb)")
    Set gCtx.rngFechaDoc_SB = gCtx.tblDatos.ListColumns("Fecha" & vbLf & "Documento" & vbLf & "(RetailWeb)")
    Set gCtx.rngSite_SB = gCtx.tblDatos.ListColumns("Sucursal" & vbLf & "(RetailWeb)")
    Set gCtx.rngTotalBruto_SB = gCtx.tblDatos.ListColumns("Total" & vbLf & "Bruto" & vbLf & "(RetailWeb)")
    Set gCtx.rngSubtotal_SB = gCtx.tblDatos.ListColumns("Subtotal" & vbLf & "(RetailWeb)")
    Set gCtx.rngValorizado_SB = gCtx.tblDatos.ListColumns("Valorizado" & vbLf & "(RetailWeb)")
    Set gCtx.rngTieneScan_SB = gCtx.tblDatos.ListColumns("¿Tiene" & vbLf & "scan?" & vbLf & "(RetailWeb)")
    Set gCtx.rngEstadoDelPago_SB = gCtx.tblDatos.ListColumns("Estado del Pago" & vbLf & "(RetailWeb)")
    Set gCtx.rngComentarios_SB = gCtx.tblDatos.ListColumns("Comentarios" & vbLf & "(RetailWeb)")
    Set gCtx.rngZona = gCtx.tblDatos.ListColumns("Zona")
    Set gCtx.rngAN = gCtx.tblDatos.ListColumns("AN")
    Set gCtx.rngMails = gCtx.tblDatos.ListColumns("Mails")
    Set gCtx.rngObservacionesDelPago_SB = gCtx.tblDatos.ListColumns("Observaciones del Pago" & vbLf & "(RetailWeb)")

    Set gCtx.rngDifCostos = gCtx.tblDatos.ListColumns("Diferencia" & vbLf & "VS" & vbLf & "RetailWeb")
    Set gCtx.rngDifConNC = gCtx.tblDatos.ListColumns("Diferencia" & vbLf & "con NC" & vbLf & "asociada")
    Set gCtx.rngDifSap = gCtx.tblDatos.ListColumns("Diferencia" & vbLf & "SAP")
    Set gCtx.rngNombreArchivo = gCtx.tblDatos.ListColumns("Nombre Archivo")
    Set gCtx.rngComentarios_User = gCtx.tblDatos.ListColumns("Comentarios" & vbLf & "(Usuario)")
    Set gCtx.rngSite = gCtx.tblDatos.ListColumns("Sucursal")
    Set gCtx.rngEstado = gCtx.tblDatos.ListColumns("Estado")
    Set gCtx.rngEstadoDelPago = gCtx.tblDatos.ListColumns("Estado del Pago" & vbLf & "(Usuario)")
    Set gCtx.rngPagado = gCtx.tblDatos.ListColumns("¿Pagado?")
    Set gCtx.rngEstadoCambiado = gCtx.tblDatos.ListColumns("¿Estado" & vbLf & "cambiado?")
    Set gCtx.rngCompensacion = gCtx.tblDatos.ListColumns("Compensación")
    Set gCtx.rngReferencia = gCtx.tblDatos.ListColumns("Referencia")
    Set gCtx.rngFechaBase = gCtx.tblDatos.ListColumns("Fecha" & vbLf & "base")
    Set gCtx.rngTexto = gCtx.tblDatos.ListColumns("Texto")
    Set gCtx.rngCeBe = gCtx.tblDatos.ListColumns("CeBe")
    Set gCtx.rngNombreSite = gCtx.tblDatos.ListColumns("Nombre Sucursal")
    Set gCtx.rngSupl = gCtx.tblDatos.ListColumns("Supl.")
    Set gCtx.rngNuevaRuta = gCtx.tblDatos.ListColumns("Nueva Ruta")
    Set gCtx.rngFechaDeFactura = gCtx.tblDatos.ListColumns("Fecha")
    Set gCtx.rngTotalBrutoFactura = gCtx.tblDatos.ListColumns("Total" & vbLf & "Bruto")
    Set gCtx.rngII = gCtx.tblDatos.ListColumns("Impuestos" & vbLf & "Internos")
    
    'Percepciones
    Set gCtx.rngIIBBBSAS = gCtx.tblDatos.ListColumns("IIBB" & vbLf & "BS. AS." & vbLf & "J100")
    Set gCtx.rngIIBBCABA = gCtx.tblDatos.ListColumns("IIBB" & vbLf & "CABA" & vbLf & "J101")
    Set gCtx.rngIIBBChubut = gCtx.tblDatos.ListColumns("IIBB" & vbLf & "Chubut" & vbLf & "J102")
    Set gCtx.rngIIBBTucuman = gCtx.tblDatos.ListColumns("IIBB" & vbLf & "Tucumán" & vbLf & "J103")
    Set gCtx.rngIIBBSalta = gCtx.tblDatos.ListColumns("IIBB" & vbLf & "Salta" & vbLf & "J104")
    Set gCtx.rngIIBBNeuquen = gCtx.tblDatos.ListColumns("IIBB" & vbLf & "Neuquén" & vbLf & "J105")
    Set gCtx.rngIIBBSantaFe = gCtx.tblDatos.ListColumns("IIBB" & vbLf & "Santa Fé" & vbLf & "J106")
    Set gCtx.rngIIBBCatamarca = gCtx.tblDatos.ListColumns("IIBB" & vbLf & "Catamarca" & vbLf & "J107")
    Set gCtx.rngIIBBChaco = gCtx.tblDatos.ListColumns("IIBB" & vbLf & "Chaco" & vbLf & "J108")
    Set gCtx.rngIIBBCordoba = gCtx.tblDatos.ListColumns("IIBB" & vbLf & "Córdoba" & vbLf & "J109")
    Set gCtx.rngIIBBCorrientes = gCtx.tblDatos.ListColumns("IIBB" & vbLf & "Corrientes" & vbLf & "J110")
    Set gCtx.rngIIBBEntreRios = gCtx.tblDatos.ListColumns("IIBB" & vbLf & "Entre Ríos" & vbLf & "J111")
    Set gCtx.rngIIBBFormosa = gCtx.tblDatos.ListColumns("IIBB" & vbLf & "Formosa" & vbLf & "J112")
    Set gCtx.rngIIBBJujuy = gCtx.tblDatos.ListColumns("IIBB" & vbLf & "Jujuy" & vbLf & "J113")
    Set gCtx.rngIIBBLaPampa = gCtx.tblDatos.ListColumns("IIBB" & vbLf & "La Pampa" & vbLf & "J114")
    Set gCtx.rngIIBBLaRioja = gCtx.tblDatos.ListColumns("IIBB" & vbLf & "La Rioja" & vbLf & "J115")
    Set gCtx.rngIIBBMendoza = gCtx.tblDatos.ListColumns("IIBB" & vbLf & "Mendoza" & vbLf & "J116")
    Set gCtx.rngIIBBMisiones = gCtx.tblDatos.ListColumns("IIBB" & vbLf & "Misiones" & vbLf & "J117")
    Set gCtx.rngIIBBRioNegro = gCtx.tblDatos.ListColumns("IIBB" & vbLf & "Rio Negro" & vbLf & "J118")
    Set gCtx.rngIIBBSanJuan = gCtx.tblDatos.ListColumns("IIBB" & vbLf & "San Juan" & vbLf & "J119")
    Set gCtx.rngIIBBSantiago = gCtx.tblDatos.ListColumns("IIBB" & vbLf & "Santiago del Estero" & vbLf & "J120")
    Set gCtx.rngIIBBSanLuis = gCtx.tblDatos.ListColumns("IIBB" & vbLf & "San Luis" & vbLf & "J121")
    Set gCtx.rngIIBBSantaCruz = gCtx.tblDatos.ListColumns("IIBB" & vbLf & "Santa Cruz" & vbLf & "J122")
    Set gCtx.rngIIBBTierraDelFuego = gCtx.tblDatos.ListColumns("IIBB" & vbLf & "Tierra del Fuego" & vbLf & "J123")
    Set gCtx.rngMuniCord = gCtx.tblDatos.ListColumns("Perc. Munic. Córdoba" & vbLf & "MCOR")
    Set gCtx.rngPercIVA = gCtx.tblDatos.ListColumns("Perc. IVA" & vbLf & "J1AP")

    
    Set gCtx.rngIVA = gCtx.tblDatos.ListColumns("IVA" & vbLf & "21,00%")
    Set gCtx.rngSubtotalFactura = gCtx.tblDatos.ListColumns("Subtotal" & vbLf & "con IVA 21,00%")
    
    Set gCtx.rngIVA105 = gCtx.tblDatos.ListColumns("IVA" & vbLf & "10,50%")
    Set gCtx.rngSubtotalFactura105 = gCtx.tblDatos.ListColumns("Subtotal" & vbLf & "con IVA 10,50%")

    Set gCtx.rngRemitoRef = gCtx.tblDatos.ListColumns("Referencia" & vbLf & "RetailWeb")
    Set gCtx.rngVTOCAE = gCtx.tblDatos.ListColumns("VTO. CAE")
    Set gCtx.rngCAE = gCtx.tblDatos.ListColumns("CAE")
    Set gCtx.rngMensajesSap = gCtx.tblDatos.ListColumns("Mensajes" & vbLf & "SAP")
    Set gCtx.rngFechaBaseCambiada = gCtx.tblDatos.ListColumns("¿Fecha base cambiada?")
    Set gCtx.rngTipoDoc = gCtx.tblDatos.ListColumns("Tipo" & vbLf & "Doc.")

    gCtx.rutaCarpeta = GetRutaCarpeta()
    gCtx.linkSB = GetConfigValue("LinkSB", "")
    gCtx.dominio = Left(gCtx.linkSB, InStr(gCtx.linkSB, ".com") - 1)
    
    gCtx.ultimaFila = gCtx.tblDatos.ListRows.Count + gCtx.tblDatos.DataBodyRange.Row - 1
    
    Call largoyletraRef
    
    Set gCtx.dictALICUOTAS = CreateObject("Scripting.Dictionary")

    For Each filaPerc In gCtx.tblPercepciones.ListRows
        PERC = filaPerc.Range(1, gCtx.rngTP_Perc.index)
        ALIC = filaPerc.Range(1, gCtx.rngAlicuota_Perc.index)
        gCtx.dictALICUOTAS(PERC) = CDbl(ALIC)
    Next filaPerc

End Sub
