VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public sheetDataBase As Worksheet

Public ControlarCambios As Boolean

'Tablas
Public tblDatos As ListObject
Public TblProveedores As ListObject
Public tblCondPago As ListObject
Public tblCORS As ListObject
Public tblRet As ListObject
Public tblPercepciones As ListObject
Public tblIndicadores As ListObject
Public tblDataBase As ListObject

Public rutaCarpeta As String
Public textoBtn As String
Public passwordSB As String

Public IE_NuevaVentana As Object
Public IE_VentanaFactura As Object
Public INPUT_stocknumber_IE As Object

Public diccDocumentos As Object


Public NombreArchivo As String

'Alicuotas
Public ali_IIBBSalta
Public ali_IIBBNqn
Public ali_IIBBCABAGen
Public ali_IIBBLaRioja
Public ali_IIBBCABABeb
Public ali_MuniCord
Public ali_PercIVA
Public ali_IIBBCatamarca
Public ali_IIBBEntreRios
Public ali_IIBBCorrientes
Public ali_IIBBMendoza
Public ali_IIBBMisiones

'Rangos
Public rngEstadoDelPago
Public rngPagado
Public rngComentarios_User
Public rngEstado
Public rngDifSap
Public rngDifCostos
Public rngReferencia
Public rngNombreArchivo
Public rngSite
Public rngEstadoCambiado
Public rngCompensacion

Public rngFechaBase
Public rngTexto
Public rngCeBe
Public rngNombreSite
Public rngSupl
Public rngNuevaRuta
Public rngFechaDeFactura
Public rngTotalBrutoFactura
Public rngSubtotalFactura
Public rngII
Public rngDifConNC
Public rngIVA
Public rngVTOCAE
Public rngCAE
Public rngIVA105
Public rngSubtotalFactura105

Public rngObservacionesDelPago_SB

'DataBase
Public rngRetailWeb_DB
Public rngRefPDF_DB
Public rngReferencia_DB
Public rngSite_DB
Public rngTipoDoc_DB
Public rngVendor_DB
Public rngFecha_DB
Public rngTotal_DB
Public rngSubtotal_DB
Public rngII_DB
Public rngIVA_DB
Public rngPerc1_DB
Public rngPerc2_DB
Public rngPerc3_DB
Public rngPerc4_DB
Public rngCAE_DB
Public rngVTOCAE_DB
Public rngFechaBase_DB
Public rngEstado_DB
Public rngComentarios_DB

'IIBB
Public rngIIBBBSAS
Public rngIIBBCABA
Public rngIIBBChubut
Public rngIIBBTucuman
Public rngIIBBSalta
Public rngIIBBNeuquen
Public rngIIBBSantaFe
Public rngIIBBCatamarca
Public rngIIBBChaco
Public rngIIBBCordoba
Public rngIIBBCorrientes
Public rngIIBBEntreRios
Public rngIIBBFormosa
Public rngIIBBJujuy
Public rngIIBBLaPampa
Public rngIIBBLaRioja
Public rngIIBBMendoza
Public rngIIBBMisiones
Public rngIIBBRioNegro
Public rngIIBBSanJuan
Public rngIIBBSantiago
Public rngIIBBSanLuis
Public rngIIBBSantaCruz
Public rngIIBBTierraDelFuego
Public rngMuniCord
Public rngPercIVA

'Proveedores
Public rngVendor_Prov
Public rngNombre_Prov
Public rngAnalista_Prov
Public rngEsPyme_Prov
Public rngCondPago_Prov
Public rngCUIT_Prov
Public rngDescripcion_Prov

'Cond Pago
Public rngCod_CondPago
Public rngDescripcion_CondPago

'Percepciones
Public rngTP_Perc
Public rngAlicuota_Perc
Public rngDenominacion_Perc

'RetailWeb
Public rngVendorProveedor_SB
Public rngNombreProveedor_SB
Public rngRetailWeb_SB
Public rngFechaNeg_SB
Public rngFechaDoc_SB
Public rngSite_SB
Public rngTotalBruto_SB
Public rngSubtotal_SB
Public rngValorizado_SB
Public rngTieneScan_SB
Public rngEstadoDelPago_SB
Public rngComentarios_SB
Public rngZona
Public rngAN
Public rngMails

Public rngRemitoRef
Public rngMensajesSap
Public rngFechaBaseCambiada
Public rngTipoDoc

Public timeout
Public linkSB
Public dominio
Public reporteSB As Boolean
Public endoso As Boolean
Public ultimaFila

Public vendorActual

Public montoToleranciaSAP
Public montoToleranciaSB
Public montoDOA
Public montoFCE

Public nuevoNombre
Public largoReferencia
Public letra

Public SELECTION_USER As Integer

Public TARGET_SELECTION

'Diccionarios
Public dictALICUOTAS As Object

Public PASS


Public vendors() As String ' Declarar el array
Public alicuotas() As String
