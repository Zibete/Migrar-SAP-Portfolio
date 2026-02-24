ARCHIVO: form_BloqueoB.frm
RUTA: <REDACTED_PATH>\form_BloqueoB.frm
==================================================

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_BloqueoB 
   Caption         =   "UserForm1"
   ClientHeight    =   7395
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11175
   OleObjectBlob   =   "form_BloqueoB.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "form_BloqueoB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
' File: form_BloqueoB.frm
' Purpose: VBA module extracted from legacy Excel app (portfolio version).
' Note: Cosmetic formatting only. No behavior changes.
'------------------------------------------------------------------------------
Dim objOutlook As Object
Dim objSeleccion As Object
Dim objMail As Object
Dim objAdjunto As Object
Dim DOCS_SAP As Collection
Dim DOCS_COUNT
Dim i
Private Sub btn_Leido_Click()

    For Each objMail In objSeleccion
        cuerpo = objMail.Body
        If objMail.Subject = "Error de validaciÃ³n de CAI-CAE-CAEA" Then
            If InStr(cuerpo, "Documento Nr.: " & DOCS_SAP(i)) > 0 Then
                If objMail.Unread = True Then
                    objMail.Unread = False
                    btn_Leido.Caption = "Marcar no leÃ­do"
                    Exit For
                End If
                If objMail.Unread = False Then
                    objMail.Unread = True
                    btn_Leido.Caption = "Marcar leÃ­do"
                    Exit For
                End If
            End If
        End If
    Next objMail




















==================================================Private Sub btn_Anterior_Click()
    btn_Leido.Caption = "Marcar leÃ­do"
    i = i - 1
    lbl_Titulo = "Verificar bloqueos B: Correo " & i & " de " & DOCS_COUNT
    tb_DOC_SAP = DOCS_SAP(i)
    tb_Proveedor = ""
    tb_CUIT = ""
    tb_Referencia = ""
    tb_Fecha = ""
    tb_Total = ""
    tb_TipoDoc = ""
    tb_CAE = ""
    btn_VerificarARCA.Enabled = False
    If i = DOCS_COUNT Then btn_Siguiente.Enabled = False
    If i < DOCS_COUNT Then btn_Siguiente.Enabled = True
    If i = 1 Then btn_Anterior.Enabled = False
    If i > 1 Then btn_Anterior.Enabled = True
End SubPrivate Sub btn_Salir_Click(): Unload Me: End Sub
Private Sub btn_Siguiente_Click()
    btn_Leido.Caption = "Marcar leÃ­do"
    i = i + 1
    lbl_Titulo = "Verificar bloqueos B: Correo " & i & " de " & DOCS_COUNT
    tb_DOC_SAP = DOCS_SAP(i)
    tb_Proveedor = ""
    tb_CUIT = ""
    tb_Referencia = ""
    tb_Fecha = ""
    tb_Total = ""
    tb_TipoDoc = ""
    tb_CAE = ""
    btn_VerificarARCA.Enabled = False
    If i = DOCS_COUNT Then btn_Siguiente.Enabled = False
    If i < DOCS_COUNT Then btn_Siguiente.Enabled = True
    If i = 1 Then btn_Anterior.Enabled = False
    If i > 1 Then btn_Anterior.Enabled = True
End Sub
    If Not IsObject(App) Then
        Set SapGuiAuto = GetObject("SAPGUI")
        Set App = SapGuiAuto.GetScriptingEngine
    End If
    If Not IsObject(Connection) Then Set Connection = App.Children(0)
    If Not IsObject(session) Then Set session = Connection.Children(0)
    If IsObject(WScript) Then
        WScript.ConnectObject session, "on"
        WScript.ConnectObject App, "on"
    End If
    
    session.findById("wnd[0]").resizeWorkingPane 97, 22, False

    session.findById("wnd[0]/tbar[0]/okcd").Text = "/NFB03"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/txtRF05L-BELNR").Text = tb_DOC_SAP
    session.findById("wnd[0]/usr/ctxtRF05L-BUKRS").Text = "<REDACTED>"
    session.findById("wnd[0]").sendVKey 0
    On Error Resume Next
    session.findById("wnd[1]").sendVKey 0
    On Error GoTo 0
    
    tb_Referencia = session.findById("wnd[0]/usr/txtBKPF-XBLNR").Text
    tb_Fecha = session.findById("wnd[0]/usr/ctxtBKPF-BLDAT").Text
    
    TTL = session.findById("wnd[0]/usr/cntlCTRL_CONTAINERBSEG/shellcont/shell").GetCellValue(0, "DMBTR")
    VDR = session.findById("wnd[0]/usr/cntlCTRL_CONTAINERBSEG/shellcont/shell").GetCellValue(0, "KTONR")
    NMB = session.findById("wnd[0]/usr/cntlCTRL_CONTAINERBSEG/shellcont/shell").GetCellValue(0, "KOBEZ")
    
    tb_Total = Replace(TTL, "-", "")
    tb_Proveedor = NMB
    
    Set rngProveedor = gCtx.rngVendor_Prov.DataBodyRange.Find(What:=VDR, LookAt:=xlWhole)
    tb_CUIT = Hoja3.Cells(rngProveedor.Row, gCtx.rngCUIT_Prov.Range.Column)
    tb_Proveedor = NMB
    
    session.findById("wnd[0]/tbar[1]/btn[5]").press
    
    tb_TipoDoc = session.findById("wnd[1]/usr/ctxtBKPF-BLART").Text
    
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[0]").sendVKey 12
    
    lbl_Progreso = "Buscando en SAP: zzzzlocsapbc_corrcai..."
    
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/Nzzzzlocsapbc_corrcai"
    
    session.findById("wnd[0]").sendVKey 0
    
    session.findById("wnd[0]/usr/ctxtP_BUKRS").Text = "<REDACTED>"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/txtS_BELNR-LOW").Text = tb_DOC_SAP
    session.findById("wnd[0]/usr/txtS_GJAHR-LOW").Text = Year(Date)
    session.findById("wnd[0]/usr/ctxtP_EVENT").Text = "SAPA2"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    tb_CAE = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").GetCellValue(0, "J_1APAC")
    
    session.findById("wnd[0]").sendVKey 12
    session.findById("wnd[0]").sendVKey 12
    
    lbl_Progreso = ""
    
End Sub
    lbl_Progreso = "Buscando en SAP: FB03..."Sub buscar_SAP()
    Width = 365
    Height = 355
    
    lbl_Progreso = "Iniciando..."
    
    Set objOutlook = CreateObject("Outlook.Application")
    Set objSeleccion = objOutlook.ActiveExplorer.Selection
    Set DOCS_SAP = New Collection
    
    For Each objMail In objSeleccion
        cuerpo = objMail.Body
        If objMail.Subject = "Error de validaciÃ³n de CAI-CAE-CAEA" Then
            If InStr(cuerpo, "Documento Nr.:") > 0 Then
                DOC_SAP = Mid(cuerpo, InStr(cuerpo, "Documento Nr.:") + Len("Documento Nr.: "), 10)
                DOCS_SAP.Add DOC_SAP
                Debug.Print DOC_SAP
            End If
        End If
    Next objMail
    
    i = 1
    
    DOCS_COUNT = DOCS_SAP.Count
    
    If DOCS_COUNT > 0 Then
        lbl_Titulo = "Verificar bloqueos B: Correo 1 de " & DOCS_COUNT
        tb_DOC_SAP = DOCS_SAP(i)
    Else
        lbl_Titulo = "Verificar bloqueos B: No se encontrÃ³ ningÃºn correo"
        tb_DOC_SAP = ""
    End If
    
    btn_Siguiente.Enabled = False
    btn_Anterior.Enabled = False
    btn_VerificarARCA.Enabled = False
    btn_Leido.Caption = "Marcar leÃ­do"
    
    If DOCS_COUNT > 1 Then btn_Siguiente.Enabled = True
    
    lbl_Progreso = ""
    
End SubPrivate Sub UserForm_Initialize()    Application.Cursor = xlWait
    lbl_Progreso = "Abriendo SAP..."
    
    Call ABRIRSAP
    Call buscar_SAP
    
    btn_VerificarARCA.Enabled = True
    
    lbl_Progreso = ""
    Application.Cursor = xlDefault
End SubPrivate Sub btn_VerificarSAP_Click()
    Application.Cursor = xlDefault
    
End Sub
    objShell.Run cmd
    If Not CheckBox_CAEA Then LINK_ARCA = "https://servicioscf.afip.gob.ar/publico/comprobantes/cae.aspx"
    If CheckBox_CAEA Then LINK_ARCA = "https://servicioscf.afip.gob.ar/publico/comprobantes/caea.aspx"
    
    Fecha = Replace(tb_Fecha, ".", "/")
    
    If tb_TipoDoc = "XL" Then TP_DOC = "1"
    If tb_TipoDoc = "XM" Then TP_DOC = "3"
    If tb_TipoDoc = "X7" Then TP_DOC = "201"
    If tb_TipoDoc = "X8" Then TP_DOC = "203"
    
    LTR = InStr(1, tb_Referencia, "A")
    PDV = Left(tb_Referencia, LTR - 1)
    COMP = Mid(tb_Referencia, LTR + 1)
    TTL = tb_Total
    TTL = Replace(TTL, ",", "")
    
    Dim objShell As Object
    Set objShell = CreateObject("WScript.Shell")
  
    cmd = """" & GetPythonwExePath() & """ """ & ResolveScriptPath("ARCA.py") & """ """ _
    & LINK_ARCA & """ """ _
    & tb_CUIT & """ """ & tb_CAE & """ """ & Fecha & """ """ & TP_DOC & """ """ _
    & PDV & """ """ & COMP & """ """ & TTL & """ """ & Hoja3.Range("CUIT" & Chr$(80) & Chr$(65) & Chr$(69)) & """    
    Application.Cursor = xlWait
    lbl_Progreso = "Verificando en ARCA..."Private Sub btn_VerificarARCA_Click()nEnd Sub
