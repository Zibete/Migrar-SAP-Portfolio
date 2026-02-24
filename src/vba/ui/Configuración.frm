VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Configuración 
   ClientHeight    =   10425
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15480
   OleObjectBlob   =   "Configuración.frx":0000
   StartUpPosition =   2  'Centrar en pantalla
End
Attribute VB_Name = "Configuración"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim suprimirEventoChange As Boolean
Private Sub EliminarDuplicadosNO_Click()
    Hoja3.Range("EliminarDuplicados") = "NO"
End Sub
Private Sub EliminarDuplicadosSI_Click()

    If suprimirEventoChange Then Exit Sub
    
    respuesta = MsgBox("Cuando se encuentren archivos duplicados, se procederá con su eliminación. ¿Desea continuar?", vbYesNo + vbQuestion, "Confirmación")

    If respuesta = vbYes Then
        Hoja3.Range("EliminarDuplicados") = "SI"
    End If
        
End Sub

Private Sub Frame5_Click()

End Sub

Private Sub PagoPendienteNO_Click()
    Hoja3.Range("PagoPendiente") = "NO"
End Sub

Private Sub PagoPendienteSI_Click()
    Hoja3.Range("PagoPendiente") = "SI"
End Sub

Private Sub PagoPendienteTODOS_Click()
    Hoja3.Range("PagoPendiente") = "TODOS"
End Sub

Private Sub ToggleButton_CUBO_Click()

    If ToggleButton_CUBO.Value = True Then
        ToggleButton_SB.Value = False
        Hoja3.Range("origenDatos") = "CUBO"
        ToggleButton_DatosNO.Value = True
        ToggleButton_DatosSI.Enabled = False
        ToggleButton_DatosNO.Enabled = False
    Else
        ToggleButton_SB.Value = True
    End If
    
End Sub

Private Sub ToggleButton_SB_Click()

    If ToggleButton_SB.Value = True Then
        ToggleButton_CUBO.Value = False
        Hoja3.Range("origenDatos") = "RW"
        ToggleButton_DatosSI.Enabled = True
        ToggleButton_DatosNO.Enabled = True
    Else
        ToggleButton_CUBO.Value = True
    End If
    
End Sub
Private Sub ToggleButton_DatosNO_Click()

    If ToggleButton_DatosNO.Value = True Then
        ToggleButton_DatosSI.Value = False
    Else
        ToggleButton_DatosSI.Value = True
        Exit Sub
    End If
    
    Hoja3.Range("mantenerDatos") = "NO"
    
End Sub
Private Sub ToggleButton_DatosSI_Click()

    If ToggleButton_DatosSI.Value = True Then
        ToggleButton_DatosNO.Value = False
    Else
        ToggleButton_DatosNO.Value = True
        Exit Sub
    End If
    
    Hoja3.Range("mantenerDatos") = "SI"
    
End Sub

Private Sub UserForm_Initialize()

    suprimirEventoChange = True

    Me.Height = 363
    Me.Width = 605

    LabelTitulo = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)
    tbUser = Environ("UserName")
    tbPassword = GetPass()
    
    tb_montoFCE = Format(gCtx.montoFCE, "##,##0.00")
    tb_montoDOA = Format(gCtx.montoDOA, "##,##0.00")
    tb_montoToleranciaSB = Format(gCtx.montoToleranciaSB, "##,##0.00")
    
    tb_montoToleranciaSAP = gCtx.montoToleranciaSAP
    tb_CUITPae = Hoja3.Range("CUIT" & Chr$(80) & Chr$(65) & Chr$(69))

    btnEditar1.Caption = "Editar"
    btnEditar2.Caption = "Editar"

    If Hoja3.Range("PagoPendiente") = "SI" Then PagoPendienteSI = True
    If Hoja3.Range("PagoPendiente") = "TODOS" Then PagoPendienteTODOS = True
    If Hoja3.Range("PagoPendiente") = "NO" Then PagoPendienteNO = True
    
    If Hoja3.Range("EliminarDuplicados") = "SI" Then EliminarDuplicadosSI = True
    If Hoja3.Range("EliminarDuplicados") = "NO" Then EliminarDuplicadosNO = True
    
    If Hoja3.Range("origenDatos") = "CUBO" Then
        
        ToggleButton_CUBO.Value = True
        ToggleButton_DatosSI.Enabled = False
        ToggleButton_DatosNO.Enabled = False
        Hoja3.Range("mantenerDatos") = "NO"
        ToggleButton_DatosNO.Value = True
        
    End If
    
    If Hoja3.Range("origenDatos") = "RW" Then ToggleButton_SB.Value = True
    
    If Hoja3.Range("mantenerDatos") = "SI" Then ToggleButton_DatosSI.Value = True
    If Hoja3.Range("mantenerDatos") = "NO" Then ToggleButton_DatosNO.Value = True
            
    suprimirEventoChange = False

End Sub
Private Sub btnAdmin_Click()
    Unload Me
    Password.Show
End Sub
Private Sub btnEditar1_Click()

    If btnEditar1.Caption = "Guardar cambios" Then
    
        tbPassword.Enabled = False
        tbPassword.PasswordChar = "*"
        btnEditar1.Caption = "Editar"
        Hoja3.Range("PasswordSB") = tbPassword
        
    ElseIf btnEditar1.Caption = "Editar" Then
    
        tbPassword.Enabled = True
        tbPassword.PasswordChar = ""
        
    End If
    
End Sub
Private Sub btnEditar2_Click()

    If btnEditar2.Caption = "Guardar cambios" Then
    
        tb_montoFCE.Enabled = False
        tb_montoDOA.Enabled = False
        tb_montoToleranciaSB.Enabled = False
        
        gCtx.montoDOA = CDbl(Replace(tb_montoDOA, ".", ""))
        gCtx.montoFCE = CDbl(Replace(tb_montoFCE, ".", ""))
        gCtx.montoToleranciaSB = CDbl(Replace(tb_montoToleranciaSB, ".", ""))
        
        tb_montoFCE = Format(gCtx.montoFCE, "##,##0.00")
        tb_montoDOA = Format(gCtx.montoDOA, "##,##0.00")
        tb_montoToleranciaSB = Format(gCtx.montoToleranciaSB, "##,##0.00")
        
        Hoja3.Range("montoDOA") = gCtx.montoDOA
        Hoja3.Range("montoFCE") = gCtx.montoFCE
        Hoja3.Range("montoToleranciaSB") = gCtx.montoToleranciaSB

        btnEditar2.Caption = "Editar"
    
    ElseIf btnEditar2.Caption = "Editar" Then
    
        tb_montoFCE.Enabled = True
        tb_montoDOA.Enabled = True
        tb_montoToleranciaSB.Enabled = True
    
    End If
    
End Sub
Private Sub tbPassword_Change()
    btnEditar1.Caption = "Guardar cambios"
End Sub
Private Sub tb_montoFCE_Change()
    If suprimirEventoChange Then Exit Sub
    btnEditar2.Caption = "Guardar cambios"
    gCtx.montoFCE = Replace(gCtx.montoFCE, ".", ",")
End Sub
Private Sub tb_montoDOA_Change()
    If suprimirEventoChange Then Exit Sub
    btnEditar2.Caption = "Guardar cambios"
    gCtx.montoDOA = Replace(gCtx.montoDOA, ".", ",")
End Sub
Private Sub tb_montoToleranciaSB_Change()
    If suprimirEventoChange Then Exit Sub
    btnEditar2.Caption = "Guardar cambios"
    montoTolerancia = Replace(montoTolerancia, ".", ",")
End Sub
Private Sub tb_montoFCE_AfterUpdate()
    suprimirEventoChange = True
    gCtx.montoFCE = Format(gCtx.montoFCE, "##,##0.00")
    suprimirEventoChange = False
End Sub
Private Sub tb_montoDOA_AfterUpdate()
    suprimirEventoChange = True
    gCtx.montoDOA = Format(gCtx.montoDOA, "##,##0.00")
    suprimirEventoChange = False
End Sub
Private Sub tb_montoToleranciaSB_AfterUpdate()
    suprimirEventoChange = True
    montoTolerancia = Format(montoTolerancia, "##,##0.00")
    suprimirEventoChange = False
End Sub

Private Sub Salir_Click()
    Unload Me
    ProtectHoja2ForUi
End Sub
Private Sub tb_montoFCE_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    SoloNumeros KeyAscii, gCtx.montoFCE
End Sub
Private Sub tb_montoDOA_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    SoloNumeros KeyAscii, gCtx.montoDOA
End Sub
Private Sub tb_montoToleranciaSB_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    SoloNumeros KeyAscii, montoTolerancia
End Sub
Public Sub SoloNumeros(KeyAscii As MSForms.ReturnInteger, ByVal contenido As String)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 8 And KeyAscii <> 46 Then
        KeyAscii = 0
    End If
    If KeyAscii = 46 And InStr(contenido, ",") > 0 Then
        KeyAscii = 0
    End If
End Sub
