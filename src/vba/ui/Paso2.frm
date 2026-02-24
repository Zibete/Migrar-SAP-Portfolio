VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Paso2 
   ClientHeight    =   8640.001
   ClientLeft      =   150
   ClientTop       =   585
   ClientWidth     =   14250
   OleObjectBlob   =   "Paso2.frx":0000
   StartUpPosition =   2  'Centrar en pantalla
End
Attribute VB_Name = "Paso2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
' File: Paso2.frm
' Purpose: VBA module extracted from legacy Excel app (portfolio version).
' Note: Cosmetic formatting only. No behavior changes.
'------------------------------------------------------------------------------

Private Sub UserForm_Initialize()

    asignaciones

    Me.Height = 325
    Me.Width = 438

    For y = 1 To 8
        Me.Controls("lbl_num" & y).Visible = False
        Me.Controls("lbl_cod" & y).Visible = False
        Me.Controls("lbl_perc" & y).Visible = False
        Me.Controls("tb_ali_al" & y).Visible = False
        Me.Controls("tb_ali_doc" & y).Visible = False
    Next y
            
    y = 1
    
    For Each encabezado In gCtx.tblDatos.HeaderRowRange
        If (Left(encabezado.Value, 4) = "IIBB" Or Left(encabezado.Value, 4) = "Perc") Then
            If encabezado.EntireColumn.Hidden = False Then
                For i = 1 To Selection.Rows.Count
                    fila = Selection.Cells(i, 1).Row
                    PERC_DOC = Hoja2.Cells(fila, encabezado.Column)
                    If PERC_DOC <> "" And Selection.Cells(i, 1).EntireRow.Hidden = False Then
                    
                        Me.Controls("lbl_perc" & y) = Left(Replace(encabezado.Value, vbLf, " "), Len(encabezado.Value) - 5)
                        
                        PERC = Right(encabezado.Value, 4)
                        
                        SUBT = Hoja2.Cells(fila, gCtx.rngSubtotalFactura.Range.Column)
                        II = Hoja2.Cells(fila, gCtx.rngII.Range.Column)
                        
                        If PERC <> "MCOR" Then IISUB = CDbl(II) + CDbl(SUBT)
                        If PERC = "MCOR" Then IISUB = CDbl(SUBT)

                        Me.Controls("tb_ali_doc" & y) = Format((PERC_DOC / IISUB) * 100, "0.00000")
                                                
                        If PERC = "J101" And II <> "" Then 'Tiene II
                            If II > SUBT Then PERC = "J101Cig" 'CIGARRILLOS
                        End If
                        
                        Me.Controls("lbl_cod" & y) = PERC
                        Me.Controls("tb_ali_al" & y) = Format(gCtx.dictALICUOTAS(PERC), "0.00000")
                        
                        Me.Controls("lbl_num" & y).Visible = True
                        Me.Controls("lbl_cod" & y).Visible = True
                        Me.Controls("lbl_perc" & y).Visible = True
                        Me.Controls("tb_ali_al" & y).Visible = True
                        Me.Controls("tb_ali_doc" & y).Visible = True
                        
                        y = y + 1
                        Exit For
                        
                    End If
                Next i
            End If
        End If
    Next encabezado

    Call VerificarDatosTextBox
    
    btn_continuar.Caption = "Continuar"

End Sub

Private Sub VerificarDatosTextBox()
    Dim textBox As MSForms.control
    Dim todosLosTextBoxConDatos As Boolean
    todosLosTextBoxConDatos = True
    For Each textBox In Me.Controls
        If TypeName(textBox) = "TextBox" And textBox.Visible Then
            If Trim(textBox.Value) = "" Then
                todosLosTextBoxConDatos = False
                Exit For
            End If
        End If
    Next textBox
    btn_continuar.Enabled = todosLosTextBoxConDatos
End Sub

Private Sub btn_Salir_Click()
    Unload Me
End Sub

Private Sub btn_continuar_Click()

    If btn_continuar.Caption <> "Continuar" Then
        Dim textBox As MSForms.control
        For Each textBox In Me.Controls
            If TypeName(textBox) = "TextBox" And textBox.Visible And textBox.Enabled Then
                If textBox.Value <> "" Then
                    i = Right(textBox.Name, 1)
                    PERC = Me.Controls("lbl_cod" & i)
                    ALIC = textBox.Value
                    For Each filaPerc In gCtx.tblPercepciones.ListRows
                        codPerc = filaPerc.Range(1, gCtx.rngTP_Perc.index)
                        If codPerc = PERC Then
                            filaPerc.Range(1, gCtx.rngAlicuota_Perc.index) = CDbl(ALIC) 'Actualizamos tabla
                            gCtx.dictALICUOTAS(PERC) = CDbl(ALIC) 'Actualizamos dicc
                            Exit For
                        End If
                    Next filaPerc
                End If
            End If
        Next textBox
    End If

    Unload Me
    Paso3.Show
    
End Sub

Private Sub tb_ali_al1_AfterUpdate(): tb_ali_al1 = Format(tb_ali_al1, "0.00000"): End Sub

Private Sub tb_ali_al2_AfterUpdate(): tb_ali_al2 = Format(tb_ali_al2, "0.00000"): End Sub

Private Sub tb_ali_al3_AfterUpdate(): tb_ali_al3 = Format(tb_ali_al3, "0.00000"): End Sub

Private Sub tb_ali_al4_AfterUpdate(): tb_ali_al4 = Format(tb_ali_al4, "0.00000"): End Sub

Private Sub tb_ali_al5_AfterUpdate(): tb_ali_al5 = Format(tb_ali_al5, "0.00000"): End Sub

Private Sub tb_ali_al6_AfterUpdate(): tb_ali_al6 = Format(tb_ali_al6, "0.00000"): End Sub

Private Sub tb_ali_al7_AfterUpdate(): tb_ali_al7 = Format(tb_ali_al7, "0.00000"): End Sub

Private Sub tb_ali_al8_AfterUpdate(): tb_ali_al8 = Format(tb_ali_al8, "0.00000"): End Sub

Private Sub tb_ali_al1_Change(): VerificarDatosTextBox: btn_continuar.Caption = "Guardar y Continuar": End Sub

Private Sub tb_ali_al2_Change(): VerificarDatosTextBox: btn_continuar.Caption = "Guardar y Continuar": End Sub

Private Sub tb_ali_al3_Change(): VerificarDatosTextBox: btn_continuar.Caption = "Guardar y Continuar": End Sub

Private Sub tb_ali_al4_Change(): VerificarDatosTextBox: btn_continuar.Caption = "Guardar y Continuar": End Sub

Private Sub tb_ali_al5_Change(): VerificarDatosTextBox: btn_continuar.Caption = "Guardar y Continuar": End Sub

Private Sub tb_ali_al6_Change(): VerificarDatosTextBox: btn_continuar.Caption = "Guardar y Continuar": End Sub

Private Sub tb_ali_al7_Change(): VerificarDatosTextBox: btn_continuar.Caption = "Guardar y Continuar": End Sub

Private Sub tb_ali_al8_Change(): VerificarDatosTextBox: btn_continuar.Caption = "Guardar y Continuar": End Sub

Private Sub tb_ali_al1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger): SoloNumeros KeyAscii, tb_ali_al1: End Sub

Private Sub tb_ali_al2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger): SoloNumeros KeyAscii, tb_ali_al2: End Sub

Private Sub tb_ali_al3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger): SoloNumeros KeyAscii, tb_ali_al3: End Sub

Private Sub tb_ali_al4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger): SoloNumeros KeyAscii, tb_ali_al4: End Sub

Private Sub tb_ali_al5_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger): SoloNumeros KeyAscii, tb_ali_al5: End Sub

Private Sub tb_ali_al6_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger): SoloNumeros KeyAscii, tb_ali_al6: End Sub

Private Sub tb_ali_al7_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger): SoloNumeros KeyAscii, tb_ali_al7: End Sub

Private Sub tb_ali_al8_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger): SoloNumeros KeyAscii, tb_ali_al8: End Sub

Public Sub SoloNumeros(KeyAscii As MSForms.ReturnInteger, ByVal contenido As String)
    If KeyAscii = 46 And contenido = "" Then KeyAscii = 0
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 8 And KeyAscii <> 46 Then
        KeyAscii = 0
    End If
    If KeyAscii = 46 Then KeyAscii = 44
    If KeyAscii = 44 And InStr(contenido, ",") > 0 Then KeyAscii = 0
End Sub
