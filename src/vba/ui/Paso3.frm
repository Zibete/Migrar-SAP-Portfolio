VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Paso3 
   ClientHeight    =   6600
   ClientLeft      =   -345
   ClientTop       =   -1500
   ClientWidth     =   9705.001
   OleObjectBlob   =   "Paso3.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Paso3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
' File: Paso3.frm
' Purpose: VBA module extracted from legacy Excel app (portfolio version).
' Note: Cosmetic formatting only. No behavior changes.
'------------------------------------------------------------------------------
Dim interrumpirEjecucion As Boolean
Public registros
Public CompensacionFC
Public nuevaDiferencia
Public valorNetoFC, valorNetoNC
Public valorDiferencia
Public TOLERANCIA
Public contabilizarAFavor As Boolean
Public selección
Public IND

Public filaNC
Public filaFC

Dim ESCIG As Boolean
Dim ESBEB As Boolean
Dim ESMED As Boolean

Dim Errores(27), Avisos(27), mensajeError, mensajeAviso As String

Private Sub Salir_Click()
    Unload Me
    ProtectHoja2ForUi
End Sub

Private Sub UserForm_Initialize()

    asignaciones

    Dim esPyme As Boolean
    
    btnInterrumpir.Enabled = False
    
    Me.Height = 358
    Me.Width = 490
    
    Me.Lista.Height = 114
    Me.Lista.Clear
    
    Me.LabelNuevaDiferencia.ForeColor = RGB(0, 0, 0)
    Me.SeCancelan.Enabled = False
    Me.AsociarFCyNC = False
    
    If gCtx.SELECTION_USER = 2 Then

        For y = 1 To Selection.Rows.Count

            fila = Selection.Cells(y, 1).Row
            Estado = Hoja2.Cells(fila, gCtx.rngEstado.Range.Column)

            If Estado <> "Revisar datos" And Estado <> "" And Hoja2.Rows(fila).EntireRow.Hidden = False Then
                If Left(Hoja2.Cells(fila, gCtx.rngTipoDoc.Range.Column), 2) = "FC" Then
                
                    filaFC = fila

                    valorDiferencia = Format(Hoja2.Cells(fila, gCtx.rngDifCostos.Range.Column), "##,##0.00")
                    
                    neto21 = Hoja2.Cells(fila, gCtx.rngSubtotalFactura.Range.Column)
                    neto105 = Hoja2.Cells(fila, gCtx.rngSubtotalFactura105.Range.Column)
                    netoII = Hoja2.Cells(fila, gCtx.rngII.Range.Column)
                    
                    valorNetoFC = Format(neto21 * 1 + neto105 * 1 + netoII * 1, "##,##0.00")

                    siteFC = Hoja2.Cells(fila, gCtx.rngSite.Range.Column)
                    
                ElseIf Left(Hoja2.Cells(fila, gCtx.rngTipoDoc.Range.Column), 2) = "NC" Then
                
                    filaNC = fila
                
                    neto21 = Hoja2.Cells(fila, gCtx.rngSubtotalFactura.Range.Column)
                    neto105 = Hoja2.Cells(fila, gCtx.rngSubtotalFactura105.Range.Column)
                    netoII = Hoja2.Cells(fila, gCtx.rngII.Range.Column)
                    
                    valorNetoNC = Format(neto21 * 1 + neto105 * 1 + netoII * 1, "##,##0.00")

                    siteNC = Hoja2.Cells(fila, gCtx.rngSite.Range.Column)
                    
                End If
            End If
            
        Next y

        If valorDiferencia <> "" And valorNetoNC <> "" Then
            If siteFC = siteNC Then

                Me.Lista.Height = 35
                Me.AsociarFCyNC = True
                
                If valorNetoFC = valorNetoNC Then
                    SeCancelan.Enabled = True
                    SeCancelan = True
                End If
            
            End If
        End If
    End If
    
    Dim dictPERC As Object
    Set dictPERC = CreateObject("Scripting.Dictionary")
    
    i = 0
    
    For y = 1 To Selection.Rows.Count
    
        dictPERC.RemoveAll
    
        fila = Selection.Cells(y, 1).Row
                
        Referencia = Hoja2.Cells(fila, gCtx.rngReferencia.Range.Column)
        Estado = Hoja2.Cells(fila, gCtx.rngEstado.Range.Column)
        tipoDoc = Hoja2.Cells(fila, gCtx.rngTipoDoc.Range.Column)
        
        If Estado <> "Revisar datos" And Estado <> "" And Referencia <> "" And Estado <> "Completar" _
        And Estado <> "Eliminado" And Hoja2.Rows(fila).EntireRow.Hidden = False Then
        
            Vendor = Hoja3.Range("Vend")
            
            If Vendor = "Varios" Then Vendor = Hoja2.Cells(fila, gCtx.rngVendorProveedor_SB.Range.Column)
            
            esPyme = True
            ESCIG = False
            ESBEB = False
            ESMED = False

            Set rngProveedor = gCtx.rngVendor_Prov.DataBodyRange.Find(What:=Vendor, LookAt:=xlWhole)

            If Hoja3.Cells(rngProveedor.Row, gCtx.rngEsPyme_Prov.Range.Column) = "NO" Then esPyme = False
            If Hoja3.Cells(rngProveedor.Row, gCtx.rngDescripcion_Prov.Range.Column) = "Cigarrillos" Then ESCIG = True
            If Hoja3.Cells(rngProveedor.Row, gCtx.rngDescripcion_Prov.Range.Column) = "Bebidas" Then ESBEB = True
            If Hoja3.Cells(rngProveedor.Row, gCtx.rngDescripcion_Prov.Range.Column) = "Medicamentos" Then ESMED = True
            
            
            With dictPERC
                .Add "J100", Hoja2.Cells(fila, gCtx.rngIIBBBSAS.Range.Column)
                If ESCIG Then .Add "J101Cig", Hoja2.Cells(fila, gCtx.rngIIBBCABA.Range.Column)
                If Not ESCIG Then .Add "J101", Hoja2.Cells(fila, gCtx.rngIIBBCABA.Range.Column)
                .Add "J102", Hoja2.Cells(fila, gCtx.rngIIBBChubut.Range.Column)
                .Add "J103", Hoja2.Cells(fila, gCtx.rngIIBBTucuman.Range.Column)
                .Add "J104", Hoja2.Cells(fila, gCtx.rngIIBBSalta.Range.Column)
                .Add "J105", Hoja2.Cells(fila, gCtx.rngIIBBNeuquen.Range.Column)
                .Add "J106", Hoja2.Cells(fila, gCtx.rngIIBBSantaFe.Range.Column)
                .Add "J107", Hoja2.Cells(fila, gCtx.rngIIBBCatamarca.Range.Column)
                .Add "J108", Hoja2.Cells(fila, gCtx.rngIIBBChaco.Range.Column)
                .Add "J109", Hoja2.Cells(fila, gCtx.rngIIBBCordoba.Range.Column)
                .Add "J110", Hoja2.Cells(fila, gCtx.rngIIBBCorrientes.Range.Column)
                .Add "J111", Hoja2.Cells(fila, gCtx.rngIIBBEntreRios.Range.Column)
                .Add "J112", Hoja2.Cells(fila, gCtx.rngIIBBFormosa.Range.Column)
                .Add "J113", Hoja2.Cells(fila, gCtx.rngIIBBJujuy.Range.Column)
                .Add "J114", Hoja2.Cells(fila, gCtx.rngIIBBLaPampa.Range.Column)
                .Add "J115", Hoja2.Cells(fila, gCtx.rngIIBBLaRioja.Range.Column)
                .Add "J116", Hoja2.Cells(fila, gCtx.rngIIBBMendoza.Range.Column)
                .Add "J117", Hoja2.Cells(fila, gCtx.rngIIBBMisiones.Range.Column)
                .Add "J118", Hoja2.Cells(fila, gCtx.rngIIBBRioNegro.Range.Column)
                .Add "J119", Hoja2.Cells(fila, gCtx.rngIIBBSanJuan.Range.Column)
                .Add "J120", Hoja2.Cells(fila, gCtx.rngIIBBSantiago.Range.Column)
                .Add "J121", Hoja2.Cells(fila, gCtx.rngIIBBSanLuis.Range.Column)
                .Add "J122", Hoja2.Cells(fila, gCtx.rngIIBBSantaCruz.Range.Column)
                .Add "J123", Hoja2.Cells(fila, gCtx.rngIIBBTierraDelFuego.Range.Column)
                .Add "MCOR", Hoja2.Cells(fila, gCtx.rngMuniCord.Range.Column)
                .Add "J1AP", Hoja2.Cells(fila, gCtx.rngPercIVA.Range.Column)
                .Add "IVA", Hoja2.Cells(fila, gCtx.rngIVA.Range.Column)
                .Add "IVA105", Hoja2.Cells(fila, gCtx.rngIVA105.Range.Column)
                .Add "II", Hoja2.Cells(fila, gCtx.rngII.Range.Column)
            End With
            
            IND = ""
        
            For Each IND_ITE In gCtx.tblIndicadores.ListColumns
                If IND_ITE.index <> 1 Then
                    For Each PERC In gCtx.tblIndicadores.ListRows
                        PERC_COD = PERC.Range(1, 1).Value
                        PERC_DIC = dictPERC(PERC_COD)
                        ALI_NEC = PERC.Range(1, IND_ITE.index)
                        If ALI_NEC <> "" And PERC_DIC <> "" Then
                            If gCtx.dictALICUOTAS(PERC_COD) = ALI_NEC Then
                                IND = Left(IND_ITE.Name, 2)
                            Else
                                IND = ""
                                Debug.Print ("NO es " & IND_ITE & " porque " & PERC_COD & " pide alícuota " & ALI_NEC & " y la alícuota está fijada en " & gCtx.dictALICUOTAS(PERC_COD))
                                Exit For
                            End If
                        ElseIf ALI_NEC = "" And PERC_DIC = "" Then
                            IND = Left(IND_ITE.Name, 2)
                        Else
                            Debug.Print ("NO es " & IND_ITE & " porque " & PERC_COD & " pide """ & ALI_NEC & """ y en diccionario hay " & """" & PERC_DIC & """")
                            IND = ""
                            Exit For
                        End If
                    Next PERC
                End If
                If IND <> "" Then Exit For
            Next IND_ITE
        
            If IND = "" Then IND = "Z0"
            Debug.Print ("Es: " & IND)
            
'-----------------------------
            Total = Hoja2.Cells(fila, gCtx.rngTotalBrutoFactura.Range.Column).Value
            Total = Format(Total, "##,##0.00")
                
            Me.Lista.AddItem
            Me.Lista.List(i, 0) = i + 1 'Num
            Me.Lista.List(i, 1) = Hoja2.Cells(fila, gCtx.rngFechaDeFactura.Range.Column).Value 'Fecha
            Me.Lista.List(i, 3) = Total 'Importe
            Me.Lista.List(i, 4) = Hoja2.Cells(fila, gCtx.rngSupl.Range.Column).Value 'Supl
            Me.Lista.List(i, 5) = Hoja2.Cells(fila, gCtx.rngSite.Range.Column).Value 'Sucursal
            
            Me.Lista.List(i, 6) = IND

            If Left(tipoDoc, 2) = "FC" Then
                If esPyme And CDbl(Total) >= gCtx.montoFCE Then
                    claseDoc = "X7"
                    If Len(Referencia) = 14 Then Referencia = Mid(Referencia, 2)
                    countX7 = countX7 + 1
                Else
                    claseDoc = "XL"
                    countXL = countXL + 1
                End If
            ElseIf Left(tipoDoc, 2) = "NC" Then
                If esPyme And Me.AsociarFCyNC Then
                    If CDbl(Hoja2.Cells(filaFC, gCtx.rngTotalBrutoFactura.Range.Column)) >= gCtx.montoFCE Then
                        claseDoc = "X8"
                        If Len(Referencia) = 14 Then Referencia = Mid(Referencia, 2)
                        countX8 = countX8 + 1
                    Else
                        claseDoc = "XM"
                        countXM = countXM + 1
                    End If
                Else
                    If Left(tipoDoc, 3) = "NCE" Then
                        claseDoc = "X8"
                        If Len(Referencia) = 14 Then Referencia = Mid(Referencia, 2)
                        countX8 = countX8 + 1
                    Else
                        claseDoc = "XM"
                        countXM = countXM + 1
                    End If
                End If
            ElseIf Left(tipoDoc, 2) = "ND" Then
                If Left(tipoDoc, 3) = "NDE" Then
                    claseDoc = "X9"
                    If Len(Referencia) = 14 Then Referencia = Mid(Referencia, 2)
                    countX9 = countX9 + 1
                Else
                    claseDoc = "XN"
                    countXN = countXN + 1
                End If
            End If
            
            Me.Lista.List(i, 2) = UCase(Referencia)
            Me.Lista.List(i, 7) = claseDoc
            
            i = i + 1
        
        End If
    Next y
    
    registros = i
    
    LabelRegistros = "Se van a contabilizar " & registros & " registros del proveedor " & Hoja3.Range("nombreProveedor")
    
    If countXL <> 0 Then LabelXL = "XL: " & countXL
    If countXN <> 0 Then LabelXN = "XN: " & countXN
    If countXM <> 0 Then LabelXM = "XM: " & countXM
    If countX7 <> 0 Then LabelX7 = "X7: " & countX7
    If countX8 <> 0 Then LabelX8 = "X8: " & countX8
    If countX9 <> 0 Then LabelX9 = "X9: " & countX9
    
End Sub

Private Sub AsociarFCyNC_Change()

    Me.LabelDiferencia = valorDiferencia
    
    If AsociarFCyNC Then
    
        If Not SeCancelan Then

            Me.Lbl1 = "Diferencia Factura:"
            Me.Lbl2 = "Neto Nota de Crédito:"
            Me.Lbl3 = "Diferencia:"
        
            SeCancelan = False
            
            NC = valorNetoNC
            
            Me.LabelNC = NC
            nuevaDiferencia = Format(valorDiferencia - NC, "##,##0.00")
            Me.LabelNuevaDiferencia = nuevaDiferencia
        
            If CDbl(nuevaDiferencia) >= gCtx.montoToleranciaSB Then
                Me.LabelNuevaDiferencia.ForeColor = RGB(255, 0, 0)
            Else
                Me.LabelNuevaDiferencia.ForeColor = RGB(0, 0, 0)
            End If

        End If

    Else
    
        SeCancelan = False
    
        NC = "0,00"
    
        NC = valorNetoNC
        
        Me.LabelNC = NC
        nuevaDiferencia = Format(valorDiferencia - NC, "##,##0.00")
        Me.LabelNuevaDiferencia = nuevaDiferencia
    
        If CDbl(nuevaDiferencia) >= gCtx.montoToleranciaSB Then
            Me.LabelNuevaDiferencia.ForeColor = RGB(255, 0, 0)
        Else
            Me.LabelNuevaDiferencia.ForeColor = RGB(0, 0, 0)
        End If
       
    End If
    
End Sub
Private Sub SeCancelan_Change()

    If SeCancelan = True Then
    
        AsociarFCyNC = True
    
        Me.Lbl1 = "Neto Factura:"
        Me.Lbl2 = "Neto Nota de Crédito:"
        Me.Lbl3 = "Diferencia:"
    
        Me.LabelDiferencia = valorNetoFC
        Me.LabelNC = valorNetoNC
        Me.LabelNuevaDiferencia = CDbl(valorNetoFC) - CDbl(valorNetoNC)
    
    Else
    
        AsociarFCyNC_Change
    
    End If

End Sub
Private Sub btnInterrumpir_Click()

    interrumpirEjecucion = True
    Progreso = "Interrumpiendo ejecución..."
    
End Sub

End Sub
