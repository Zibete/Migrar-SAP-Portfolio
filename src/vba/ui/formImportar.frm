ARCHIVO: formImportar.frm
RUTA: <REDACTED_PATH>\formImportar.frm
==================================================

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formImportar 
   ClientHeight    =   8400.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17220
   OleObjectBlob   =   "formImportar.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "formImportar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
' File: formImportar.frm
' Purpose: VBA module extracted from legacy Excel app (portfolio version).
' Note: Cosmetic formatting only. No behavior changes.
'------------------------------------------------------------------------------
Private esNumerico As Boolean
Private esLetra As Boolean
Private Sub UserForm_Initialize()

    Me.Height = 428
    Me.Width = 741

    For Each vendor_prov In gCtx.rngVendor_Prov.DataBodyRange

























































































































































































==================================================End Sub    Unload MePrivate Sub Salir_Click()End Sub    End If        End If            Exit Sub            KeyAscii = 0        If Len(tb_search) >= 6 Then        End If            Exit Sub            KeyAscii = 0        If Not (KeyAscii >= 48 And KeyAscii <= 57) Then    If esNumerico Then    End If        End If            Exit Sub            esLetra = True            esNumerico = False        Else            esLetra = False            esNumerico = True        If KeyAscii >= 48 And KeyAscii <= 57 Then    If Len(tb_search) = 0 Then    If KeyAscii = 8 Then Exit SubPrivate Sub tb_search_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)End Function    Next i        End If            Exit Function            ExisteEnListBox = True        If lst.List(i, 0) = valor Then    For i = 0 To lst.ListCount - 1    ExisteEnListBox = FalsePrivate Function ExisteEnListBox(lst As MSForms.ListBox, valor As String) As Boolean    
End Sub    Next i        End If            ListBox2.RemoveItem i            End If                ListBox1.List(y, 1) = ListBox2.List(i, 1)                ListBox1.List(y, 0) = ListBox2.List(i, 0)                ListBox1.AddItem                y = ListBox1.ListCount            If Not ExisteEnListBox(ListBox1, ListBox2.List(i, 0)) Then        If moverTodos Or ListBox2.Selected(i) Then    
    For i = ListBox2.ListCount - 1 To 0 Step -1    Next i        End If            Exit For            moverTodos = False        If ListBox2.Selected(i) Then    
    For i = 0 To ListBox2.ListCount - 1    
    moverTodos = TruePrivate Sub SpinButton1_SpinDown()End Sub    Next i        End If            ListBox1.RemoveItem i            End If                ListBox2.List(y, 1) = ListBox1.List(i, 1)                ListBox2.List(y, 0) = ListBox1.List(i, 0)                ListBox2.AddItem                y = ListBox2.ListCount            If Not ExisteEnListBox(ListBox2, ListBox1.List(i, 0)) Then        If moverTodos Or ListBox1.Selected(i) Then    For i = ListBox1.ListCount - 1 To 0 Step -1    Next i        End If            Exit For            moverTodos = False        If ListBox1.Selected(i) Then    For i = 0 To ListBox1.ListCount - 1    moverTodos = TruePrivate Sub SpinButton1_SpinUp()
End Sub    End If        Next nombre_Prov            End If                End If                    
                    i = i + 1                    ListBox1.List(i, 0) = Hoja3.Cells(nombre_Prov.Row, gCtx.rngVendor_Prov.Range.Column)                    ListBox1.List(i, 1) = nombre                    ListBox1.AddItem                If Not ExisteEnListBox(ListBox2, Hoja3.Cells(nombre_Prov.Row, gCtx.rngVendor_Prov.Range.Column)) Then            
            
            
            If UCase(nombre) Like "*" & UCase(tb_search) & "*" Then            If Trim(analistaProv) <> "" Then nombre = nombre & " [" & analistaProv & "]"            If Trim(descripcionProv) <> "" Then nombre = nombre & " (" & descripcionProv & ")"            analistaProv = Hoja3.Cells(nombre_Prov.Row, gCtx.rngAnalista_Prov.Range.Column)            descripcionProv = Hoja3.Cells(nombre_Prov.Row, gCtx.rngDescripcion_Prov.Range.Column)            
            nombre = nombre_Prov        For Each nombre_Prov In gCtx.rngNombre_Prov.DataBodyRange    ElseIf esLetra Then        Next vendor_prov            End If                End If                    
                    i = i + 1                    
                    ListBox1.List(i, 1) = nombreProv                    If Trim(analistaProv) <> "" Then nombreProv = nombreProv & " [" & analistaProv & "]"                    If Trim(descripcionProv) <> "" Then nombreProv = nombreProv & " (" & descripcionProv & ")"                    analistaProv = Hoja3.Cells(vendor_prov.Row, gCtx.rngAnalista_Prov.Range.Column)                    descripcionProv = Hoja3.Cells(vendor_prov.Row, gCtx.rngDescripcion_Prov.Range.Column)                    
                    nombreProv = Hoja3.Cells(vendor_prov.Row, gCtx.rngNombre_Prov.Range.Column)                    ListBox1.List(i, 0) = vendor_prov                    ListBox1.AddItem
    ListBox1.Clear
    
    If esNumerico Then
n        For Each vendor_prov In gCtx.rngVendor_Prov.DataBodyRange
n            If UCase(vendor_prov) Like "*" & UCase(tb_search) & "*" Then
n            If Not ExisteEnListBox(ListBox2, vendor_prov.Value) ThenPrivate Sub tb_search_Change()           
End Sub    
    Unload ProgressBar   
    Call Importar_5_Finalizar(y, "Importar")    End With        .Lbl2.Caption = capt2 & " (" & Format(.pb2.Value / .pb2.Max, "0%") & ")"        capt2 = "Finalizando..."        .Lbl1.Caption = capt1 & " (" & Format(.pb1.Value / .pb1.Max, "0%") & ")"        .pb1.Value = .pb1.Value + 1 
    With ProgressBar       
    Call Importar_3_SB_to_MIGRAR(y, "Importar")    End With        .Lbl2.Caption = capt2 & " (" & Format(.pb2.Value / .pb2.Max, "0%") & ")"        capt2 = "Ordenando los datos descargados"        .Lbl1.Caption = capt1 & " (" & Format(.pb1.Value / .pb1.Max, "0%") & ")"        .pb1.Value = .pb1.Value + 1
    With ProgressBar    End If            
        End If        
            Call AbrirRetailWebCubo            End With                .Lbl2.Caption = capt2 & " (" & Format(.pb2.Value / .pb2.Max, "0%") & ")"                .Lbl1.Caption = capt1 & " (" & Format(.pb1.Value / .pb1.Max, "0%") & ")"                .pb2.Value = 1                .pb2.Max = 2                .pb1.Value = 1                .pb1.Max = .pb1.Max + 1        
            With ProgressBar        
            capt2 = "Descargando nuevos datos"        If Hoja3.Range("mantenerDatos") = "NO" Then    
    If Hoja3.Range("origenDatos") = "RW" Then    End With        .pb1.Max = 2        .Lbl2.Caption = "Preparando todo..." & " (0%)"        .Lbl1.Caption = capt1 & " (0%)"        .Show vbModeless    With ProgressBar    
    'Abrir    
    capt1 = "Importando datos de RetailWeb"    
    Unload Me    
    y = gCtx.tblDatos.Range.Row + 1    gCtx.tblDatos.DataBodyRange.ClearContents    gCtx.tblDatos.AutoFilter.ShowAllData    
    gCtx.ControlarCambios = False    End If        Hoja3.Range("CUIT") = "Varios"        Hoja3.Range("nombreProveedor") = "Varios"        
        Hoja3.Range("Vend") = "Varios"        Next i            gCtx.vendors(i) = ListBox2.List(i, 0)    
        For i = 0 To ListBox2.ListCount - 1    
    Else        Hoja3.Range("nombreProveedor") = ListBox2.List(0, 1)        Hoja3.Range("Vend") = ListBox2.List(0, 0)    
    ReDim gCtx.vendors(ListBox2.ListCount - 1)
        
    If ListBox2.ListCount = 1 Then
n    
        gCtx.vendors(0) = ListBox2.List(0, 0)        
    End If        Exit Sub    
        MsgBox "Seleccione al menos un proveedor"    If ListBox2.ListCount = 0 ThenPrivate Sub btn_Aceptar_Click()    
End Sub    Next vendor_prov           
        i = i + 1        
        ListBox1.List(i, 1) = nombreProv        If Trim(analistaProv) <> "" Then nombreProv = nombreProv & " [" & analistaProv & "]"        If Trim(descripcionProv) <> "" Then nombreProv = nombreProv & " (" & descripcionProv & ")"        analistaProv = Hoja3.Cells(vendor_prov.Row, gCtx.rngAnalista_Prov.Range.Column)        descripcionProv = Hoja3.Cells(vendor_prov.Row, gCtx.rngDescripcion_Prov.Range.Column)        
        nombreProv = Hoja3.Cells(vendor_prov.Row, gCtx.rngNombre_Prov.Range.Column)        ListBox1.List(i, 0) = vendor_prov        ListBox1.AddItem
