ARCHIVO: FormRutaCarpeta.frm
RUTA: <REDACTED_PATH>\FormRutaCarpeta.frm
==================================================

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormRutaCarpeta 
   ClientHeight    =   4005
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10800
   OleObjectBlob   =   "FormRutaCarpeta.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "FormRutaCarpeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
' File: FormRutaCarpeta.frm
' Purpose: VBA module extracted from legacy Excel app (portfolio version).
' Note: Cosmetic formatting only. No behavior changes.
'------------------------------------------------------------------------------
Public RutaDescarga As String
Dim objOutlook As Object
Dim objSeleccion As Object
Dim objMail As Object
Dim objAdjunto As Object
Dim numeroCorreos
Private Sub UserForm_Initialize()
    Me.Width = 545
    Me.Height = 210
    rutaActual = Hoja3.Range("Ruta")
    Hoja3.Range("RutaDescarga") = rutaActual
    rutaActual.Enabled = False










































==================================================
                gCtx.nuevoNombre = "Fecha base " & fechaRecepcion & "-" & countNombrePDF & ".pdf"
                
                If Dir(RutaDescarga & gCtx.nuevoNombre) <> "" Then
                    Do
                        countNombrePDF = countNombrePDF + 1
                        gCtx.nuevoNombre = "Fecha base " & fechaRecepcion & "-" & countNombrePDF & ".pdf"
                    Loop While Dir(RutaDescarga & gCtx.nuevoNombre) <> ""
                End If
                
                countDescargas = countDescargas + 1

                Me.Progreso = "Correo " & countCorreo & " de " & numeroCorreos & ": Descargando adjunto " & countDescargas & "..."
                
                objAdjunto.SaveAsFile RutaDescarga & gCtx.nuevoNombre
                
            End If
        Next objAdjunto
            
        If CheckBoxLeido.Value = True Then objMail.Unread = False ' Marcar el correo como leÃ­do
        objMail.Save ' Guardar los cambios

    Next objMail
    
    
    
    ' Liberar la memoria
    Set objAdjunto = Nothing
    Set objMail = Nothing
    Set objSeleccion = Nothing
    Set objOutlook = Nothing
    
    
    
        respuesta = MsgBox("Se descargaron " & countDescargas & " archivos adjuntos y los correos se marcaron como leÃ­dos." & vbCrLf & _
        "Â¿Desea procesar los PDF descargados?", vbYesNo, "Pregunta")
        If respuesta = vbYes Then
            Unload Me
            Importar_1_SeleccionarDocumentos (Hoja3.Range("RutaDescarga"))
        End If

    
    
    'Application.Cursor = xlDefault
    
    Exit Sub
    
errorDisco:
    
    MsgBox "Error " & Err.Number & ": " & Err.Description & vbLf & vbLf & _
    "No se pudo descargar el archivo. Verifique si hay espacio en el disco y vuelva a intentar", vbCritical, "Error"

    Hoja2.Select
    
End Sub
    On Error GoTo errorDisco
    
    countDescargas = 0
    countCorreo = 0
    
    For Each objMail In objSeleccion
    
        countCorreo = countCorreo + 1
        countPDF = 0
    
        fechaRecepcion = Format(objMail.ReceivedTime, "dd.mm.yyyy")
    
        For Each objAdjunto In objMail.Attachments
                
            If UCase(Right(objAdjunto.Filename, 4)) = ".PDF" Then
            
                countNombrePDF = countNombrePDF + 1Sub DescargarAdjuntosYMarcarComoLeido()End Sub    If Right(Me.rutaActual.Value, 1) <> "\" Then Me.rutaActual.Value = Me.rutaActual.Value & "\"Private Sub rutaActual_Change()End Sub    End With        End If            Hoja3.Range("RutaDescarga") = rutaActual            rutaActual = gCtx.rutaCarpeta            gCtx.rutaCarpeta = .SelectedItems(1)        If .Show = -1 Then        .InitialFileName = rutaActual        .AllowMultiSelect = False        .Title = "Seleccione una carpeta"    With Application.FileDialog(msoFileDialogFolderPicker)Private Sub Modificar_Click()End Sub    Unload Me    End If        Exit Sub        MsgBox "La ruta proporcionada NO existe"    Else        DescargarAdjuntosYMarcarComoLeido        RutaDescarga = rutaActual    If Dir(rutaActual, vbDirectory) <> "" ThenPrivate Sub Continuar_Click()Private Sub Cancelar_Click(): Unload Me: End SubEnd Sub    End If        Me.LabelTitulo = "Descargar de Outlook: " & numeroCorreos & " correos seleccionados"    Else        Me.LabelTitulo = "Descargar de Outlook: " & numeroCorreos & " correo seleccionado"    
    If numeroCorreos = 1 Then    
    numeroCorreos = objSeleccion.Count    Set objSeleccion = objOutlook.ActiveExplorer.Selection    Set objOutlook = CreateObject("Outlook.Application")
