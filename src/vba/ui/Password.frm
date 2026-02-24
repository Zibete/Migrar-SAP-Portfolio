VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Password 
   ClientHeight    =   2700
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5160
   OleObjectBlob   =   "Password.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
' File: Password.frm
' Purpose: VBA module extracted from legacy Excel app (portfolio version).
' Note: Cosmetic formatting only. No behavior changes.
'------------------------------------------------------------------------------

Private Sub EntrarAdmin_Click()
    Dim adminPwd As String

    adminPwd = GetWorkbookUnprotectPassword()

    If UCase(Me.user) = "XMAP07" Then

        If adminPwd = "" Then
            MENSAJE = MsgBox("La password de administrador no esta configurada. Defina MIGRAR_PASSWORD en el entorno.", vbExclamation)
            Exit Sub
        End If

        If Me.Password = adminPwd Then
       
            UnprotectHoja2Safe
    
            nombre = ThisWorkbook.Name
            Application.Windows(nombre).DisplayWorkbookTabs = True
            Application.Windows(nombre).DisplayHeadings = True
            Application.ScreenUpdating = True
            
            gCtx.ControlarCambios = False
            
            Unload Me
        
        Else
            MENSAJE = MsgBox("La password es incorrecta", vbCritical)
        End If
    Else
        MENSAJE = MsgBox("El usuario no es admin", vbCritical)
    End If

End Sub

Private Sub VolverAdmin_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Me.Width = 270
    Me.Height = 165
    Me.user = Environ("UserName")
End Sub
