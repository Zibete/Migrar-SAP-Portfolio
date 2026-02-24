VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Password_SB 
   Caption         =   "Password RetailWeb"
   ClientHeight    =   1980
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4890
   OleObjectBlob   =   "Password_SB.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Password_SB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
' File: Password_SB.frm
' Purpose: VBA module extracted from legacy Excel app (portfolio version).
' Note: Cosmetic formatting only. No behavior changes.
'------------------------------------------------------------------------------

Private Sub Entrar_Click()

    If Me.passwordSB <> "" Then
        Hoja3.Range("passwordSB") = Me.passwordSB
        Unload Me
    Else
        MENSAJE = MsgBox("Ingrese una password", vbCritical)
    End If

End Sub

Private Sub passwordSB_Change()
    If Me.passwordSB = "" Then
        Me.Entrar.Enabled = False
    Else
        Me.Entrar.Enabled = True
    End If
End Sub

Private Sub UserForm_Initialize()
    Me.Width = 255
    Me.Height = 130
    If Me.passwordSB = "" Then
        Me.Entrar.Enabled = False
    Else
        Me.Entrar.Enabled = True
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True ' Evita que el formulario se cierre
    End If
End Sub
