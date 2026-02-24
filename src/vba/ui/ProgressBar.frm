VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressBar 
   ClientHeight    =   1800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7665
   OleObjectBlob   =   "ProgressBar.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub pb1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As stdole.OLE_XPOS_PIXELS, ByVal y As stdole.OLE_YPOS_PIXELS)

End Sub

Private Sub UserForm_Initialize()

    Me.Height = 120 'alto
    Me.Width = 396 'ancho

End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If CloseMode = 0 Then
        Cancel = True
    End If
    
End Sub
