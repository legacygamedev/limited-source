VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Main Form"
   ClientHeight    =   5055
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8535.001
   LinkTopic       =   "Form1"
   ScaleHeight     =   337
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   569
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOG 
         Caption         =   "Open Graphic"
      End
      Begin VB.Menu mnuOS 
         Caption         =   "Open Sprite"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case vbKeyLeft:
        MsgBox "left"
        SM3D.SetTriStripTopLeft Pos_X - 1, Pos_Y, 32, 32
    Case vbKeyRight:
        SM3D.SetTriStripTopLeft Pos_X + 1, Pos_Y, 32, 32
    Case vbKeyUp:
        SM3D.SetTriStripTopLeft Pos_X, Pos_Y - 1, 32, 32
    Case vbKeyDown:
        SM3D.SetTriStripTopLeft Pos_X, Pos_Y + 1, 32, 32
End Select

MsgBox KeyAscii
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SM_RENDER = False
    Set SM3D = Nothing
End Sub

Private Sub picScreen_Click()
    SM_RENDER = False
    Set SM3D = Nothing
End Sub

Private Sub mnuQuit_Click()
    SM_RENDER = False
    Set SM3D = Nothing
    Unload Me
End Sub
