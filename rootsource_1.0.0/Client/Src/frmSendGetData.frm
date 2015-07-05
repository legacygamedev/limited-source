VERSION 5.00
Begin VB.Form frmSendGetData 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   825
   ClientLeft      =   0
   ClientTop       =   -15
   ClientWidth     =   4785
   Icon            =   "frmSendGetData.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSendGetData.frx":08CA
   ScaleHeight     =   55
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   319
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   225
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   4035
   End
End
Attribute VB_Name = "frmSendGetData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Caption = GAME_NAME & " (esc to cancel)"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Call DestroyTCP
        frmSendGetData.Hide
        frmMainMenu.Show
    End If
End Sub

' When the form close button is pressed
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Call DestroyTCP
        frmSendGetData.Hide
        frmMainMenu.Show
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MoveForm(Me)
End Sub
