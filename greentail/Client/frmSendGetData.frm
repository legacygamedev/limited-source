VERSION 5.00
Begin VB.Form frmSendGetData 
   BackColor       =   &H00789298&
   BorderStyle     =   0  'None
   ClientHeight    =   435
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   4485
   ControlBox      =   0   'False
   Icon            =   "frmSendGetData.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSendGetData.frx":014A
   ScaleHeight     =   29
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   299
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   45
      TabIndex        =   0
      Top             =   135
      Width           =   4380
   End
End
Attribute VB_Name = "frmSendGetData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyEscape) Then
        Call DestroyDirectX
        Call StopMidi
        InGame = False
        frmMirage.Socket.Close
        frmMainMenu.Visible = True
        Connucted = False
        Unload Me
    End If
End Sub

Private Sub Form_Load()
Dim i As Long
Dim Ending As String
    For i = 1 To 3
        If i = 1 Then Ending = ".GIF"
        If i = 2 Then Ending = ".JPG"
        If i = 3 Then Ending = ".PNG"
 
        If FileExist("GUI\Loading" & Ending) Then frmSendGetData.Picture = LoadPicture(App.Path & "\GUI\Loading" & Ending)
    Next i
End Sub

