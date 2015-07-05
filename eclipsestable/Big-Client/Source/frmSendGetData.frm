VERSION 5.00
Begin VB.Form frmSendGetData 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   ClientHeight    =   435
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   4425
   ControlBox      =   0   'False
   ForeColor       =   &H00E0E0E0&
   Icon            =   "frmSendGetData.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSendGetData.frx":0FC2
   ScaleHeight     =   29
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   295
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblQuit 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Menu"
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
      Left            =   240
      TabIndex        =   0
      Top             =   105
      Width           =   3945
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
        Call GameDestroy
    End If
End Sub

Private Sub Form_Load()
    Dim I As Long
    Dim Ending As String
    For I = 1 To 3
        If I = 1 Then
            Ending = ".gif"
        End If
        If I = 2 Then
            Ending = ".jpg"
        End If
        If I = 3 Then
            Ending = ".png"
        End If

        If FileExists("GUI\Loading" & Ending) Then
            frmSendGetData.Picture = LoadPicture(App.Path & "\GUI\Loading" & Ending)
        End If
    Next I
End Sub

Private Sub lblQuit_Click()
    Call GameDestroy
End Sub
