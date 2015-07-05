VERSION 5.00
Begin VB.Form frmSendGetData 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "M:RPGe (Loading...)"
   ClientHeight    =   6555
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   7665
   ControlBox      =   0   'False
   Icon            =   "frmSendGetData.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   7665
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image3 
      Height          =   195
      Left            =   7320
      Picture         =   "frmSendGetData.frx":0442
      Top             =   120
      Width           =   195
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   0
      Picture         =   "frmSendGetData.frx":25E8
      Top             =   0
      Width           =   7680
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   735
      Left            =   1560
      TabIndex        =   0
      Top             =   3360
      Width           =   4575
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   6195
      Left            =   0
      Picture         =   "frmSendGetData.frx":51B7
      Top             =   360
      Width           =   7680
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
        End
    End If
End Sub

Private Sub Image3_Click()
Unload Me
Set frmSendGetData = Nothing
End Sub
