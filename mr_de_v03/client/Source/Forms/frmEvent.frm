VERSION 5.00
Begin VB.Form frmEvent 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alert"
   ClientHeight    =   900
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5250
   ControlBox      =   0   'False
   Icon            =   "frmEvent.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   900
   ScaleWidth      =   5250
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblInformation 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Alert"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001D2B34&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   4785
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Ok"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2055
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "frmEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        MenuState CurrentState
    End If
End Sub

Private Sub lblMenu_Click()
    MenuState CurrentState
End Sub
