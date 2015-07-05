VERSION 5.00
Begin VB.Form frmStatus 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   1605
   ClientLeft      =   0
   ClientTop       =   -15
   ClientWidth     =   6495
   ControlBox      =   0   'False
   Icon            =   "frmStatus.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   6435
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ------------------------------------------
' --               Asphodel               --
' ------------------------------------------

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Not KeyAscii = vbKeyEscape Then Exit Sub
    Select Case CurrentWindow
        Case 0
            DestroyGame
        Case Else
            Windows(CurrentWindow).Visible = True
            Me.Visible = False
    End Select
End Sub
