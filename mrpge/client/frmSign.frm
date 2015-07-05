VERSION 5.00
Begin VB.Form frmSign 
   BorderStyle     =   0  'None
   Caption         =   "Sign"
   ClientHeight    =   3120
   ClientLeft      =   2250
   ClientTop       =   3285
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   208
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C00000&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   0
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3120
      Width           =   4695
   End
   Begin VB.TextBox txtMessage 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H00FFFFFF&
      Height          =   3135
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmSign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdOK_LostFocus()
Unload Me
End Sub

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub

Private Sub txtMessage_GotFocus()
cmdOK.SetFocus
End Sub

