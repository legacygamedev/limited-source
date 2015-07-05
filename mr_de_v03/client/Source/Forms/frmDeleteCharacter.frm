VERSION 5.00
Begin VB.Form frmDeleteCharacter 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delete Account"
   ClientHeight    =   855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   1815
   Icon            =   "frmDeleteCharacter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   1815
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Back"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblDelete 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Delete Character"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmDeleteCharacter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Unload(Cancel As Integer)
    Me.Visible = False
    frmChars.Visible = True
End Sub

Private Sub Label2_Click()
    Me.Visible = False
    frmChars.Visible = True
End Sub

Private Sub lblDelete_Click()
    Me.Visible = False
    frmChars.Visible = False
    If ConnectToServer = True Then
        Call SetStatus("Connected, sending adventurer deletion request...")
        Call SendDelChar(frmChars.lstChars.ListIndex + 1)
    End If
End Sub
