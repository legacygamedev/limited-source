VERSION 5.00
Begin VB.Form frmAFK 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Steel Warrior"
   ClientHeight    =   1725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "frmAFK.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1725
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "You have been caught AFK training, and have been automatically jailed by the server."
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      Height          =   255
      Left            =   3600
      TabIndex        =   0
      Top             =   1440
      Width           =   855
   End
End
Attribute VB_Name = "frmAFK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()
End
End Sub
