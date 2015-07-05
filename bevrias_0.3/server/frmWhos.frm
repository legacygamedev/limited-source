VERSION 5.00
Begin VB.Form frmWhos 
   Caption         =   "Whos The Creator"
   ClientHeight    =   2655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2655
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Background"
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton Command1 
         Caption         =   "Close"
         Height          =   255
         Left            =   3480
         TabIndex        =   2
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   $"frmWhos.frx":0000
         Height          =   1815
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4215
      End
   End
End
Attribute VB_Name = "frmWhos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmWhos.Visible = False
End Sub
