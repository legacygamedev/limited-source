VERSION 5.00
Begin VB.Form frmCredits 
   Caption         =   "Credits"
   ClientHeight    =   1335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1335
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Credits"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton Command1 
         Caption         =   "Close"
         Height          =   255
         Left            =   3480
         TabIndex        =   1
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Johansson_tk@hotmail.com - Creators MSN"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   3615
      End
      Begin VB.Label Label1 
         Caption         =   "www.Bevrias.com - Offical site for Bevrias ORPGE"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   3615
      End
   End
End
Attribute VB_Name = "frmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmCredits.Visible = False
End Sub
