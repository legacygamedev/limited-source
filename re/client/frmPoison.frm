VERSION 5.00
Begin VB.Form frmPoison 
   Caption         =   "Poison Editor"
   ClientHeight    =   1320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3105
   Icon            =   "frmPoison.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1320
   ScaleWidth      =   3105
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Ok"
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   120
      Max             =   500
      Min             =   1
      TabIndex        =   0
      Top             =   360
      Value           =   1
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Poison Strength: 1"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmPoison"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    PoisonStrength = HScroll1.Value
    Unload Me
End Sub

Private Sub HScroll1_Change()
    Label1.Caption = "Poison Strength: " & HScroll1.Value
End Sub
