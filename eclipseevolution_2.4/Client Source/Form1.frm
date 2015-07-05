VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   8085
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   4600
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   4545
      ScaleWidth      =   2445
      TabIndex        =   0
      Top             =   0
      Width           =   2500
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim var() As String
Dim stringy As String

stringy = "," & "2a" & ",hasoi"
var = Split(stringy, ",")
Label1.Caption = var(0)
Label2.Caption = var(1)
End Sub
