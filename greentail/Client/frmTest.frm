VERSION 5.00
Begin VB.Form frmTest 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2.146
   ScaleMode       =   0  'User
   ScaleWidth      =   10.435
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   2640
      TabIndex        =   0
      Top             =   960
      Width           =   2175
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim AppPath As String
       If Right(AppPath, 1) <> "\" Then AppPath = AppPath & "\"
     
       Call clsFormSkin.fn_CreateSkin(frmTest, 585, 310, AppPath & "MainMenu.bmp", RGB(255, 0, 210))
End Sub

