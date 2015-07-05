VERSION 5.00
Begin VB.Form frmPartyRangeLevel 
   Caption         =   "Party Range Level"
   ClientHeight    =   1590
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2775
   LinkTopic       =   "Form1"
   ScaleHeight     =   1590
   ScaleWidth      =   2775
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   1080
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Party Range"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      Begin VB.CommandButton Command3 
         Caption         =   "Default"
         Height          =   255
         Left            =   1200
         TabIndex        =   3
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmPartyRangeLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmPartyRangeLevel.Visible = False
End Sub

Private Sub Command2_Click()
Call PutVar(App.Path & "\Data.ini", "ADDED", "PARTYRANGELVL", Text1.text)
End Sub

Private Sub Command3_Click()
Call PutVar(App.Path & "\Data.ini", "ADDED", "PARTYRANGELVL", 5)
Text1.text = "5"
End Sub

Private Sub Form_Load()
Text1.text = GetVar(App.Path & "\Data.ini", "ADDED", "PARTYRANGELVL")
End Sub
