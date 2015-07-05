VERSION 5.00
Begin VB.Form frmMOTD 
   Caption         =   "MOTD"
   ClientHeight    =   1575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   ScaleHeight     =   1575
   ScaleWidth      =   4815
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame ww 
      Caption         =   "Set The Welcome Message"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.CommandButton Command1 
         Caption         =   "Close"
         Height          =   255
         Left            =   3720
         TabIndex        =   6
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton Command57 
         Caption         =   "Default"
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton Command67 
         Caption         =   "Save"
         Height          =   255
         Left            =   1080
         TabIndex        =   4
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton Command68 
         Caption         =   "Clear"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton Command94 
         Caption         =   "MOTD"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4335
      End
   End
End
Attribute VB_Name = "frmMOTD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
frmmotd.Visible = False
End Sub

Private Sub Form_Load()
Text2.text = GetVar(App.Path & "\motd.ini", "MOTD", "Msg")
End Sub
Private Sub Command57_Click()
Dim NewMOTD As String
Text2.text = "Welcome to my game made with Bevrias Engine, www.Bevrias-Engine.tk."
NewMOTD = "Welcome to my game made with Bevrias Engine, www.Bevrias-Engine.tk."
Call PutVar(App.Path & "\motd.ini", "MOTD", "Msg", NewMOTD)
End Sub

Private Sub Command67_Click()
Dim NewMOTD As String
NewMOTD = Text2.text
Call PutVar(App.Path & "\motd.ini", "MOTD", "Msg", NewMOTD)
End Sub

Private Sub Command68_Click()
Dim NewMOTD As String
Text2.text = " "
NewMOTD = " "
Call PutVar(App.Path & "\motd.ini", "MOTD", "Msg", NewMOTD)
End Sub

Private Sub Command94_Click()
    AFileName = "MOTD.ini"
    Unload frmEditor
    frmEditor.Show
End Sub
