VERSION 5.00
Begin VB.Form frmEditMOTD 
   Caption         =   "Edit MOTD"
   ClientHeight    =   1455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   ScaleHeight     =   1455
   ScaleWidth      =   3975
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmmotd 
      Caption         =   "Set the MOTD"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.CommandButton Command1 
         Caption         =   "Back To Admin Panel"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   3495
      End
      Begin VB.CommandButton Command57 
         Caption         =   "Default"
         Height          =   255
         Left            =   2400
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton Command67 
         Caption         =   "Save"
         Height          =   255
         Left            =   1320
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   3495
      End
      Begin VB.CommandButton Command68 
         Caption         =   "Clear"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmEditMOTD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmEditMOTD.Visible = False
frmadmin.Visible = True
End Sub

Private Sub Command67_Click()
Dim MOTD As String
MOTD = Text2.Text
Call SendData("SETMOTD" & SEP_CHAR & MOTD & SEP_CHAR & END_CHAR)
End Sub

Private Sub Command57_Click()
Dim MOTD As String
Text2.Text = "Welcome to my game made with Bevrias Engine, www.Bevrias-Engine.tk."
MOTD = "Welcome to my game made with Bevrias Engine, www.Bevrias-Engine.tk."
Call SendData("SETMOTD" & SEP_CHAR & MOTD & SEP_CHAR & END_CHAR)
End Sub


Private Sub Command68_Click()
Dim MOTD As String
Text2.Text = " "
MOTD = " "
Call SendData("SETMOTD" & SEP_CHAR & MOTD & SEP_CHAR & END_CHAR)
End Sub
