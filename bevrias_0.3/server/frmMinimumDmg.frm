VERSION 5.00
Begin VB.Form frmMinimumDmg 
   Caption         =   "Minimum Damage"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   1080
      ScaleHeight     =   2025
      ScaleWidth      =   2385
      TabIndex        =   13
      Top             =   480
      Width           =   2415
      Visible         =   0   'False
      Begin VB.CommandButton Command10 
         Caption         =   "Close"
         Height          =   255
         Left            =   1320
         TabIndex        =   20
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Maximum = 10.000"
         Height          =   255
         Left            =   840
         TabIndex        =   19
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Minimum = 0"
         Height          =   255
         Left            =   840
         TabIndex        =   18
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Player Minimum Damage:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Maximum = 10.000"
         Height          =   255
         Left            =   840
         TabIndex        =   16
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Minimum = 0"
         Height          =   255
         Left            =   840
         TabIndex        =   15
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Enemy Minimum Damage:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   2535
      End
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Close"
      Height          =   255
      Left            =   3720
      TabIndex        =   12
      Top             =   2880
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "Player Minimum Damage"
      Height          =   1335
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   4455
      Begin VB.CommandButton Command8 
         Caption         =   "Help"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Default"
         Height          =   255
         Left            =   1560
         TabIndex        =   10
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Clear"
         Height          =   255
         Left            =   840
         TabIndex        =   9
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Save"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Enemy Minimum Damage"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton Command4 
         Caption         =   "Help"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Default"
         Height          =   255
         Left            =   1560
         TabIndex        =   4
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Clear"
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4215
      End
   End
End
Attribute VB_Name = "frmMinimumDmg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim minimum As String
minimum = Text1.text
Call PutVar(App.Path & "\Data.ini", "ADDED", "EnemyMinimumDmg", minimum)
End Sub

Private Sub Command10_Click()
Picture1.Visible = False
End Sub

Private Sub Command2_Click()
Text1.text = ""
End Sub

Private Sub Command3_Click()
Text1.text = "1"
End Sub

Private Sub Command4_Click()
Picture1.Visible = True
End Sub

Private Sub Command5_Click()
Dim minimum As String
minimum = Text2.text
Call PutVar(App.Path & "\Data.ini", "ADDED", "PlayerMinimumDmg", minimum)
End Sub

Private Sub Command6_Click()
Text2.text = ""
End Sub

Private Sub Command7_Click()
Text2.text = "1"
End Sub

Private Sub Command8_Click()
Picture1.Visible = True
End Sub

Private Sub Command9_Click()
frmMinimumDmg.Visible = False
End Sub

Private Sub Form_Load()
Text1.text = GetVar(App.Path & "\Data.ini", "ADDED", "EnemyMinimumDmg")
Text2.text = GetVar(App.Path & "\Data.ini", "ADDED", "PlayerMinimumDmg")
End Sub
