VERSION 5.00
Begin VB.Form frmExpToAttacker 
   Caption         =   "Experience to Attacker"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   0
      ScaleHeight     =   3105
      ScaleWidth      =   4665
      TabIndex        =   13
      Top             =   0
      Width           =   4695
      Visible         =   0   'False
      Begin VB.CommandButton Command6 
         Caption         =   "Close"
         Height          =   255
         Left            =   3840
         TabIndex        =   14
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Experience to Enemy = This is how much experience you will lose when you are killed by an enemy, not a player."
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   2520
         Width           =   4455
      End
      Begin VB.Label Label2 
         Caption         =   $"frmExpToAttacker.frx":0000
         Height          =   855
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   4455
      End
      Begin VB.Label Label1 
         Caption         =   $"frmExpToAttacker.frx":00CE
         Height          =   1575
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   4455
      End
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Close"
      Height          =   255
      Left            =   4080
      TabIndex        =   12
      Top             =   2880
      Width           =   615
   End
   Begin VB.Frame Frame2 
      Caption         =   "Experience to Enemy"
      Height          =   1335
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   4455
      Begin VB.CommandButton Command10 
         Caption         =   "Help"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Clear"
         Height          =   255
         Left            =   1680
         TabIndex        =   10
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Default"
         Height          =   255
         Left            =   840
         TabIndex        =   9
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton Command7 
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
      Caption         =   "Experience to Attacker"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton Command4 
         Caption         =   "Help"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Clear"
         Height          =   255
         Left            =   1800
         TabIndex        =   4
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Default"
         Height          =   255
         Left            =   960
         TabIndex        =   3
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox eText1 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4215
      End
   End
End
Attribute VB_Name = "frmExpToAttacker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim attacker As String
attacker = eText1.text
If IsNumeric(attacker) = True Then
Call PutVar(App.Path & "\Data.ini", "ADDED", "ExpToAttacker", attacker)
Else
Call PutVar(App.Path & "\Data.ini", "ADDED", "ExpToAttacker", "10")
End If
End Sub

Private Sub Command10_Click()
Picture1.Visible = True
End Sub

Private Sub Command11_Click()
frmExpToAttacker.Visible = False
End Sub

Private Sub Command2_Click()
eText1.text = "12"
End Sub

Private Sub Command3_Click()
eText1.text = ""
End Sub

Private Sub Command4_Click()
Picture1.Visible = True
End Sub

Private Sub Command5_Click()
frmExpToAttacker.Visible = False
End Sub

Private Sub Command6_Click()
Picture1.Visible = False
End Sub

Private Sub Command7_Click()
Dim attacker As String
attacker = Text2.text
Call PutVar(App.Path & "\Data.ini", "ADDED", "ExpLostToEnemy", attacker)
End Sub

Private Sub Command8_Click()
Text2.text = "5"
End Sub

Private Sub Command9_Click()
Text2.text = ""
End Sub

Private Sub Form_Load()
eText1.text = GetVar(App.Path & "\Data.ini", "ADDED", "ExpToAttacker")
Text2.text = GetVar(App.Path & "\Data.ini", "ADDED", "ExpLostToEnemy")
End Sub
