VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Regen"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   3495
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command15 
      Caption         =   "Help"
      Height          =   255
      Left            =   2640
      TabIndex        =   33
      Top             =   3840
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   120
      ScaleHeight     =   1905
      ScaleWidth      =   3225
      TabIndex        =   25
      Top             =   840
      Width           =   3255
      Visible         =   0   'False
      Begin VB.CommandButton Command14 
         Caption         =   "Close"
         Height          =   255
         Left            =   2400
         TabIndex        =   32
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Maximum = 100"
         Height          =   255
         Left            =   720
         TabIndex        =   31
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Minimum = 0"
         Height          =   255
         Left            =   720
         TabIndex        =   30
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "RegenSpeed:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "1 = On"
         Height          =   255
         Left            =   720
         TabIndex        =   28
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "0 = Off"
         Height          =   255
         Left            =   720
         TabIndex        =   27
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Regen:"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Close"
      Height          =   255
      Left            =   1920
      TabIndex        =   24
      Top             =   3840
      Width           =   735
   End
   Begin VB.Frame Frame6 
      Caption         =   "SP Regen Speed"
      Height          =   615
      Left            =   120
      TabIndex        =   20
      Top             =   3120
      Width           =   3255
      Begin VB.CommandButton Command12 
         Caption         =   "Default"
         Height          =   255
         Left            =   2280
         TabIndex        =   23
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Save"
         Height          =   255
         Left            =   1560
         TabIndex        =   22
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "SP Regen"
      Height          =   615
      Left            =   120
      TabIndex        =   16
      Top             =   2520
      Width           =   3255
      Begin VB.CommandButton Command10 
         Caption         =   "Default"
         Height          =   255
         Left            =   2280
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Save"
         Height          =   255
         Left            =   1560
         TabIndex        =   18
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "MP Regen Speed"
      Height          =   615
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   3255
      Begin VB.CommandButton Command8 
         Caption         =   "Default"
         Height          =   255
         Left            =   2280
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Save"
         Height          =   255
         Left            =   1560
         TabIndex        =   14
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "MP Regen"
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   3255
      Begin VB.CommandButton Command6 
         Caption         =   "Default"
         Height          =   255
         Left            =   2280
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Save"
         Height          =   255
         Left            =   1560
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "HP Regen Speed"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   3255
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Default"
         Height          =   255
         Left            =   2280
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Save"
         Height          =   255
         Left            =   1560
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "HP Regen"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      Begin VB.CommandButton Command2 
         Caption         =   "Default"
         Height          =   255
         Left            =   2280
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save"
         Height          =   255
         Left            =   1560
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim regen As String
regen = Text1.text
Call PutVar(App.Path & "\Data.ini", "CONFIG", "HPRegen", regen)
End Sub

Private Sub Command10_Click()
Text5.text = "1"
End Sub

Private Sub Command11_Click()
Dim regen As String
regen = Text2.text
Call PutVar(App.Path & "\Data.ini", "CONFIG", "HPRegenSpeed", regen)
End Sub

Private Sub Command12_Click()
Text6.text = "2"
End Sub

Private Sub Command13_Click()
frmOptions.Visible = False
End Sub

Private Sub Command14_Click()
Picture1.Visible = False
End Sub

Private Sub Command15_Click()
Picture1.Visible = True
End Sub

Private Sub Command2_Click()
Text1.text = "1"
End Sub

Private Sub Command3_Click()
Dim regen As String
regen = Text3.text
Call PutVar(App.Path & "\Data.ini", "CONFIG", "MPRegen", regen)
End Sub

Private Sub Command4_Click()
Text2.text = "1"
End Sub

Private Sub Command5_Click()
Dim regen As String
regen = Text4.text
Call PutVar(App.Path & "\Data.ini", "CONFIG", "MPRegenSpeed", regen)
End Sub

Private Sub Command6_Click()
Text3.text = "1"
End Sub

Private Sub Command7_Click()
Dim regen As String
regen = Text5.text
Call PutVar(App.Path & "\Data.ini", "CONFIG", "SPRegen", regen)
End Sub

Private Sub Command8_Click()
Text4.text = "1"
End Sub

Private Sub Command9_Click()
Dim regen As String
regen = Text6.text
Call PutVar(App.Path & "\Data.ini", "CONFIG", "SPRegenSpeed", regen)
End Sub

Private Sub Form_Load()
Text1.text = GetVar(App.Path & "\Data.ini", "CONFIG", "HPRegen")
Text2.text = GetVar(App.Path & "\Data.ini", "CONFIG", "HPRegenSpeed")
Text3.text = GetVar(App.Path & "\Data.ini", "CONFIG", "MPRegen")
Text4.text = GetVar(App.Path & "\Data.ini", "CONFIG", "MPRegenSpeed")
Text5.text = GetVar(App.Path & "\Data.ini", "CONFIG", "SPRegen")
Text6.text = GetVar(App.Path & "\Data.ini", "CONFIG", "SPRegenSpeed")
End Sub

