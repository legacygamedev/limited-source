VERSION 5.00
Begin VB.Form frmSave 
   Caption         =   "Save"
   ClientHeight    =   2775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2055
   LinkTopic       =   "Form1"
   ScaleHeight     =   2775
   ScaleWidth      =   2055
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Objects"
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
      Begin VB.CommandButton Command8 
         Caption         =   "Close"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Save Arrows"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Save Logs"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Save Maps"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Save NPCS"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Save Spells"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save Shops"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton S 
         Caption         =   "Save Items"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save Classes"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call SaveClasses
End Sub

Private Sub Command2_Click()
Call SaveShops
End Sub

Private Sub Command3_Click()
Call SaveSpells
End Sub

Private Sub Command4_Click()
Call SaveNpcs
End Sub

Private Sub Command5_Click()
Call CheckMaps
End Sub

Private Sub Command6_Click()
Call SaveLogs
End Sub

Private Sub Command7_Click()
Call CheckArrows
End Sub

Private Sub Command8_Click()
frmSave.Visible = False
End Sub

Private Sub S_Click()
Call SaveItems
End Sub
