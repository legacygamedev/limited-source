VERSION 5.00
Begin VB.Form frmPKLevel 
   Caption         =   "PK Level"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   3375
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame13 
      Caption         =   "PK Level Information"
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   2880
      Width           =   3135
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "-Minimum Value: 1 -Maximum Value: 10000"
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "PK Level"
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.CommandButton Command92 
         Caption         =   "Clear"
         Height          =   255
         Left            =   1200
         TabIndex        =   13
         Top             =   1800
         Width           =   495
      End
      Begin VB.CommandButton Command89 
         Caption         =   "Default"
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   1800
         Width           =   735
      End
      Begin VB.CommandButton Command88 
         Caption         =   "Set"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1800
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Close"
         Height          =   255
         Left            =   2040
         TabIndex        =   8
         Top             =   2400
         Width           =   975
      End
      Begin VB.CommandButton Command91 
         Caption         =   "Clear"
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton Command90 
         Caption         =   "Default"
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command69 
         Caption         =   "Set"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox txtPKLEVEL 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   2895
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label Label10 
         Caption         =   "Be PKED Level:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label11 
         Caption         =   "PK Other Level:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmPKLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmPKLevel.Visible = False
End Sub

Private Sub Form_Load()
txtPKLEVEL.text = GetVar(App.Path & "\Data.ini", "ADDED", "PKLEVEL")
Text6.text = GetVar(App.Path & "\Data.ini", "ADDED", "BEPKED")
End Sub
Private Sub Command69_Click()
Dim pklevell As String
pklevell = txtPKLEVEL.text
If IsNumeric(pklevell) = True Then
Call PutVar(App.Path & "\Data.ini", "ADDED", "PKLEVEL", pklevell)
Else
Call PutVar(App.Path & "\Data.ini", "ADDED", "PKLEVEL", "10")
End If
End Sub
Private Sub Command88_Click()
Dim pklevell As String
pklevell = Text6.text
If IsNumeric(pklevell) = True Then
Call PutVar(App.Path & "\Data.ini", "ADDED", "BEPKED", pklevell)
Else
Call PutVar(App.Path & "\Data.ini", "ADDED", "BEPKED", "10")
End If
End Sub

Private Sub Command89_Click()
Text6.text = "10"
Call PutVar(App.Path & "\Data.ini", "ADDED", "BEPKED", 10)
End Sub

Private Sub Command90_Click()
Dim ss
ss = 10
txtPKLEVEL.text = ss
Call PutVar(App.Path & "\Data.ini", "ADDED", "PKLEVEL", 10)
End Sub

Private Sub Command91_Click()
txtPKLEVEL.text = ""
Call PutVar(App.Path & "\Data.ini", "CONFIG", "ADDED", "10")
End Sub

Private Sub Command92_Click()
Text6.text = ""
Call PutVar(App.Path & "\Data.ini", "ADDED", "BEPKED", "10")
End Sub
