VERSION 5.00
Begin VB.Form frmCredits 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Engine Credits"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   -45
   ClientWidth     =   6000
   ControlBox      =   0   'False
   Icon            =   "frmCredits.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCredits.frx":0FC2
   ScaleHeight     =   6000
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      FillColor       =   &H0000FF00&
      ForeColor       =   &H0000FF00&
      Height          =   1455
      Left            =   960
      ScaleHeight     =   1425
      ScaleWidth      =   4065
      TabIndex        =   0
      Top             =   1440
      Width           =   4095
      Begin VB.Label lblCredit 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.ariae.co.nr/"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   12
         Top             =   1110
         Width           =   3855
      End
      Begin VB.Label lblCredit 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Draken"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   10
         Top             =   840
         Width           =   1575
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000FF00&
         Height          =   1425
         Left            =   0
         Top             =   0
         Width           =   4065
      End
      Begin VB.Label lblCredit 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Baron"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   8
         Left            =   3000
         TabIndex        =   9
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblCredit 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Pingu"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   7
         Left            =   2160
         TabIndex        =   8
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblCredit 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "GodSentDeath"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   6
         Left            =   2160
         TabIndex        =   7
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblCredit 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Unreal"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   5
         Left            =   3000
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblCredit 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Coke"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   4
         Left            =   2160
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblCredit 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "- Special Thanks To -"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   3
         Left            =   2160
         TabIndex        =   4
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label lblCredit 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Topher"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblCredit 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Emblem"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblCredit 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "- Engine Developers -"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   1815
      End
   End
   Begin VB.Label picCancel 
      BackStyle       =   0  'Transparent
      Caption         =   "Back To Main Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   11
      Top             =   5400
      Width           =   1320
   End
End
Attribute VB_Name = "frmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim I As Long
Dim Ending As String
    For I = 1 To 3
        If I = 1 Then Ending = ".gif"
        If I = 2 Then Ending = ".jpg"
        If I = 3 Then Ending = ".png"

        If FileExist("GUI\Credits" & Ending) Then frmCredits.Picture = LoadPicture(App.Path & "\GUI\Credits" & Ending)
    Next I
End Sub

Private Sub lblCredit_Click(Index As Integer)
    Select Case Index
        Case 1
            Shell "explorer ""http://ariae.co.nr/""", vbNormalFocus
        Case 2
            Shell "explorer ""http://united-fiction.co.nr/""", vbNormalFocus
        Case 9
            MsgBox "GNOME DOMINATION!!!!", vbOKOnly Or vbInformation
    End Select
End Sub

Private Sub picCancel_Click()
    frmMainMenu.Visible = True
    frmCredits.Visible = False
End Sub

