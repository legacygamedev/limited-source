VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmNews 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit News"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3240
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Colour"
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label lblNews 
      Caption         =   "News:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4455
   End
   Begin VB.Label lblCaption 
      Caption         =   "Caption:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmNews"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Red
Public Green
Public Blue

Private Sub btnOK_Click()
Call PutVar(App.Path & "\News.ini", "Data", "ServerNews", Text1.text)
Call PutVar(App.Path & "\News.ini", "Data", "Desc", Text2.text)
Unload Me
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
frmColor.Visible = True
End Sub

Private Sub Form_Load()
Text1.text = GetVar(App.Path & "\News.ini", "Data", "ServerNews")
Text2.text = GetVar(App.Path & "\News.ini", "Data", "Desc")
Red = GetVar(App.Path & "\News.ini", "Color", "Red")
Green = GetVar(App.Path & "\News.ini", "Color", "Green")
Blue = GetVar(App.Path & "\News.ini", "Color", "Blue")
Text1.ForeColor = RGB(Red, Green, Blue)
Text2.ForeColor = RGB(Red, Green, Blue)
Text1.ForeColor = RGB(Red, Green, Blue)
Text2.ForeColor = RGB(Red, Green, Blue)
End Sub
