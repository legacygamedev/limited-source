VERSION 5.00
Begin VB.Form frmNews 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "News Editor"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdColor 
      Caption         =   "Change Color"
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
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
      Caption         =   "News Content:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblCaption 
      Caption         =   "News Title:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmNews"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public RED
Public GREEN
Public BLUE

Private Sub cmdOK_Click()
    Call PutVar(App.Path & "\News.ini", "Data", "NewsTitle", Text1.Text)
    Call PutVar(App.Path & "\News.ini", "Data", "NewsBody", Text2.Text)

    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdColor_Click()
    frmColor.Visible = True
End Sub

Private Sub Form_Load()
    Text1.Text = GetVar(App.Path & "\News.ini", "Data", "NewsTitle")
    Text2.Text = GetVar(App.Path & "\News.ini", "Data", "NewsBody")

    RED = GetVar(App.Path & "\News.ini", "Color", "Red")
    GREEN = GetVar(App.Path & "\News.ini", "Color", "Green")
    BLUE = GetVar(App.Path & "\News.ini", "Color", "Blue")

    Text1.ForeColor = RGB(RED, GREEN, BLUE)
    Text2.ForeColor = RGB(RED, GREEN, BLUE)
End Sub
