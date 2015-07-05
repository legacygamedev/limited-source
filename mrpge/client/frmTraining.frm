VERSION 5.00
Begin VB.Form frmTraining 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "M:RPGe (Training)"
   ClientHeight    =   6540
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   7680
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTraining.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   7680
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbStat 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      ItemData        =   "frmTraining.frx":0442
      Left            =   2160
      List            =   "frmTraining.frx":0458
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   3120
      Width           =   4575
   End
   Begin VB.Image Image4 
      Height          =   195
      Left            =   7080
      Picture         =   "frmTraining.frx":049F
      Top             =   120
      Width           =   195
   End
   Begin VB.Image Image3 
      Height          =   195
      Left            =   7320
      Picture         =   "frmTraining.frx":2612
      Top             =   120
      Width           =   195
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Train:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   2760
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   0
      Picture         =   "frmTraining.frx":47B8
      Top             =   0
      Width           =   7680
   End
   Begin VB.Image picTrain 
      Height          =   480
      Left            =   480
      Picture         =   "frmTraining.frx":7387
      Top             =   2760
      Width           =   1320
   End
   Begin VB.Image picCancel 
      Height          =   480
      Left            =   480
      Picture         =   "frmTraining.frx":7833
      Top             =   3360
      Width           =   1320
   End
   Begin VB.Image Image2 
      Height          =   6195
      Left            =   0
      Picture         =   "frmTraining.frx":7D13
      Top             =   360
      Width           =   7680
   End
End
Attribute VB_Name = "frmTraining"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    cmbStat.ListIndex = 0
End Sub

Private Sub Image3_Click()
Unload Me
End Sub

Private Sub Image4_Click()
frmTraining.WindowState = vbMinimized
End Sub

Private Sub picTrain_Click()
    Call SendData("usestatpoint" & SEP_CHAR & cmbStat.ListIndex & SEP_CHAR & END_CHAR)
End Sub

Private Sub picCancel_Click()
    Unload Me
End Sub

