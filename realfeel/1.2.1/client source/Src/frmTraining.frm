VERSION 5.00
Begin VB.Form frmTraining 
   BackColor       =   &H00004080&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2520
   ClientLeft      =   15
   ClientTop       =   -90
   ClientWidth     =   2970
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTraining.frx":0000
   ScaleHeight     =   168
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   198
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picHeaderTrain 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   -10
      Picture         =   "frmTraining.frx":1C49A
      ScaleHeight     =   780
      ScaleWidth      =   3000
      TabIndex        =   3
      Top             =   -10
      Width           =   3000
   End
   Begin VB.PictureBox picCancel 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1485
      Picture         =   "frmTraining.frx":23EBC
      ScaleHeight     =   375
      ScaleWidth      =   1500
      TabIndex        =   2
      Top             =   2160
      Width           =   1500
   End
   Begin VB.PictureBox picTrain 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -20
      Picture         =   "frmTraining.frx":25C4A
      ScaleHeight     =   375
      ScaleWidth      =   1500
      TabIndex        =   1
      Top             =   2160
      Width           =   1500
   End
   Begin VB.ComboBox cmbStat 
      BackColor       =   &H00F5763F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0009E7F2&
      Height          =   405
      ItemData        =   "frmTraining.frx":279D8
      Left            =   120
      List            =   "frmTraining.frx":279E8
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1440
      Width           =   2655
   End
End
Attribute VB_Name = "frmTraining"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub picCancel_Click()
    Unload Me
End Sub

Private Sub picTrain_Click()
    Call SendData("usestatpoint" & SEP_CHAR & cmbStat.ListIndex & SEP_CHAR & END_CHAR)
End Sub

Private Sub Form_Load()
    cmbStat.ListIndex = 0
End Sub

