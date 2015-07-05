VERSION 5.00
Begin VB.Form frmStats 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8580
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1800
   ScaleWidth      =   8580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   240
      Top             =   5400
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Alignment"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   3480
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label mid 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Neutral"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   600
      Width           =   6975
   End
   Begin VB.Label lblRight 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Holy"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6840
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lblLeft 
      BackStyle       =   0  'Transparent
      Caption         =   "Evil"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.Image gball 
      Height          =   210
      Left            =   7920
      Picture         =   "frmStats.frx":0000
      Top             =   1080
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image eball 
      Height          =   210
      Left            =   600
      Picture         =   "frmStats.frx":02AA
      Top             =   1080
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image bar 
      Height          =   195
      Left            =   840
      Picture         =   "frmStats.frx":0554
      Top             =   1080
      Width           =   7035
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Close"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7680
      TabIndex        =   0
      Top             =   1560
      Width           =   495
   End
End
Attribute VB_Name = "frmStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim P As Long

Private Sub Form_Load()
    Left = frmMirage.Left
    Top = frmMirage.Top

    P = 5000
    
    If P >= 5000 Then
        gball.Visible = True
        eball.Visible = False
        gball.Left = bar.Left + (((bar.Width - gball.Width) / 10000) * P)
        gball.Top = bar.Top - 15
    Else
        gball.Visible = False
        eball.Visible = True
        eball.Left = bar.Left + (((bar.Width - eball.Width) / 10000) * P)
        eball.Top = bar.Top - 15
    End If
    
End Sub

Private Sub Label1_Click()
    frmStats.Hide
End Sub

Private Sub Timer1_Timer()

    P = Player(MyIndex).Alignment
    
    If P >= 5000 Then
        gball.Visible = True
        eball.Visible = False
        gball.Left = bar.Left + (((bar.Width - gball.Width) / 10000) * P)
        gball.Top = bar.Top - 15
    Else
        gball.Visible = False
        eball.Visible = True
        eball.Left = bar.Left + (((bar.Width - eball.Width) / 10000) * P)
        eball.Top = bar.Top - 15
    End If
    
End Sub
