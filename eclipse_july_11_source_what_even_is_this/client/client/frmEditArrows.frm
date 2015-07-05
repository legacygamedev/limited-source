VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL3N.OCX"
Begin VB.Form frmEditArrows 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Arrow Editor"
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   3270
   ControlBox      =   0   'False
   Icon            =   "frmEditArrows.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   3270
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   3555
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   6271
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   397
      TabMaxWidth     =   1984
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Edit Arrow"
      TabPicture(0)   =   "frmEditArrows.frx":0FC2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblRange"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblArrow"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblAmount"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdOk"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "scrlArrow"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Picture1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "scrlRange"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtName"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "scrlAmount"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      Begin VB.HScrollBar scrlAmount 
         Height          =   255
         Left            =   120
         Max             =   500
         Min             =   1
         TabIndex        =   12
         Top             =   2760
         Value           =   1
         Width           =   2775
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   2775
      End
      Begin VB.HScrollBar scrlRange 
         Height          =   255
         Left            =   120
         Max             =   30
         Min             =   1
         TabIndex        =   9
         Top             =   2160
         Value           =   1
         Width           =   2775
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   2400
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   34
         TabIndex        =   4
         Top             =   1560
         Width           =   540
         Begin VB.PictureBox picEmoticon 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Left            =   15
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   5
            Top             =   15
            Width           =   480
            Begin VB.PictureBox picArrows 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   480
               Left            =   0
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   6
               Top             =   0
               Width           =   480
            End
         End
      End
      Begin VB.HScrollBar scrlArrow 
         Height          =   255
         Left            =   120
         Max             =   500
         Min             =   1
         TabIndex        =   3
         Top             =   1200
         Value           =   1
         Width           =   2775
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "Ok"
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
         Left            =   240
         TabIndex        =   2
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cancel"
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
         Left            =   1560
         TabIndex        =   1
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label lblAmount 
         Caption         =   "Amount:"
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
         Left            =   120
         TabIndex        =   14
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   420
      End
      Begin VB.Label lblArrow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Arrow:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   435
      End
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         Caption         =   "Range:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   120
         TabIndex        =   7
         Top             =   1920
         Width           =   435
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   960
      TabIndex        =   13
      Top             =   1800
      Width           =   45
   End
End
Attribute VB_Name = "frmEditArrows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
    Call ArrowEditorOk
End Sub

Private Sub Command1_Click()
    Call ArrowEditorCancel
End Sub

Private Sub scrlArrow_Change()
    lblArrow.Caption = "Arrow: " & scrlArrow.Value
    picArrows.Top = (scrlArrow.Value * 32) * -1
End Sub

Private Sub scrlRange_Change()
    lblRange.Caption = "Range: " & scrlRange.Value
End Sub
Private Sub scrlAmount_Change()
    lblAmount.Caption = "Amount: " & scrlAmount
End Sub
