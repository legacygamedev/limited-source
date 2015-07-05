VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL3N.OCX"
Begin VB.Form frmEmoticonEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emoticon Editor"
   ClientHeight    =   2535
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   3255
   ControlBox      =   0   'False
   Icon            =   "frmEmoticonEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   3255
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   108
      Left            =   2760
      Top             =   2040
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2265
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   3995
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
      TabCaption(0)   =   "Edit Emoction"
      TabPicture(0)   =   "frmEmoticonEditor.frx":0FC2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblEmoticon"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Picture1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "scrlEmoticon"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtCommand"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdOk"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Command1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
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
         TabIndex        =   10
         Top             =   1920
         Width           =   1215
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
         TabIndex        =   9
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txtCommand 
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
         MaxLength       =   15
         TabIndex        =   8
         Text            =   "/"
         Top             =   1560
         Width           =   2775
      End
      Begin VB.HScrollBar scrlEmoticon 
         Height          =   255
         Left            =   120
         Max             =   1000
         TabIndex        =   3
         Top             =   360
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
         TabIndex        =   1
         Top             =   720
         Width           =   540
         Begin VB.PictureBox picEmoticon 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Left            =   15
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   4
            Top             =   15
            Width           =   480
            Begin VB.PictureBox picEmoticons 
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Command :"
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
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblEmoticon 
         AutoSize        =   -1  'True
         Caption         =   "0"
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
         Left            =   840
         TabIndex        =   5
         Top             =   720
         Width           =   75
      End
      Begin VB.Label Label5 
         Caption         =   "Emoticon :"
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
         TabIndex        =   2
         Top             =   720
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmEmoticonEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
Dim i As Long

    For i = 0 To MAX_EMOTICONS
        If Trim(Emoticons(i).Command) = Trim(txtCommand.Text) And i <> EditorIndex - 1 And Trim(txtCommand.Text) <> "/" Then
            MsgBox "There already is a " & Trim(txtCommand.Text) & " command!"
            Exit Sub
        End If
    Next i
    Call EmoticonEditorOk
End Sub

Private Sub Command1_Click()
    Call EmoticonEditorCancel
End Sub

Private Sub Form_Load()
    picEmoticons.Top = (scrlEmoticon.Value * 32) * -1
End Sub

Private Sub scrlEmoticon_Change()
    picEmoticons.Top = (scrlEmoticon.Value * 32) * -1
    lblEmoticon.Caption = scrlEmoticon.Value
End Sub

Private Sub Timer1_Timer()
    If picEmoticons.Left < -(10 * 32) Then picEmoticons.Left = 0
    picEmoticons.Left = picEmoticons.Left - 32
End Sub

Private Sub txtCommand_Change()
Dim i As String
i = txtCommand.Text
    If Mid(i, 1, 1) <> "/" Then
        If Trim(i) = "" Then
            txtCommand.Text = "/"
            Exit Sub
        End If
        txtCommand.Text = "/" & i
        txtCommand.SelStart = Len(txtCommand.Text)
    End If
End Sub
