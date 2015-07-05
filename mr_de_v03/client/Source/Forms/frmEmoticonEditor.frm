VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmEmoticonEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emoticon Editor"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4800
   ControlBox      =   0   'False
   Icon            =   "frmEmoticonEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   202
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   2865
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   5054
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   397
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Edit Emoticon"
      TabPicture(0)   =   "frmEmoticonEditor.frx":0E42
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "SSTab2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdOk"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtCommand"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
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
         TabIndex        =   3
         Text            =   "/"
         Top             =   600
         Width           =   4215
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
         Left            =   360
         TabIndex        =   2
         Top             =   2520
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
         Left            =   2880
         TabIndex        =   1
         Top             =   2520
         Width           =   1215
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   1455
         Left            =   135
         TabIndex        =   4
         Top             =   960
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   2566
         _Version        =   393216
         Tabs            =   1
         TabsPerRow      =   2
         TabHeight       =   397
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Picture"
         TabPicture(0)   =   "frmEmoticonEditor.frx":0E5E
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblEmoticon"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label5"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "scrlEmoticon"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Picture1"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).ControlCount=   4
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   540
            Left            =   120
            ScaleHeight     =   14.703
            ScaleMode       =   0  'User
            ScaleWidth      =   7.884
            TabIndex        =   6
            Top             =   360
            Width           =   540
            Begin VB.PictureBox picEmoticon 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   480
               Left            =   15
               ScaleHeight     =   16
               ScaleMode       =   0  'User
               ScaleWidth      =   10.667
               TabIndex        =   7
               Top             =   15
               Width           =   480
            End
         End
         Begin VB.HScrollBar scrlEmoticon 
            Height          =   255
            Left            =   120
            Max             =   1000
            TabIndex        =   5
            Top             =   1080
            Value           =   1
            Width           =   3975
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
            Left            =   3045
            TabIndex        =   9
            Top             =   735
            Width           =   735
         End
         Begin VB.Label lblEmoticon 
            Alignment       =   1  'Right Justify
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
            Left            =   4020
            TabIndex        =   8
            Top             =   750
            Width           =   75
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
         TabIndex        =   10
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmEmoticonEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
Dim i As Long

    For i = 1 To MAX_EMOTICONS
        If i <> EditorIndex Then
            If Trim$(Emoticons(i).Command) = Trim$(txtCommand.Text) Then
                If Trim$(txtCommand.Text) <> "/" Then
                    MsgBox "There already is a " & Trim$(txtCommand.Text) & " command!"
                    Exit Sub
                End If
            End If
        End If
    Next
    Call EmoticonEditorOk
End Sub

Private Sub Command1_Click()
    Call EmoticonEditorCancel
End Sub

Private Sub Form_Load()
    'picEmoticon.Top = (scrlEmoticon.Value * 32) * -1
    BltEmoticonEditor
End Sub

Private Sub scrlEmoticon_Change()
    'picEmoticon.Top = (scrlEmoticon.Value * 32)
    BltEmoticonEditor
    lblEmoticon.Caption = scrlEmoticon.Value
End Sub

Private Sub txtCommand_Change()
Dim i As String

    i = txtCommand.Text
    
    If Left$(i, 1) <> "/" Then
        If Trim$(i) = "" Then
            txtCommand.Text = "/"
            Exit Sub
        End If
        
        txtCommand.Text = "/" & i
        txtCommand.SelStart = Len(txtCommand.Text)
    End If
End Sub

Private Sub BltEmoticonEditor()
Dim rec As RECT
Dim drec As RECT
 
     With rec
        .Top = scrlEmoticon.Value * 32
        .Bottom = .Top + 32
        .Left = 0
        .Right = .Right + 32
     End With
     
     With drec
        .Top = 0
        .Bottom = 32
        .Left = 0
        .Right = 32
     End With
     
     DD_EmoticonSurf.BltToDC picEmoticon.hdc, rec, drec
     picEmoticon.Refresh
End Sub
