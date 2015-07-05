VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmEmoticonEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emoticon Editor"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3075
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   3075
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   108
      Left            =   2880
      Top             =   0
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4425
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   7805
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
      TabCaption(0)   =   "Edit Emoticon"
      TabPicture(0)   =   "frmEmoticonEditor.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "SSTab2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtCommand"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdOk"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
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
         Left            =   1440
         TabIndex        =   4
         Top             =   3960
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
         Left            =   120
         TabIndex        =   3
         Top             =   3960
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
         TabIndex        =   2
         Text            =   "/"
         Top             =   600
         Width           =   2535
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   2655
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   4683
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
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
         TabPicture(0)   =   "frmEmoticonEditor.frx":001C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label5"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lblEmoticon"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "chkPic"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Picture1"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "scrlEmoticon"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).ControlCount=   5
         TabCaption(1)   =   "Sound"
         TabPicture(1)   =   "frmEmoticonEditor.frx":0038
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "chkSound"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "lstSound"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).ControlCount=   2
         Begin VB.ListBox lstSound 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1815
            Left            =   -74880
            TabIndex        =   14
            Top             =   720
            Width           =   2055
         End
         Begin VB.CheckBox chkSound 
            Caption         =   "Use sound?"
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
            Left            =   -74400
            TabIndex        =   11
            Top             =   360
            Width           =   1095
         End
         Begin VB.HScrollBar scrlEmoticon 
            Height          =   255
            Left            =   120
            Max             =   1000
            TabIndex        =   10
            Top             =   2160
            Value           =   1
            Width           =   2055
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   540
            Left            =   840
            ScaleHeight     =   34
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   34
            TabIndex        =   7
            Top             =   1320
            Width           =   540
            Begin VB.PictureBox picEmoticon 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   480
               Left            =   15
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   8
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
                  TabIndex        =   9
                  Top             =   0
                  Width           =   480
               End
            End
         End
         Begin VB.CheckBox chkPic 
            Caption         =   "Use picture?"
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
            Left            =   600
            TabIndex        =   6
            Top             =   360
            Width           =   1095
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
            Left            =   1080
            TabIndex        =   13
            Top             =   840
            Width           =   315
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
            Left            =   240
            TabIndex        =   12
            Top             =   840
            Width           =   735
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
         TabIndex        =   1
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

' Copyright (c) 2007-2008 Elysium Source. All rights reserved.
' This code is licensed under the Elysium General License.

Private Sub cmdOk_Click()
Dim I As Long

    For I = 0 To MAX_EMOTICONS
        If Trim$(Emoticons(I).Command) = Trim$(txtCommand.Text) And I <> EditorIndex - 1 And Trim$(txtCommand.Text) <> "/" Then
            MsgBox "There already is a " & Trim$(txtCommand.Text) & " command!"
            Exit Sub
        End If
    Next I
    If chkSound.Value = 0 And chkPic.Value = 0 Then
        MsgBox "You need to select a picture or a sound!"
        Exit Sub
    End If
    Call StopSound
    Call EmoticonEditorOk
End Sub

Private Sub Command1_Click()
    Call StopSound
    Call EmoticonEditorCancel
End Sub

Private Sub Form_Load()
    picEmoticons.Top = (scrlEmoticon.Value * 32) * -1
    Call ListSounds(App.Path & "\SFX\", 3)
    lstSound.Text = Emoticons(EditorIndex - 1).Sound
End Sub

Private Sub lstSound_Click()
    If chkSound.Value = 1 Then Call PlaySound(lstSound.Text)
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
Dim I As String
I = txtCommand.Text
    If Mid(I, 1, 1) <> "/" Then
        If Trim$(I) = vbNullString Then
            txtCommand.Text = "/"
            Exit Sub
        End If
        txtCommand.Text = "/" & I
        txtCommand.SelStart = Len(txtCommand.Text)
    End If
End Sub
