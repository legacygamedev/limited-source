VERSION 5.00
Begin VB.Form frmEmoticonEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emoticon Editor"
   ClientHeight    =   2550
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   3255
   ControlBox      =   0   'False
   Icon            =   "frmEmoticonEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   3255
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Edit Emoticon"
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   2400
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   34
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   840
         Width           =   540
         Begin VB.PictureBox picEmoticon 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Left            =   15
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   9
            TabStop         =   0   'False
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
               TabIndex        =   10
               TabStop         =   0   'False
               Top             =   0
               Width           =   480
            End
         End
      End
      Begin VB.HScrollBar scrlEmoticon 
         Height          =   255
         Left            =   240
         Max             =   1000
         TabIndex        =   4
         Top             =   480
         Value           =   1
         Width           =   2775
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
         Left            =   240
         MaxLength       =   15
         TabIndex        =   3
         Text            =   "/"
         Top             =   1680
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
         Top             =   2040
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
         Left            =   1800
         TabIndex        =   1
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         TabIndex        =   7
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblEmoticon 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   960
         TabIndex        =   6
         Top             =   840
         Width           =   75
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   735
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   108
      Left            =   3120
      Top             =   0
   End
End
Attribute VB_Name = "frmEmoticonEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdOk_Click()
    Dim I As Long

    For I = 0 To MAX_EMOTICONS
        If Trim$(Emoticons(I).Command) = Trim$(txtCommand.Text) And I <> EditorIndex - 1 And Trim$(txtCommand.Text) <> "/" Then
            MsgBox "There already is a " & Trim$(txtCommand.Text) & " command!"
            Exit Sub
        End If
    Next I
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
    If picEmoticons.Left < -(10 * 32) Then
        picEmoticons.Left = 0
    End If
    picEmoticons.Left = picEmoticons.Left - 32
End Sub

Private Sub txtCommand_Change()
    Dim I As String
    I = txtCommand.Text
    If Mid$(I, 1, 1) <> "/" Then
        If Trim$(I) = vbNullString Then
            txtCommand.Text = "/"
            Exit Sub
        End If
        txtCommand.Text = "/" & I
        txtCommand.SelStart = Len(txtCommand.Text)
    End If
End Sub
