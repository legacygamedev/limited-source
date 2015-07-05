VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmEditor 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "FriendCodes Pocket Entertainment System"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   8685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picQuote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7560
      Picture         =   "frmEditor.frx":0000
      ScaleHeight     =   21.739
      ScaleMode       =   0  'User
      ScaleWidth      =   65.217
      TabIndex        =   19
      Top             =   1080
      Width           =   1125
   End
   Begin VB.PictureBox picPreview 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7560
      Picture         =   "frmEditor.frx":1686
      ScaleHeight     =   21.739
      ScaleMode       =   0  'User
      ScaleWidth      =   65.217
      TabIndex        =   15
      Top             =   3720
      Width           =   1125
   End
   Begin VB.PictureBox picText 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7560
      Picture         =   "frmEditor.frx":2D0C
      ScaleHeight     =   21.739
      ScaleMode       =   0  'User
      ScaleWidth      =   65.217
      TabIndex        =   18
      Top             =   3720
      Width           =   1125
   End
   Begin RichTextLib.RichTextBox rtbNews 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   7858
      _Version        =   393217
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmEditor.frx":4392
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picCode 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7560
      Picture         =   "frmEditor.frx":4416
      ScaleHeight     =   21.739
      ScaleMode       =   0  'User
      ScaleWidth      =   65.217
      TabIndex        =   17
      Top             =   720
      Width           =   1125
   End
   Begin VB.PictureBox picSubmit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7560
      Picture         =   "frmEditor.frx":5A9C
      ScaleHeight     =   21.739
      ScaleMode       =   0  'User
      ScaleWidth      =   65.217
      TabIndex        =   16
      Top             =   4080
      Width           =   1125
   End
   Begin VB.PictureBox picCyan 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8310
      Picture         =   "frmEditor.frx":7122
      ScaleHeight     =   21.739
      ScaleMode       =   0  'User
      ScaleWidth      =   21.739
      TabIndex        =   14
      Top             =   2880
      Width           =   375
   End
   Begin VB.PictureBox picBlue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7935
      Picture         =   "frmEditor.frx":78D0
      ScaleHeight     =   21.739
      ScaleMode       =   0  'User
      ScaleWidth      =   21.739
      TabIndex        =   13
      Top             =   2880
      Width           =   375
   End
   Begin VB.PictureBox picPink 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7560
      Picture         =   "frmEditor.frx":807E
      ScaleHeight     =   21.739
      ScaleMode       =   0  'User
      ScaleWidth      =   21.739
      TabIndex        =   12
      Top             =   2880
      Width           =   375
   End
   Begin VB.PictureBox PicYellow 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8310
      Picture         =   "frmEditor.frx":882C
      ScaleHeight     =   21.739
      ScaleMode       =   0  'User
      ScaleWidth      =   21.739
      TabIndex        =   11
      Top             =   2520
      Width           =   375
   End
   Begin VB.PictureBox picGreen 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7935
      Picture         =   "frmEditor.frx":8FDA
      ScaleHeight     =   21.739
      ScaleMode       =   0  'User
      ScaleWidth      =   21.739
      TabIndex        =   10
      Top             =   2520
      Width           =   375
   End
   Begin VB.PictureBox picRed 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7560
      Picture         =   "frmEditor.frx":9788
      ScaleHeight     =   21.739
      ScaleMode       =   0  'User
      ScaleWidth      =   21.739
      TabIndex        =   9
      Top             =   2520
      Width           =   375
   End
   Begin VB.PictureBox picColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7560
      Picture         =   "frmEditor.frx":9F36
      ScaleHeight     =   21.739
      ScaleMode       =   0  'User
      ScaleWidth      =   65.217
      TabIndex        =   8
      Top             =   2160
      Width           =   1125
   End
   Begin VB.PictureBox picRight 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8310
      Picture         =   "frmEditor.frx":B5BC
      ScaleHeight     =   21.739
      ScaleMode       =   0  'User
      ScaleWidth      =   21.739
      TabIndex        =   7
      Top             =   360
      Width           =   375
   End
   Begin VB.PictureBox picCenter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7935
      Picture         =   "frmEditor.frx":BD6A
      ScaleHeight     =   21.739
      ScaleMode       =   0  'User
      ScaleWidth      =   21.739
      TabIndex        =   6
      Top             =   360
      Width           =   375
   End
   Begin VB.PictureBox picLeft 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7560
      Picture         =   "frmEditor.frx":C518
      ScaleHeight     =   21.739
      ScaleMode       =   0  'User
      ScaleWidth      =   21.739
      TabIndex        =   5
      Top             =   360
      Width           =   375
   End
   Begin VB.PictureBox picUnderline 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8310
      Picture         =   "frmEditor.frx":CCC6
      ScaleHeight     =   21.739
      ScaleMode       =   0  'User
      ScaleWidth      =   21.739
      TabIndex        =   4
      Top             =   0
      Width           =   375
   End
   Begin VB.PictureBox picItalic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7935
      Picture         =   "frmEditor.frx":D474
      ScaleHeight     =   21.739
      ScaleMode       =   0  'User
      ScaleWidth      =   21.739
      TabIndex        =   3
      Top             =   0
      Width           =   375
   End
   Begin VB.PictureBox picBold 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7560
      Picture         =   "frmEditor.frx":DC22
      ScaleHeight     =   21.739
      ScaleMode       =   0  'User
      ScaleWidth      =   21.739
      TabIndex        =   2
      Top             =   0
      Width           =   375
   End
   Begin RichTextLib.RichTextBox rtbCopy 
      Height          =   4455
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   7858
      _Version        =   393217
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmEditor.frx":E3D0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub picBlue_Click()
rtbNews.Text = rtbNews.Text & "[color=blue][/color]"
End Sub

Private Sub picBold_Click()
rtbNews.Text = rtbNews.Text & "[b][/b]"
End Sub

Private Sub picCenter_Click()
rtbNews.Text = rtbNews.Text & "[center][/center]"
End Sub

Private Sub picCode_Click()
rtbNews.Text = rtbNews.Text & "[code][/code]"
End Sub

Private Sub picText_Click()
picText.Visible = False
picPreview.Visible = True
rtbNews.Text = SavedText

' Make sure all settings are nullified
rtbNews.SelStart = 0
rtbNews.SelLength = Len(rtbNews.Text)
rtbNews.SelBold = False
rtbNews.SelItalic = False
rtbNews.SelUnderline = False
rtbNews.SelColor = RGB(0, 0, 0)
rtbNews.SelLength = 0

' Unlock the control
rtbNews.Locked = False
End Sub

Private Sub picCyan_Click()
rtbNews.Text = rtbNews.Text & "[color=cyan][/color]"
End Sub

Private Sub picGreen_Click()
rtbNews.Text = rtbNews.Text & "[color=green][/color]"
End Sub

Private Sub picItalic_Click()
rtbNews.Text = rtbNews.Text & "[i][/i]"
End Sub

Private Sub picLeft_Click()
rtbNews.Text = rtbNews.Text & "[left][/left]"
End Sub

Private Sub picPink_Click()
rtbNews.Text = rtbNews.Text & "[color=pink][/color]"
End Sub

Private Sub picPreview_Click()
picPreview.Visible = False
picText.Visible = True
' Save the text
SavedText = rtbNews.Text
rtbCopy = rtbNews
Call FilterText(frmEditor, rtbCopy.Text)
rtbNews.Locked = True
End Sub

Private Sub picRed_Click()
rtbNews.Text = rtbNews.Text & "[color=red][/color]"
End Sub

Private Sub picRight_Click()
rtbNews.Text = rtbNews.Text & "[right][/right]"
End Sub

Private Sub picUnderline_Click()
rtbNews.Text = rtbNews.Text & "[u][/u]"
End Sub

Private Sub PicYellow_Click()
rtbNews.Text = rtbNews.Text & "[color=yellow][/color]"
End Sub
