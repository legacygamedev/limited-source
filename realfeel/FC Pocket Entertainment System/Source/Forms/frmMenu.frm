VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmMenu 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "FriendCodes Pocket Entertainment System"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Fake editor link"
      Height          =   1575
      Left            =   2760
      TabIndex        =   21
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Fake Download link"
      Height          =   1575
      Left            =   4560
      TabIndex        =   20
      Top             =   720
      Width           =   975
   End
   Begin InetCtlsObjects.Inet cInet 
      Left            =   7080
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.PictureBox picmMusic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      Picture         =   "frmMenu.frx":0000
      ScaleHeight     =   21.739
      ScaleMode       =   0  'User
      ScaleWidth      =   65.217
      TabIndex        =   19
      Top             =   2760
      Width           =   1125
   End
   Begin VB.PictureBox picMusic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      Picture         =   "frmMenu.frx":1686
      ScaleHeight     =   21.739
      ScaleMode       =   0  'User
      ScaleWidth      =   65.217
      TabIndex        =   18
      Top             =   2760
      Width           =   1125
   End
   Begin VB.TextBox txtChatbar 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1125
      TabIndex        =   17
      Top             =   4170
      Visible         =   0   'False
      Width           =   6450
   End
   Begin VB.PictureBox picmSignIn 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      Picture         =   "frmMenu.frx":2D0C
      ScaleHeight     =   21.739
      ScaleMode       =   0  'User
      ScaleWidth      =   65.217
      TabIndex        =   16
      Top             =   480
      Width           =   1125
   End
   Begin VB.PictureBox picSignIn 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      Picture         =   "frmMenu.frx":4392
      ScaleHeight     =   21.739
      ScaleMode       =   0  'User
      ScaleWidth      =   65.217
      TabIndex        =   15
      Top             =   480
      Width           =   1125
   End
   Begin VB.PictureBox picmRegister 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      Picture         =   "frmMenu.frx":5A18
      ScaleHeight     =   21.739
      ScaleMode       =   0  'User
      ScaleWidth      =   65.217
      TabIndex        =   14
      Top             =   0
      Width           =   1125
   End
   Begin VB.PictureBox picRegister 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      Picture         =   "frmMenu.frx":709E
      ScaleHeight     =   21.739
      ScaleMode       =   0  'User
      ScaleWidth      =   65.217
      TabIndex        =   13
      Top             =   0
      Width           =   1125
   End
   Begin VB.PictureBox picmContacts 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      Picture         =   "frmMenu.frx":8724
      ScaleHeight     =   21.739
      ScaleMode       =   0  'User
      ScaleWidth      =   65.217
      TabIndex        =   12
      Top             =   3720
      Width           =   1125
   End
   Begin VB.PictureBox picContacts 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      Picture         =   "frmMenu.frx":9DAA
      ScaleHeight     =   21.739
      ScaleMode       =   0  'User
      ScaleWidth      =   65.217
      TabIndex        =   11
      Top             =   3720
      Width           =   1125
   End
   Begin VB.PictureBox picmMembers 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      Picture         =   "frmMenu.frx":B430
      ScaleHeight     =   21.739
      ScaleMode       =   0  'User
      ScaleWidth      =   65.217
      TabIndex        =   10
      Top             =   3240
      Width           =   1125
   End
   Begin VB.PictureBox picMembers 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      Picture         =   "frmMenu.frx":CAB6
      ScaleHeight     =   21.739
      ScaleMode       =   0  'User
      ScaleWidth      =   65.217
      TabIndex        =   9
      Top             =   3240
      Width           =   1125
   End
   Begin VB.PictureBox picmNewArticles 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   0
      Picture         =   "frmMenu.frx":E13C
      ScaleHeight     =   43.478
      ScaleMode       =   0  'User
      ScaleWidth      =   65.217
      TabIndex        =   8
      Top             =   960
      Width           =   1125
   End
   Begin VB.PictureBox picNewArticles 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   0
      Picture         =   "frmMenu.frx":10E06
      ScaleHeight     =   43.478
      ScaleMode       =   0  'User
      ScaleWidth      =   65.217
      TabIndex        =   7
      Top             =   960
      Width           =   1125
   End
   Begin VB.PictureBox picmChat 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      Picture         =   "frmMenu.frx":13AD0
      ScaleHeight     =   21.739
      ScaleMode       =   0  'User
      ScaleWidth      =   65.217
      TabIndex        =   6
      Top             =   2280
      Width           =   1125
   End
   Begin VB.PictureBox picChat 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      Picture         =   "frmMenu.frx":15156
      ScaleHeight     =   21.739
      ScaleMode       =   0  'User
      ScaleWidth      =   65.217
      TabIndex        =   5
      Top             =   2280
      Width           =   1125
   End
   Begin VB.PictureBox picFriendCodes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      Picture         =   "frmMenu.frx":167DC
      ScaleHeight     =   21.739
      ScaleMode       =   0  'User
      ScaleWidth      =   65.217
      TabIndex        =   4
      Top             =   4120
      Width           =   1125
   End
   Begin VB.PictureBox picmArticles 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      Picture         =   "frmMenu.frx":17E62
      ScaleHeight     =   21.739
      ScaleMode       =   0  'User
      ScaleWidth      =   65.217
      TabIndex        =   2
      Top             =   1800
      Width           =   1125
   End
   Begin RichTextLib.RichTextBox rtbNews 
      Height          =   4455
      Left            =   1125
      TabIndex        =   1
      Top             =   0
      Width           =   6450
      _ExtentX        =   11377
      _ExtentY        =   7858
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMenu.frx":194E8
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
   Begin RichTextLib.RichTextBox rtbCopy 
      Height          =   4455
      Left            =   1125
      TabIndex        =   0
      Top             =   0
      Width           =   6450
      _ExtentX        =   11377
      _ExtentY        =   7858
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmMenu.frx":1956C
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
   Begin VB.PictureBox picArticles 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      Picture         =   "frmMenu.frx":195F0
      ScaleHeight     =   21.739
      ScaleMode       =   0  'User
      ScaleWidth      =   65.217
      TabIndex        =   3
      Top             =   1800
      Width           =   1125
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Call Inet.Download("www.dualsolace.com/Downloads/", "RealFeel111.zip")
End Sub

Private Sub Command2_Click()
frmEditor.Visible = True
End Sub

Private Sub Form_Load()
Dim choice As Integer, f As Integer
f = FreeFile

' Make folder if it doesn't exist
If Not Exists(App.Path & "\Files\") Then MkDir (App.Path & "\Files\")

' Make file if it doesn't exist
If Not Exists(App.Path & "\Dynamic Library\data.dat") Then
    Open App.Path & "\Dynamic Library\data.dat" For Output As #f
        Print #f, "-Do not delete the lines below! It will cause errors!"
        Print #f, "Register False"
        Print #f, "[ENDFILE]"
    Close #f
End If

If Read(App.Path & "\Dynamic Library\data.dat", "Register ") = "True" Then Exit Sub

If Not Exists("C:\Windows\System32\msinet.ocx") Or Not Exists("C:\Windows\System32\Mswinsck.ocx") Or Not Exists("C:\Windows\System32\RICHTX32.ocx") Then
    choice = MsgBox("FriendCodes Pocket Entertainment System detects that not all of the dynamic library files are registered! Would you like FriendCodes Pocket Entertainment System to register them for you?", vbYesNo, "FriendCodes Pocket Entertainment System")
    If choice = vbYes Then
        If Not Exists("C:\Windows\System32\msinet.ocx") Then Call Shell(App.Path & "\Dynamic Library\msinet.bat")
        If Not Exists("C:\Windows\System32\Mswinsck.ocx") Then Call Shell(App.Path & "\Dynamic Library\Mswinsck.bat")
        If Not Exists("C:\Windows\System32\RICHTX32.ocx") Then Call Shell(App.Path & "\Dynamic Library\RICHTX32.bat")
        Call MsgBox("The library files have been properly registered!", vbOKOnly, "FriendCodes Pocket Entertainment System")
        
        ' Set the Register variable to true
        Call Place(App.Path & "\Files\data.dat", "Register ", "True")
    End If
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Hide real Register button
picmRegister.Visible = True
picRegister.Visible = False

' Hide real Sign In button
picmSignIn.Visible = True
picSignIn.Visible = False

' Hide real New Articles button
picmNewArticles.Visible = True
picNewArticles.Visible = False

' Hide real Articles button
picmArticles.Visible = True
picArticles.Visible = False

' Hide real Chat button
picmChat.Visible = True
picChat.Visible = False

' Hide real Music button
picmMusic.Visible = True
picMusic.Visible = False

' Hide real Member button
picmMembers.Visible = True
picMembers.Visible = False

' Hide real Contacts button
picmContacts.Visible = True
picContacts.Visible = False
End Sub

Private Sub picChat_Click()
' Reveal chatbar and resize the the richtextboxes
txtChatbar.Visible = True
rtbNews.Height = 4170
rtbCopy.Height = 4170
End Sub

Private Sub picmArticles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
picmArticles.Visible = False
picArticles.Visible = True
End Sub

Private Sub picmChat_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
picmChat.Visible = False
picChat.Visible = True
End Sub

Private Sub picmContacts_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
picmContacts.Visible = False
picContacts.Visible = True
End Sub

Private Sub picmMembers_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
picmMembers.Visible = False
picMembers.Visible = True
End Sub

Private Sub picmMusic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
picmMusic.Visible = False
picMusic.Visible = True
End Sub

Private Sub picmNewArticles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
picmNewArticles.Visible = False
picNewArticles.Visible = True
End Sub

Private Sub picmRegister_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
picmRegister.Visible = False
picRegister.Visible = True
End Sub

Private Sub picmSignIn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
picmSignIn.Visible = False
picSignIn.Visible = True
End Sub

Private Sub picRegister_Click()
' Add registering data
End Sub

Private Sub rtbNews_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Hide real Register button
picmRegister.Visible = True
picRegister.Visible = False

' Hide real Sign In button
picmSignIn.Visible = True
picSignIn.Visible = False

' Hide real New Articles button
picmNewArticles.Visible = True
picNewArticles.Visible = False

' Hide real Articles button
picmArticles.Visible = True
picArticles.Visible = False

' Hide real Chat button
picmChat.Visible = True
picChat.Visible = False

' Hide real Music button
picmMusic.Visible = True
picMusic.Visible = False

' Hide real Member button
picmMembers.Visible = True
picMembers.Visible = False

' Hide real Contacts button
picmContacts.Visible = True
picContacts.Visible = False
End Sub
