VERSION 5.00
Begin VB.Form frmChars 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Characters"
   ClientHeight    =   9465
   ClientLeft      =   165
   ClientTop       =   390
   ClientWidth     =   12855
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   631
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   857
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picCharsel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   3240
      ScaleHeight     =   191
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   399
      TabIndex        =   0
      Top             =   3720
      Width           =   6015
      Begin VB.PictureBox picSpriteloader 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   1680
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   8
         Top             =   2040
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   840
         Top             =   2040
      End
      Begin VB.PictureBox picChar3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   4320
         ScaleHeight     =   78.961
         ScaleMode       =   0  'User
         ScaleWidth      =   65.376
         TabIndex        =   7
         Top             =   480
         Width           =   1455
      End
      Begin VB.PictureBox picChar2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   2280
         ScaleHeight     =   78.961
         ScaleMode       =   0  'User
         ScaleWidth      =   65.376
         TabIndex        =   6
         Top             =   480
         Width           =   1455
      End
      Begin VB.PictureBox picChar1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   240
         ScaleHeight     =   78.961
         ScaleMode       =   0  'User
         ScaleWidth      =   65.376
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
      Begin VB.ListBox lstChars 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         ItemData        =   "frmChars.frx":0000
         Left            =   5160
         List            =   "frmChars.frx":0002
         TabIndex        =   1
         Top             =   2640
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.PictureBox picSel1 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   120
         ScaleHeight     =   92.26
         ScaleMode       =   0  'User
         ScaleWidth      =   76.387
         TabIndex        =   9
         Top             =   360
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.PictureBox picSel2 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   2160
         ScaleHeight     =   92.26
         ScaleMode       =   0  'User
         ScaleWidth      =   76.387
         TabIndex        =   10
         Top             =   360
         Visible         =   0   'False
         Width           =   1695
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            ForeColor       =   &H80000008&
            Height          =   1695
            Left            =   0
            ScaleHeight     =   92.26
            ScaleMode       =   0  'User
            ScaleWidth      =   76.387
            TabIndex        =   11
            Top             =   0
            Visible         =   0   'False
            Width           =   1695
         End
      End
      Begin VB.PictureBox picSel3 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   4200
         ScaleHeight     =   92.26
         ScaleMode       =   0  'User
         ScaleWidth      =   76.387
         TabIndex        =   12
         Top             =   360
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label picCancel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Back"
         Height          =   255
         Left            =   4200
         TabIndex        =   4
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label picDelChar 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Delete"
         Height          =   255
         Left            =   2160
         TabIndex        =   3
         Top             =   2520
         Width           =   1740
      End
      Begin VB.Label picUseChar 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
         Height          =   225
         Left            =   120
         TabIndex        =   2
         Top             =   2520
         Width           =   1665
      End
   End
End
Attribute VB_Name = "frmChars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim i As Long
Dim Ending As String
    For i = 1 To 3
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"
 
        If FileExist("GUI\CharacterSelect" & Ending) Then frmChars.Picture = LoadPicture(App.Path & "\GUI\CharacterSelect" & Ending)
    Next i
        If FileExist("GFX\sprites.bmp") Then frmChars.picSpriteloader.Picture = LoadPicture(App.Path & "\GFX\sprites.bmp")
End Sub

Private Sub picCancel_Click()
    Call TcpDestroy
    frmLogin.Visible = True
    Me.Visible = False
End Sub

Private Sub picChar1_Click()
    lstChars.ListIndex = 0
    picSel1.Visible = True
    picSel2.Visible = False
    picSel3.Visible = False
End Sub

Private Sub picChar2_Click()
    lstChars.ListIndex = 1
    picSel1.Visible = False
    picSel2.Visible = True
    picSel3.Visible = False
End Sub

Private Sub picChar3_Click()
    lstChars.ListIndex = 2
    picSel1.Visible = False
    picSel2.Visible = False
    picSel3.Visible = True
End Sub

Private Sub picCharsel_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    SOffsetX = x
    SOffsetY = y
End Sub


Private Sub picCharsel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        Call MovePicture(frmChars.picCharsel, Button, Shift, x, y)
End Sub

Private Sub picNewChar_Click()
    If lstChars.List(lstChars.ListIndex) <> "Free Character Slot" Then
        MsgBox "There is already a character in this slot!"
        Exit Sub
    End If
    Call MenuState(MENU_STATE_NEWCHAR)
End Sub

Private Sub picUseChar_Click()
    If lstChars.List(lstChars.ListIndex) = "Free Character Slot" Then
        Call MenuState(MENU_STATE_NEWCHAR)
    Else
        Call MenuState(MENU_STATE_USECHAR)
    End If
End Sub

Private Sub picDelChar_Click()
Dim value As Long

    If lstChars.List(lstChars.ListIndex) = "Free Character Slot" Then
        MsgBox "There is no character in this slot!"
        Exit Sub
    End If

    value = MsgBox("Are you sure you wish to delete this character?", vbYesNo, GAME_NAME)
    If value = vbYes Then
        Call MenuState(MENU_STATE_DELCHAR)
    End If
End Sub

Private Sub Timer1_Timer()
If charselsprite(1) <> 0 Then
    Call BitBlt(frmChars.picChar1.hDC, (picChar1.Width / 2) - 16, (picChar1.Height / 2) - 16, PIC_X, PIC_Y, frmChars.picSpriteloader.hDC, 3 * PIC_X, charselsprite(1) * PIC_Y, SRCCOPY)
End If
If charselsprite(2) <> 0 Then
    Call BitBlt(frmChars.picChar2.hDC, (picChar2.Width / 2) - 16, (picChar2.Height / 2) - 16, PIC_X, PIC_Y, frmChars.picSpriteloader.hDC, 3 * PIC_X, charselsprite(2) * PIC_Y, SRCCOPY)
End If
If charselsprite(3) <> 0 Then
    Call BitBlt(frmChars.picChar3.hDC, (picChar3.Width / 2) - 16, (picChar3.Height / 2) - 16, PIC_X, PIC_Y, frmChars.picSpriteloader.hDC, 3 * PIC_X, charselsprite(3) * PIC_Y, SRCCOPY)
End If
End Sub
