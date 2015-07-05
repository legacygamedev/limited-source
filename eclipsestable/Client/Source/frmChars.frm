VERSION 5.00
Begin VB.Form frmChars 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Character Selection"
   ClientHeight    =   5985
   ClientLeft      =   165
   ClientTop       =   330
   ClientWidth     =   5985
   ControlBox      =   0   'False
   Icon            =   "frmChars.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   399
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   StartUpPosition =   2  'CenterScreen
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
      Height          =   1650
      ItemData        =   "frmChars.frx":0FC2
      Left            =   810
      List            =   "frmChars.frx":0FC4
      TabIndex        =   0
      Top             =   555
      Width           =   4350
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Top             =   4800
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label picUseChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   2400
      Width           =   945
   End
   Begin VB.Label picNewChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   2400
      Width           =   1560
   End
   Begin VB.Label picCancel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label picDelChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   2400
      Width           =   1560
   End
End
Attribute VB_Name = "frmChars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim I As Long
    Dim Ending As String

    For I = 1 To 3
        If I = 1 Then
            Ending = ".gif"
        End If

        If I = 2 Then
            Ending = ".jpg"
        End If

        If I = 3 Then
            Ending = ".png"
        End If

        If FileExists("GUI\CharacterSelect" & Ending) Then
            frmChars.Picture = LoadPicture(App.Path & "\GUI\CharacterSelect" & Ending)
        End If
    Next I
End Sub

Private Sub Label1_Click()
    If AutoLogin = 1 Then
        Call WriteINI("CONFIG", "Auto", 0, (App.Path & "\config.ini"))

        Me.Visible = False
        frmMainMenu.Visible = True
    End If
End Sub

Private Sub picCancel_Click()
    Call TcpDestroy

    Me.Visible = False
    frmLogin.Visible = True
End Sub

Private Sub picNewChar_Click()
    If lstChars.List(lstChars.ListIndex) <> "Free Character Slot" Then
        MsgBox "There is already a character in this slot!"
        Exit Sub
    End If

    frmNewChar.picSprites.Picture = LoadPicture(App.Path & "\GFX\Sprites.bmp")

    Call MenuState(MENU_STATE_NEWCHAR)
End Sub

Private Sub picUseChar_Click()
    If lstChars.List(lstChars.ListIndex) = "Free Character Slot" Then
        MsgBox "There is no character in this slot!"
        Exit Sub
    End If

    frmStable.picItems.Picture = LoadPicture(App.Path & "\GFX\Items.bmp")
    frmSpriteChange.picSprites.Picture = LoadPicture(App.Path & "\GFX\Sprites.bmp")

    Call MenuState(MENU_STATE_USECHAR)
End Sub

Private Sub picDelChar_Click()
    Dim Value As Integer

    If lstChars.List(lstChars.ListIndex) = "Free Character Slot" Then
        MsgBox "There is no character in this slot!"
        Exit Sub
    End If

    Value = MsgBox("Are you sure you wish to delete this character?", vbYesNo, GAME_NAME)
    If Value = vbYes Then
        Call MenuState(MENU_STATE_DELCHAR)
    End If
End Sub
