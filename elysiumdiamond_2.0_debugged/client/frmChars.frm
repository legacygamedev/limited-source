VERSION 5.00
Begin VB.Form frmChars 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Characters"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   -45
   ClientWidth     =   6000
   ControlBox      =   0   'False
   Icon            =   "frmChars.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmChars.frx":0FC2
   ScaleHeight     =   6000
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
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
      Height          =   1830
      ItemData        =   "frmChars.frx":4991
      Left            =   240
      List            =   "frmChars.frx":4993
      TabIndex        =   0
      Top             =   1560
      Width           =   5385
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Disable Auto-Login"
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
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   5520
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.Label picUseChar 
      BackStyle       =   0  'Transparent
      Caption         =   "Use Character"
      Height          =   225
      Left            =   240
      TabIndex        =   4
      Top             =   3480
      Width           =   1065
   End
   Begin VB.Label picNewChar 
      BackStyle       =   0  'Transparent
      Caption         =   "New Character"
      Height          =   240
      Left            =   4440
      TabIndex        =   3
      Top             =   3480
      Width           =   1125
   End
   Begin VB.Label picCancel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Back To Login Screen"
      Height          =   255
      Left            =   2115
      TabIndex        =   2
      Top             =   5250
      Width           =   1710
   End
   Begin VB.Label picDelChar 
      BackStyle       =   0  'Transparent
      Caption         =   "Delete Character..."
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   3480
      Width           =   1380
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
End Sub

Private Sub Label1_Click()
    If AutoLogin = 1 Then
        Call WriteINI("CONFIG", "Auto", 0, (App.Path & "\config.ini"))
        frmChars.Label1.Visible = False
        frmMainMenu.Visible = True
        Me.Visible = False
    End If
End Sub

Private Sub picCancel_Click()
    Call TcpDestroy
    frmLogin.Visible = True
    Me.Visible = False
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
        MsgBox "There is no character in this slot!"
        Exit Sub
    End If
    Call MenuState(MENU_STATE_USECHAR)
End Sub

Private Sub picDelChar_Click()
Dim Value As Long

    If lstChars.List(lstChars.ListIndex) = "Free Character Slot" Then
        MsgBox "There is no character in this slot!"
        Exit Sub
    End If

    Value = MsgBox("Are you sure you wish to delete this character?", vbYesNo, GAME_NAME)
    If Value = vbYes Then
        Call MenuState(MENU_STATE_DELCHAR)
    End If
End Sub

