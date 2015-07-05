VERSION 5.00
Begin VB.Form frmChars 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Character Select"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   ControlBox      =   0   'False
   Icon            =   "frmChars.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   3750
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstChars 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   615
      ItemData        =   "frmChars.frx":2372
      Left            =   720
      List            =   "frmChars.frx":2374
      TabIndex        =   0
      Top             =   1545
      Width           =   2265
   End
   Begin VB.Label picCancel 
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      Height          =   195
      Left            =   1590
      TabIndex        =   4
      Top             =   3615
      Width           =   525
   End
   Begin VB.Label picDelChar 
      BackStyle       =   0  'Transparent
      Caption         =   "Delete character"
      Height          =   195
      Left            =   1245
      TabIndex        =   3
      Top             =   3210
      Width           =   1215
   End
   Begin VB.Label picNewChar 
      BackStyle       =   0  'Transparent
      Caption         =   "New character"
      Height          =   225
      Left            =   1305
      TabIndex        =   2
      Top             =   2775
      Width           =   1080
   End
   Begin VB.Label picUseChar 
      BackStyle       =   0  'Transparent
      Caption         =   "Use selected character"
      Height          =   210
      Left            =   1005
      TabIndex        =   1
      Top             =   2370
      Width           =   1695
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

        If FileExist("\core files\interface\chars" & Ending) Then frmChars.Picture = LoadPicture(App.Path & "\core files\interface\chars" & Ending)
    Next i
End Sub

Private Sub picCancel_Click()
    Call TcpDestroy
    frmLogin.Visible = True
    Me.Visible = False
End Sub

Private Sub picNewChar_Click()
    Call MenuState(MENU_STATE_NEWCHAR)
End Sub

Private Sub picUseChar_Click()
    Call MenuState(MENU_STATE_USECHAR)
End Sub

Private Sub picDelChar_Click()
Dim Value As Long

    Value = MsgBox("Are you sure you wish to delete this character?", vbYesNo, GAME_NAME)
    If Value = vbYes Then
        Call MenuState(MENU_STATE_DELCHAR)
    End If
End Sub

