VERSION 5.00
Begin VB.Form frmChars 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Characters"
   ClientHeight    =   5235
   ClientLeft      =   120
   ClientTop       =   -45
   ClientWidth     =   7500
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmChars.frx":0000
   ScaleHeight     =   5235
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstChars 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1110
      ItemData        =   "frmChars.frx":27B7
      Left            =   850
      List            =   "frmChars.frx":27B9
      TabIndex        =   0
      Top             =   1350
      Width           =   4680
   End
   Begin VB.Label picUseChar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   840
      TabIndex        =   4
      Top             =   2640
      Width           =   1305
   End
   Begin VB.Label picNewChar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   4080
      TabIndex        =   3
      Top             =   2640
      Width           =   1485
   End
   Begin VB.Label picCancel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6120
      TabIndex        =   2
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label picDelChar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   2640
      Width           =   1620
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

Private Sub picCancel_Click()
    Call TcpDestroy
    frmLogin.Visible = True
    Me.Visible = False
End Sub

Private Sub picNewChar_Click()
    If lstChars.List(lstChars.ListIndex) <> "Casilla de personaje Vacia" Then
        MsgBox "There is already a character in this slot!"
        Exit Sub
    End If
    Call MenuState(MENU_STATE_NEWCHAR)
End Sub

Private Sub picUseChar_Click()
    If lstChars.List(lstChars.ListIndex) = "Casilla de personaje Vacia" Then
        MsgBox "There is no character in this slot!"
        Exit Sub
    End If
    Call MenuState(MENU_STATE_USECHAR)
End Sub

Private Sub picDelChar_Click()
Dim Value As Long

    If lstChars.List(lstChars.ListIndex) = "Casilla de personaje Vacia" Then
        MsgBox "There is no character in this slot!"
        Exit Sub
    End If

    Value = MsgBox("Are you sure you wish to delete this character?", vbYesNo, GAME_NAME)
    If Value = vbYes Then
        Call MenuState(MENU_STATE_DELCHAR)
    End If
End Sub

