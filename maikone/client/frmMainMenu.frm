VERSION 5.00
Begin VB.Form frmMainMenu 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Maikone Engine"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   4500
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMainMenu.frx":0000
   ScaleHeight     =   6000
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblExit 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3720
      TabIndex        =   5
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label lblCredits 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   3195
      Width           =   1215
   End
   Begin VB.Label lblDeleteAcc 
      BackStyle       =   0  'Transparent
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label lblNewAcc 
      BackStyle       =   0  'Transparent
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label lblIPConfig 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   1520
      Width           =   1575
   End
   Begin VB.Label lblLogin 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   855
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private OldX As Integer
Private OldY As Integer
Private DragMode As Boolean
Dim MoveMe As Boolean

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveMe = True
    OldX = X
    OldY = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MoveMe = True Then
        Me.Left = Me.Left + (X - OldX)
        Me.top = Me.top + (Y - OldY)
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.Left = Me.Left + (X - OldX)
    Me.top = Me.top + (Y - OldY)
    MoveMe = False
End Sub

Private Sub lblCredits_Click()
    frmCredits.Visible = True
    Me.Visible = False
End Sub

Private Sub lblDeleteAcc_Click()
Dim YesNo As Long

    YesNo = MsgBox("You are on the path for a character deletion, are you sure you want to go through with this?", vbYesNo, GAME_NAME)
    If YesNo = vbYes Then
        frmDeleteAccount.Visible = True
        Me.Visible = False
    End If
End Sub

Private Sub lblExit_Click()
    Call GameDestroy
End Sub

Private Sub lblIPConfig_Click()
    frmIPConfig.Visible = True
    Me.Visible = False
End Sub

Private Sub lblLogin_Click()
    frmLogin.Visible = True
    Me.Visible = False
End Sub

Private Sub lblNewAcc_Click()
    frmNewAccount.Visible = True
    Me.Visible = False
End Sub

Private Sub Form_Load()
Dim i As Long
Dim Ending As String

    For i = 1 To 3
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"
 
        If FileExist("GUI\MainMenu" & Ending) Then frmMainMenu.Picture = LoadPicture(App.Path & "\GUI\MainMenu" & Ending)
    Next i
End Sub
