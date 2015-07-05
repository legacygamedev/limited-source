VERSION 5.00
Begin VB.Form frmNewAccount 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Maikone Engine"
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   7500
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmNewAccount.frx":0000
   ScaleHeight     =   4500
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   240
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2760
      Width           =   4335
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   240
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1680
      Width           =   4335
   End
   Begin VB.Label lblCreate 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label lblBack 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   3480
      Width           =   975
   End
End
Attribute VB_Name = "frmNewAccount"
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

Private Sub lblBack_Click()
    frmMainMenu.Visible = True
    frmNewAccount.Visible = False
End Sub

Private Sub lblCreate_Click()
Dim Msg As String
Dim i As Long

    If Trim$(txtName.Text) <> "" And Trim(txtPassword.Text) <> "" Then
        Msg = Trim$(txtName.Text)
        
        ' Prevent high ascii chars
        For i = 1 To Len(Msg)
            If Asc(Mid$(Msg, i, 1)) < 32 Or Asc(Mid$(Msg, i, 1)) > 126 Then
                Call MsgBox("You cannot use high ascii chars in your name, please reenter.", vbOKOnly, GAME_NAME)
                txtName.Text = ""
                Exit Sub
            End If
        Next i
    
        Call MenuState(MENU_STATE_NEWACCOUNT)
    End If
End Sub

Private Sub Form_Load()
Dim i As Long
Dim Ending As String

    For i = 1 To 3
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"
 
        If FileExist("GUI\CreateAcc" & Ending) Then frmNewAccount.Picture = LoadPicture(App.Path & "\GUI\CreateAcc" & Ending)
    Next i
End Sub
