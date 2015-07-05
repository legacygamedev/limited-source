VERSION 5.00
Begin VB.Form frmNewAccount 
   BackColor       =   &H00004080&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2775
   ClientLeft      =   15
   ClientTop       =   -90
   ClientWidth     =   2970
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
   ScaleHeight     =   2775
   ScaleWidth      =   2970
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H0009E7F2&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1130
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1960
      Width           =   1695
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H0009E7F2&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1130
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1600
      Width           =   1695
   End
   Begin VB.PictureBox picHeading 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   -10
      Picture         =   "frmNewAccount.frx":1C49A
      ScaleHeight     =   780
      ScaleWidth      =   3000
      TabIndex        =   4
      Top             =   -10
      Width           =   3000
   End
   Begin VB.PictureBox picConnect 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      Picture         =   "frmNewAccount.frx":23EBC
      ScaleHeight     =   375
      ScaleWidth      =   1500
      TabIndex        =   3
      Top             =   2410
      Width           =   1500
   End
   Begin VB.PictureBox picCancel 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1490
      Picture         =   "frmNewAccount.frx":25C4A
      ScaleHeight     =   375
      ScaleWidth      =   1500
      TabIndex        =   2
      Top             =   2410
      Width           =   1500
   End
End
Attribute VB_Name = "frmNewAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PicArray() As VB.PictureBox

Public Sub MakePic(ByVal n As Long)
    Set PicArray(n) = Controls.Add("VB.PictureBox", "PicArray" & CStr(n), Me)
    PicArray(n).Appearance = 0
    PicArray(n).BorderStyle = 0
    PicArray(n).AutoRedraw = True
    PicArray(n).AutoSize = True
    PicArray(n).Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "New Account", "PicFrame" & n & "Target"))
    PicArray(n).Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Account", "PicFrame" & n & "X"))
    PicArray(n).top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Account", "PicFrame" & n & "Y"))
    PicArray(n).Visible = True
End Sub

Public Sub SetArray(ByVal num As Long)
    ReDim PicArray(1 To num) As VB.PictureBox
End Sub

Private Sub picCancel_Click()
    Call LoadMenu
    frmNewAccount.Visible = False
End Sub

Private Sub picConnect_Click()
Dim Msg As String
Dim i As Long

    If Trim$(txtName.Text) <> "" And Trim$(txtPassword.Text) <> "" Then
        Msg = Trim$(txtName.Text)
        
        ' Prevent high ascii chars
        For i = 1 To Len(Msg)
            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 126 Then
                Call MsgBox("You cannot use high ascii chars in your name, please reenter.", vbOKOnly, GAME_NAME)
                txtName.Text = ""
                Exit Sub
            End If
        Next i
    
        Call MenuState(MENU_STATE_NEWACCOUNT)
    End If
End Sub
