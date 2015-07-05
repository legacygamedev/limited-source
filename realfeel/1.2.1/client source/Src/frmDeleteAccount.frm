VERSION 5.00
Begin VB.Form frmDeleteAccount 
   BackColor       =   &H00004080&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2880
   ClientLeft      =   15
   ClientTop       =   -90
   ClientWidth     =   3000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmDeleteAccount.frx":0000
   ScaleHeight     =   2880
   ScaleWidth      =   3000
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1140
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1995
      Width           =   1695
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   1155
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1545
      Width           =   1695
   End
   Begin VB.PictureBox picHeading 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   0
      Picture         =   "frmDeleteAccount.frx":1C49C
      ScaleHeight     =   765
      ScaleWidth      =   3000
      TabIndex        =   3
      Top             =   0
      Width           =   3000
   End
   Begin VB.PictureBox picConnect 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   0
      Picture         =   "frmDeleteAccount.frx":23C66
      ScaleHeight     =   390
      ScaleWidth      =   1200
      TabIndex        =   2
      Top             =   2505
      Width           =   1200
   End
   Begin VB.PictureBox picCancel 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   1905
      Picture         =   "frmDeleteAccount.frx":25508
      ScaleHeight     =   390
      ScaleWidth      =   1095
      TabIndex        =   1
      Top             =   2505
      Width           =   1095
   End
End
Attribute VB_Name = "frmDeleteAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PicArray() As VB.PictureBox

Private Sub picCancel_Click()
    Call LoadMenu
    frmDeleteAccount.Visible = False
End Sub

Private Sub picConnect_Click()
    If Trim$(txtName.Text) <> "" And Trim$(txtPassword.Text) <> "" Then
        Call MenuState(MENU_STATE_DELACCOUNT)
    End If
End Sub

Public Sub MakePic(ByVal n As Long)
    Set PicArray(n) = Controls.Add("VB.PictureBox", "PicArray" & CStr(n), Me)
    PicArray(n).Appearance = 0
    PicArray(n).BorderStyle = 0
    PicArray(n).AutoRedraw = True
    PicArray(n).AutoSize = True
    PicArray(n).Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Delete Account", "PicFrame" & n & "Target"))
    PicArray(n).Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Delete Account", "PicFrame" & n & "X"))
    PicArray(n).top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Delete Account", "PicFrame" & n & "Y"))
    PicArray(n).Visible = True
End Sub

Public Sub SetArray(ByVal num As Long)
    ReDim PicArray(1 To num) As VB.PictureBox
End Sub

