VERSION 5.00
Begin VB.Form frmNewAccount 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "M:RPGe (New Account)"
   ClientHeight    =   6540
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   7680
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
   Icon            =   "frmNewAccount.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   7680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   390
      Left            =   2160
      MaxLength       =   20
      TabIndex        =   1
      Top             =   3120
      Width           =   3375
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   2160
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   4080
      Width           =   3375
   End
   Begin VB.Image Image4 
      Height          =   195
      Left            =   7080
      Picture         =   "frmNewAccount.frx":0442
      Top             =   120
      Width           =   195
   End
   Begin VB.Image Image3 
      Height          =   195
      Left            =   7320
      Picture         =   "frmNewAccount.frx":25B5
      Top             =   120
      Width           =   195
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000C&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000C&
      Height          =   465
      Left            =   2135
      Top             =   4050
      Width           =   3435
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000C&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000C&
      Height          =   465
      Left            =   2135
      Top             =   3090
      Width           =   3435
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2160
      TabIndex        =   2
      Top             =   3720
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   0
      Picture         =   "frmNewAccount.frx":475B
      Top             =   0
      Width           =   7680
   End
   Begin VB.Image picCancel 
      Height          =   480
      Left            =   480
      Picture         =   "frmNewAccount.frx":732A
      Top             =   3360
      Width           =   1320
   End
   Begin VB.Image picConnect 
      Height          =   480
      Left            =   480
      Picture         =   "frmNewAccount.frx":780A
      Top             =   2760
      Width           =   1320
   End
   Begin VB.Image Image2 
      Height          =   6195
      Left            =   0
      Picture         =   "frmNewAccount.frx":7C9E
      Top             =   360
      Width           =   7680
   End
End
Attribute VB_Name = "frmNewAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Image3_Click()
Call GameDestroy
End Sub

Private Sub Image4_Click()
frmNewAccount.WindowState = vbMinimized
End Sub

Private Sub picCancel_Click()
    frmMainMenu.Visible = True
    frmNewAccount.Visible = False
End Sub




Private Sub picConnect_Click()
Dim Msg As String
Dim i As Long

txtName.text = LCase(txtName.text)
    If Trim(txtName.text) <> "" Then
        Msg = Trim(txtName.text)
        
        ' Prevent high ascii chars
        For i = 1 To Len(Msg)
            If Asc(Mid(Msg, i, 1)) <> 95 Then
                If Asc(Mid(Msg, i, 1)) < 97 Or Asc(Mid(Msg, i, 1)) > 123 Then
                    Call MsgBox("You cannot use high ascii chars in your name, please reenter.", vbOKOnly, GAME_NAME)
                    txtName.text = ""
                    Exit Sub
                End If
            End If
        Next i
        Dim strName As String
        Dim underscoreCount As Long
        Dim blnNextUC As Boolean
        blnNextUC = True
        underscoreCount = 0
        strName = ""
        For i = 1 To Len(txtName.text)
            If blnNextUC Then
                strName = strName & UCase(Mid(txtName.text, i, 1))
                blnNextUC = False
            Else
                If Asc(Mid(Msg, i, 1)) = 95 Then
                    blnNextUC = True
                    underscoreCount = underscoreCount + 1
                End If
                If underscoreCount < 2 Then
                    strName = strName & Mid(txtName.text, i, 1)
                Else
                    Call MsgBox("You cannot use high ascii chars in your name, please reenter.", vbOKOnly, GAME_NAME)
                    txtName.text = ""
                End If
            End If
            
        Next i
        txtName.text = strName
        Call MenuState(MENU_STATE_NEWACCOUNT)
    End If
End Sub




