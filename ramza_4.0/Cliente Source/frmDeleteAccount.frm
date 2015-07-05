VERSION 5.00
Begin VB.Form frmDeleteAccount 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Delete Account"
   ClientHeight    =   5235
   ClientLeft      =   150
   ClientTop       =   -30
   ClientWidth     =   4260
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmDeleteAccount.frx":0000
   ScaleHeight     =   5235
   ScaleWidth      =   4260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      BackColor       =   &H00404040&
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
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   720
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2520
      Width           =   2460
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00404040&
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
      Height          =   270
      Left            =   720
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1680
      Width           =   2460
   End
   Begin VB.Label picCancel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Volver al Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   4920
      Width           =   1650
   End
   Begin VB.Label Label1 
      BackColor       =   &H00789298&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de la Cuenta"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   1080
      TabIndex        =   4
      Top             =   1440
      Width           =   1545
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1440
      TabIndex        =   3
      Top             =   2280
      Width           =   825
   End
   Begin VB.Label picConnect 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Borrar Cuenta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   3120
      Width           =   1785
   End
End
Attribute VB_Name = "frmDeleteAccount"
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
 
        If FileExist("GUI\DeleteAccount" & Ending) Then frmDeleteAccount.Picture = LoadPicture(App.Path & "\GUI\DeleteAccount" & Ending)
    Next i
End Sub

Private Sub picCancel_Click()
    frmMainMenu.Visible = True
    frmDeleteAccount.Visible = False
End Sub

Private Sub picConnect_Click()
    If Trim(txtName.Text) <> "" And Trim(txtPassword.Text) <> "" Then
        If Len(Trim(txtName.Text)) < 3 Or Len(Trim(txtPassword.Text)) < 3 Then
            MsgBox "Tu nombre y contraseña debe tener por lo menos 3 caracteres"
            Exit Sub
        End If
        Call MenuState(MENU_STATE_DELACCOUNT)
    End If
End Sub

