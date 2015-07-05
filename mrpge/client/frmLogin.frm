VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "M:RPGe (Login)"
   ClientHeight    =   6540
   ClientLeft      =   0
   ClientTop       =   -120
   ClientWidth     =   7680
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   7680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPort 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   2250
      MaxLength       =   20
      TabIndex        =   3
      Text            =   "10001"
      Top             =   5640
      Width           =   3195
   End
   Begin VB.TextBox txtIP 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   2250
      MaxLength       =   20
      TabIndex        =   2
      Text            =   "localhost"
      Top             =   4680
      Width           =   3195
   End
   Begin VB.CheckBox chkSaveLogin 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5760
      MaskColor       =   &H00008000&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6000
      Width           =   255
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   2250
      MaxLength       =   20
      TabIndex        =   0
      Top             =   3000
      Width           =   3195
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
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2250
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3840
      Width           =   3195
   End
   Begin VB.Label Label5 
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
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label4 
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
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Port:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2220
      TabIndex        =   7
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "IP Address:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2220
      TabIndex        =   6
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   420
      Left            =   2220
      Top             =   5610
      Width           =   3270
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   420
      Left            =   2220
      Top             =   4655
      Width           =   3270
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   480
      Picture         =   "frmLogin.frx":0442
      Top             =   3960
      Width           =   1320
   End
   Begin VB.Image Image4 
      Height          =   195
      Left            =   7080
      Picture         =   "frmLogin.frx":0945
      Top             =   120
      Width           =   195
   End
   Begin VB.Image Image3 
      Height          =   195
      Left            =   7320
      Picture         =   "frmLogin.frx":2AB8
      Top             =   120
      Width           =   195
   End
   Begin VB.Image picConnect 
      Height          =   480
      Left            =   480
      Picture         =   "frmLogin.frx":4C5E
      Top             =   2760
      Width           =   1320
   End
   Begin VB.Image picCreate 
      Height          =   480
      Left            =   480
      Picture         =   "frmLogin.frx":510B
      Top             =   3360
      Width           =   1320
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Save Login Info?"
      Height          =   255
      Left            =   6120
      TabIndex        =   5
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Image picCancel 
      Height          =   480
      Left            =   480
      Picture         =   "frmLogin.frx":55D8
      Top             =   4560
      Width           =   1320
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   0
      Picture         =   "frmLogin.frx":5AB8
      Top             =   0
      Width           =   7680
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   420
      Left            =   2220
      Top             =   2970
      Width           =   3270
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   420
      Left            =   2220
      Top             =   3810
      Width           =   3270
   End
   Begin VB.Image Image1 
      Height          =   6195
      Left            =   0
      Picture         =   "frmLogin.frx":8687
      Top             =   360
      Width           =   7680
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkSaveLogin_Click()
Dim filenum As Long
    If chkSaveLogin.value = 1 Then
    filenum = FreeFile
        Open App.Path & "\setup.txt" For Output As #filenum
            Print #filenum, "username=" & encryptData(txtName.text)
            Print #filenum, "password=" & encryptData(txtPassword.text)
            Print #filenum, "ip=" & encryptData(txtIP.text)
            Print #filenum, "port=" & encryptData(txtPort.text)
            Print #filenum, "version=" & App.Major & "." & App.Minor & "." & App.Revision
        Close #filenum
    End If
End Sub

Private Sub Form_Load()

'On Error GoTo error
Dim user As String
Dim pass As String
Dim dummy As String
Dim ip As Integer
Dim port As Integer

Dim filenum As Long
filenum = FreeFile
On Error Resume Next
     Open App.Path & "\setup.txt" For Input As #filenum
        Input #filenum, user, pass, dummy, ip, port
    Close #filenum
    DoEvents
If user <> "" And pass <> "" And ip <> "" And port <> "" Then
txtName.text = Right(decryptData(user), Len(user) - 9)
txtPassword.text = Right(decryptData(pass), Len(pass) - 9)
txtIP.text = Right(decryptData(ip), Len(ip) - 9)
txtPort.text = Right(decryptData(port), Len(port) - 9)
chkSaveLogin.value = 1
End If
End Sub

Private Sub Image3_Click()
    Call GameDestroy
End Sub

Private Sub Image4_Click()
    frmLogin.WindowState = vbMinimized
End Sub

Private Sub picCancel_Click()
    frmMainMenu.Visible = True
    frmLogin.Visible = False
End Sub

Private Sub picConnect_Click()
    Call TcpInit2(txtIP.text, Val(txtPort.text))
    If Trim(txtName.text) <> "" And Trim(txtPassword.text) <> "" Then
        Call MenuState(MENU_STATE_LOGIN)
    End If
End Sub

Public Function encryptData(ByVal data As String) As String
    Dim i
    Dim output As String
    For i = 1 To Len(data) Step 1
        'encryptArr(i) = Chr(Asc(encryptArr(i)) + 1)
        output = output & Chr(Asc(Mid(data, i, 1)) + 1)
    Next i
    encryptData = output
End Function

Public Function decryptData(ByVal data As String) As String
    Dim i
    Dim output As String
    For i = 1 To Len(data) Step 1
        'encryptArr(i) = Chr(Asc(encryptArr(i)) + 1)
        output = output & Chr(Asc(Mid(data, i, 1)) - 1)
    Next i
    decryptData = output
End Function

Private Sub picCreate_Click()
    frmNewAccount.Visible = True
    Me.Visible = False
End Sub

Private Sub txtName_Change()
chkSaveLogin.value = 0
End Sub

Private Sub txtPassword_Change()
chkSaveLogin.value = 0
End Sub



