VERSION 5.00
Begin VB.Form frmNewChar 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Maikone Engine"
   ClientHeight    =   4515
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   7515
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmNewChar.frx":0000
   ScaleHeight     =   301
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   501
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optFemale 
      BackColor       =   &H00404000&
      Caption         =   "Female"
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
      Height          =   255
      Left            =   6240
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.OptionButton optMale 
      BackColor       =   &H00404000&
      Caption         =   "Male"
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
      Height          =   255
      Left            =   5040
      TabIndex        =   2
      Top             =   840
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cmbClass 
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
      Height          =   360
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2760
      Width           =   3375
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
      Left            =   360
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1800
      Width           =   3375
   End
   Begin VB.Label lblBack 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   360
      TabIndex        =   19
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label lblCreate 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   5880
      TabIndex        =   18
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label lblMAGI 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   6600
      TabIndex        =   17
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label lblDEF 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   6480
      TabIndex        =   16
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label lblSP 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   4680
      TabIndex        =   15
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label lblSPEED 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   6840
      TabIndex        =   14
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label lblMP 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   4680
      TabIndex        =   13
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label lblSTR 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   6480
      TabIndex        =   12
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label lblHP 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   4680
      TabIndex        =   11
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "SPEED:"
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
      Height          =   255
      Left            =   5760
      TabIndex        =   10
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "MAGI:"
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
      Height          =   375
      Left            =   5760
      TabIndex        =   9
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "SP:"
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
      Height          =   255
      Left            =   4080
      TabIndex        =   8
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "HP:"
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
      Height          =   255
      Left            =   4080
      TabIndex        =   7
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "MP:"
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
      Height          =   255
      Left            =   4080
      TabIndex        =   6
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "DEF:"
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
      Height          =   255
      Left            =   5760
      TabIndex        =   5
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "STR:"
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
      Height          =   255
      Left            =   5760
      TabIndex        =   4
      Top             =   1320
      Width           =   615
   End
End
Attribute VB_Name = "frmNewChar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private OldX As Integer
Private OldY As Integer
Private DragMode As Boolean
Dim MoveMe As Boolean

Private Sub cmbClass_Click()
    lblHP.Caption = STR(Class(cmbClass.ListIndex + 1).HP)
    lblMP.Caption = STR(Class(cmbClass.ListIndex + 1).MP)
    lblSP.Caption = STR(Class(cmbClass.ListIndex + 1).SP)
    
    lblSTR.Caption = STR(Class(cmbClass.ListIndex + 1).STR)
    lblDEF.Caption = STR(Class(cmbClass.ListIndex + 1).DEF)
    lblSPEED.Caption = STR(Class(cmbClass.ListIndex + 1).SPEED)
    lblMAGI.Caption = STR(Class(cmbClass.ListIndex + 1).MAGI)
End Sub

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
    frmChars.Visible = True
    Me.Visible = False
End Sub

Private Sub lblCreate_Click()
Dim Msg As String
Dim i As Long

    If Trim$(txtName.Text) <> "" Then
        Msg = Trim$(txtName.Text)
        
        ' Prevent high ascii chars
        For i = 1 To Len(Msg)
            If Asc(Mid$(Msg, i, 1)) < 32 Or Asc(Mid$(Msg, i, 1)) > 126 Then
                Call MsgBox("You cannot use high ascii chars in your name, please reenter.", vbOKOnly, GAME_NAME)
                txtName.Text = ""
                Exit Sub
            End If
        Next i
        
        Call MenuState(MENU_STATE_ADDCHAR)
    End If
End Sub

Private Sub Form_Load()
Dim i As Long
Dim Ending As String

    For i = 1 To 3
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"
 
        If FileExist("GUI\NewChar" & Ending) Then frmNewChar.Picture = LoadPicture(App.Path & "\GUI\NewChar" & Ending)
    Next i
End Sub

