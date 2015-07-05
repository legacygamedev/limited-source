VERSION 5.00
Begin VB.Form frmNewChar 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "New Character"
   ClientHeight    =   5985
   ClientLeft      =   90
   ClientTop       =   -60
   ClientWidth     =   5985
   ControlBox      =   0   'False
   Icon            =   "frmNewChar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmNewChar.frx":0FC2
   ScaleHeight     =   399
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optFemale 
      BackColor       =   &H00789298&
      Caption         =   "Female"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3840
      TabIndex        =   19
      Top             =   3360
      Width           =   855
   End
   Begin VB.OptionButton optMale 
      BackColor       =   &H00789298&
      Caption         =   "Male"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3840
      TabIndex        =   18
      Top             =   3120
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   3960
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   35
      TabIndex        =   8
      Top             =   2160
      Width           =   555
      Begin VB.PictureBox Picpic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   15
         ScaleHeight     =   31
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   31
         TabIndex        =   9
         Top             =   15
         Width           =   495
      End
   End
   Begin VB.PictureBox Picsprites 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5880
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   3
      Top             =   5880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer Timer2 
      Interval        =   250
      Left            =   4560
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   5040
      Top             =   480
   End
   Begin VB.ComboBox cmbClass 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   1020
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2430
      Width           =   2835
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   375
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1485
      Width           =   2760
   End
   Begin VB.Label lblSPEED 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   3240
      TabIndex        =   25
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "SPEED"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2760
      TabIndex        =   24
      Top             =   3240
      Width           =   540
   End
   Begin VB.Label lblDEF 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   2160
      TabIndex        =   23
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "DEF"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1920
      TabIndex        =   22
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label lblSTR 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   1350
      TabIndex        =   21
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "STR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1080
      TabIndex        =   20
      Top             =   3240
      Width           =   360
   End
   Begin VB.Label lblMAGI 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   4080
      TabIndex        =   17
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "MAGI"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3600
      TabIndex        =   16
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label lblSP 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   3000
      TabIndex        =   15
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "SP"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   2760
      TabIndex        =   14
      Top             =   2880
      Width           =   300
   End
   Begin VB.Label lblMP 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   2160
      TabIndex        =   13
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "MP"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1920
      TabIndex        =   12
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label lblHP 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   1320
      TabIndex        =   11
      Top             =   2880
      Width           =   405
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "HP"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1080
      TabIndex        =   10
      Top             =   2880
      Width           =   240
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Vocation"
      Height          =   210
      Left            =   1080
      TabIndex        =   7
      Top             =   2040
      Width           =   885
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Character Name"
      Height          =   210
      Left            =   405
      TabIndex        =   6
      Top             =   1230
      Width           =   1755
   End
   Begin VB.Label picAddChar 
      BackStyle       =   0  'Transparent
      Caption         =   "Create New Character"
      Height          =   240
      Left            =   2160
      TabIndex        =   5
      Top             =   3720
      Width           =   1665
   End
   Begin VB.Label picCancel 
      BackStyle       =   0  'Transparent
      Caption         =   "Back To Character Screen"
      Height          =   240
      Left            =   2010
      TabIndex        =   4
      Top             =   5340
      Width           =   1965
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Sex"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6960
      TabIndex        =   2
      Top             =   2190
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "frmNewChar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public animi As Long

Private Sub cmbClass_Click()
    lblHP.Caption = STR(Class(cmbClass.ListIndex).HP)
    lblMP.Caption = STR(Class(cmbClass.ListIndex).MP)
    lblSP.Caption = STR(Class(cmbClass.ListIndex).SP)
    
    lblSTR.Caption = STR(Class(cmbClass.ListIndex).STR)
    lblDEF.Caption = STR(Class(cmbClass.ListIndex).DEF)
    lblSPEED.Caption = STR(Class(cmbClass.ListIndex).Speed)
    lblMAGI.Caption = STR(Class(cmbClass.ListIndex).MAGI)
End Sub

Private Sub picAddChar_Click()
Dim Msg As String
Dim i As Long

    If Trim(txtName.Text) <> "" Then
        Msg = Trim(txtName.Text)
        
        If Len(Trim(txtName.Text)) < 3 Then
            MsgBox "Character name must be at least three characters in length."
            Exit Sub
        End If
        
        ' Prevent high ascii chars
        For i = 1 To Len(Msg)
            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 126 Then
                Call MsgBox("You cannot use high ascii chars in your name, please reenter.", vbOKOnly, GAME_NAME)
                txtName.Text = ""
                Exit Sub
            End If
        Next i
        
        Call MenuState(MENU_STATE_ADDCHAR)
    End If
End Sub

Private Sub picCancel_Click()
    frmChars.Visible = True
    Me.Visible = False
End Sub

Private Sub Timer1_Timer()
If cmbClass.ListIndex < 0 Then Exit Sub
If optMale.Value = True Then
    Call BitBlt(picPic.hDC, 0, 0, PIC_X, PIC_Y, picSprites.hDC, animi * PIC_X, Int(Class(cmbClass.ListIndex).MaleSprite) * PIC_Y, SRCCOPY)
Else
    Call BitBlt(picPic.hDC, 0, 0, PIC_X, PIC_Y, picSprites.hDC, animi * PIC_X, Int(Class(cmbClass.ListIndex).FemaleSprite) * PIC_Y, SRCCOPY)
End If
End Sub

Private Sub Form_Load()
Dim i As Long
Dim Ending As String
    For i = 1 To 3
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"
 
        If FileExist("GUI\NewCharacter" & Ending) Then frmNewChar.Picture = LoadPicture(App.Path & "\GUI\NewCharacter" & Ending)
    Next i
    picSprites.Picture = LoadPicture(App.Path & "\GFX\sprites.bmp")
End Sub

Private Sub Timer2_Timer()
    animi = animi + 1
If animi > 4 Then
    animi = 3
End If
End Sub
