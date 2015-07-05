VERSION 5.00
Begin VB.Form frmNewChar 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Character"
   ClientHeight    =   5250
   ClientLeft      =   135
   ClientTop       =   375
   ClientWidth     =   7500
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmNewChar.frx":0000
   ScaleHeight     =   350
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   500
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optFemale 
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
      Left            =   960
      MaskColor       =   &H8000000F&
      Picture         =   "frmNewChar.frx":2879
      TabIndex        =   14
      Top             =   2910
      Width           =   975
   End
   Begin VB.OptionButton optMale 
      Caption         =   "Male"
      DisabledPicture =   "frmNewChar.frx":2DBC
      DownPicture     =   "frmNewChar.frx":2F23
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
      Left            =   960
      MaskColor       =   &H8000000F&
      Picture         =   "frmNewChar.frx":3466
      TabIndex        =   15
      Top             =   2640
      Value           =   -1  'True
      Width           =   975
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
      Left            =   930
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3390
      Width           =   2835
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   6000
      Top             =   360
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   6480
      Top             =   360
   End
   Begin VB.PictureBox Picsprites 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4680
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   4
      Top             =   4440
      Width           =   495
      Visible         =   0   'False
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   3195
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   35
      TabIndex        =   2
      Top             =   2400
      Width           =   555
      Begin VB.PictureBox Picpic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   930
         Left            =   15
         LinkTimeout     =   65
         ScaleHeight     =   60
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   31
         TabIndex        =   3
         Top             =   15
         Width           =   495
      End
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   345
      Left            =   945
      MaxLength       =   20
      TabIndex        =   1
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Label picAddChar 
      BackStyle       =   0  'Transparent
      Height          =   360
      Left            =   5280
      TabIndex        =   13
      Top             =   3480
      Width           =   1665
   End
   Begin VB.Label lblMAGI 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Left            =   5520
      TabIndex        =   12
      Top             =   2940
      Width           =   375
   End
   Begin VB.Label lblDEF 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Left            =   5520
      TabIndex        =   11
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label lblSP 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Left            =   5520
      TabIndex        =   10
      Top             =   2070
      Width           =   375
   End
   Begin VB.Label lblSPEED 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Left            =   5520
      TabIndex        =   9
      Top             =   2730
      Width           =   375
   End
   Begin VB.Label lblMP 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Left            =   5520
      TabIndex        =   8
      Top             =   1845
      Width           =   375
   End
   Begin VB.Label lblSTR 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Left            =   5520
      TabIndex        =   7
      Top             =   2295
      Width           =   375
   End
   Begin VB.Label lblHP 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   225
      Left            =   5520
      TabIndex        =   6
      Top             =   1620
      Width           =   375
   End
   Begin VB.Label picCancel 
      BackStyle       =   0  'Transparent
      Height          =   360
      Left            =   6480
      TabIndex        =   0
      Top             =   4800
      Width           =   885
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
    lblHP.Caption = str(Class(cmbClass.ListIndex).HP)
    lblMP.Caption = str(Class(cmbClass.ListIndex).MP)
    lblSP.Caption = str(Class(cmbClass.ListIndex).SP)
    
    lblSTR.Caption = str(Class(cmbClass.ListIndex).str)
    lblDEF.Caption = str(Class(cmbClass.ListIndex).DEF)
    lblSPEED.Caption = str(Class(cmbClass.ListIndex).Speed)
    lblMAGI.Caption = str(Class(cmbClass.ListIndex).MAGI)
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
    Dim Packet As String
    Packet = "charwindow" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
If cmbClass.ListIndex < 0 Then Exit Sub
If optMale.Value = True Then
 If size1 = 0 Then
    Call BitBlt(picPic.hDC, 0, 0, PIC_X, 64, picSprites.hDC, animi * PIC_X, Int(Class(cmbClass.ListIndex).MaleSprite) * PIC_Y, SRCCOPY)
End If
 If size1 = 1 Then
    Call BitBlt(picPic.hDC, 0, 0, PIC_X, 32, picSprites.hDC, animi * PIC_X, Int(Class(cmbClass.ListIndex).MaleSprite) * PIC_Y, SRCCOPY)
 End If
End If
If optMale.Value = False Then
 If size1 = 1 Then
    Call BitBlt(picPic.hDC, 0, 0, PIC_X, 64, picSprites.hDC, animi * PIC_X, Int(Class(cmbClass.ListIndex).FemaleSprite) * PIC_Y, SRCCOPY)
   Else
    Call BitBlt(picPic.hDC, 0, 0, PIC_X, PIC_Y, picSprites.hDC, animi * PIC_X, Int(Class(cmbClass.ListIndex).FemaleSprite) * PIC_Y, SRCCOPY)
 End If
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
If animi > 11 Then
    animi = 1
End If
End Sub
