VERSION 5.00
Begin VB.Form frmNewChar 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "New Character"
   ClientHeight    =   6000
   ClientLeft      =   90
   ClientTop       =   -60
   ClientWidth     =   6000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmNewChar.frx":0000
   ScaleHeight     =   400
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrRace 
      Interval        =   1000
      Left            =   4080
      Top             =   480
   End
   Begin VB.CheckBox chkRace3 
      BackColor       =   &H00000000&
      Caption         =   "Orc"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3720
      TabIndex        =   29
      Top             =   5280
      Width           =   975
   End
   Begin VB.CheckBox chkRace2 
      BackColor       =   &H00000000&
      Caption         =   "Elf"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   28
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CheckBox chkRace1 
      BackColor       =   &H00000000&
      Caption         =   "Human"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   27
      Top             =   5280
      Width           =   1095
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1020
      Left            =   3825
      ScaleHeight     =   66
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   35
      TabIndex        =   24
      Top             =   2490
      Width           =   555
      Begin VB.PictureBox Picpic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   15
         ScaleHeight     =   62
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   31
         TabIndex        =   25
         Top             =   15
         Width           =   495
      End
   End
   Begin VB.OptionButton optFemale 
      BackColor       =   &H00FF8080&
      Caption         =   "Female"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2520
      TabIndex        =   17
      Top             =   3240
      Width           =   855
   End
   Begin VB.OptionButton optMale 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "Male"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1680
      TabIndex        =   16
      Top             =   3240
      Value           =   -1  'True
      Width           =   735
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
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   3840
      Width           =   2835
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   1560
      MaxLength       =   20
      TabIndex        =   0
      Top             =   2880
      Width           =   1800
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Race Selection"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   210
      Left            =   2040
      TabIndex        =   30
      Top             =   4920
      Width           =   1755
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Chaos Engine"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   1080
      TabIndex        =   26
      Top             =   1320
      Width           =   3705
   End
   Begin VB.Label picAddChar 
      BackStyle       =   0  'Transparent
      Caption         =   "Create New Character"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2280
      TabIndex        =   5
      Top             =   5640
      Width           =   1665
   End
   Begin VB.Label lblSPEED 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3600
      TabIndex        =   23
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "SPEED"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3000
      TabIndex        =   22
      Top             =   4560
      Width           =   540
   End
   Begin VB.Label lblDEF 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2640
      TabIndex        =   21
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "DEF"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2280
      TabIndex        =   20
      Top             =   4560
      Width           =   495
   End
   Begin VB.Label lblSTR 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1920
      TabIndex        =   19
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "STR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1560
      TabIndex        =   18
      Top             =   4560
      Width           =   360
   End
   Begin VB.Label lblMAGI 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4320
      TabIndex        =   15
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "MAGI"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3720
      TabIndex        =   14
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label lblSP 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3360
      TabIndex        =   13
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "SP"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3000
      TabIndex        =   12
      Top             =   4200
      Width           =   300
   End
   Begin VB.Label lblMP 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2640
      TabIndex        =   11
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "MP"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label lblHP 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   1920
      TabIndex        =   9
      Top             =   4200
      Width           =   405
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "HP"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1560
      TabIndex        =   8
      Top             =   4200
      Width           =   240
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Vocation"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   1680
      TabIndex        =   7
      Top             =   3600
      Width           =   885
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Character Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   1560
      TabIndex        =   6
      Top             =   2640
      Width           =   1755
   End
   Begin VB.Label picCancel 
      BackStyle       =   0  'Transparent
      Caption         =   " X"
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   5760
      TabIndex        =   4
      Top             =   120
      Width           =   285
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

' Copyright (c) 2006 Chaos Engine Source. All rights reserved.
' This code is licensed under the Chaos Engine General License.

Option Explicit
Public animi As Long

Private Sub cmbClass_Click()
    lblHP.Caption = STR(Class(cmbClass.ListIndex + 1).HP)
    lblMP.Caption = STR(Class(cmbClass.ListIndex + 1).MP)
    lblSP.Caption = STR(Class(cmbClass.ListIndex + 1).SP)
    
    lblSTR.Caption = STR(Class(cmbClass.ListIndex + 1).STR)
    lblDEF.Caption = STR(Class(cmbClass.ListIndex + 1).DEF)
    lblSPEED.Caption = STR(Class(cmbClass.ListIndex + 1).speed)
    lblMAGI.Caption = STR(Class(cmbClass.ListIndex + 1).MAGI)
End Sub


Private Sub picAddChar_Click()
Dim Msg As String
Dim I As Long

    If Trim(txtName.Text) <> "" Then
        Msg = Trim(txtName.Text)
        
        If Len(Trim(txtName.Text)) < 3 Then
            MsgBox "Character name must be at least three characters in length."
            Exit Sub
        End If
        
        ' Prevent high ascii chars
        For I = 1 To Len(Msg)
            If Asc(mid(Msg, I, 1)) < 32 Or Asc(mid(Msg, I, 1)) > 126 Then
                Call MsgBox("You cannot use high ascii chars in your name, please reenter.", vbOKOnly, GAME_NAME)
                txtName.Text = ""
                Exit Sub
            End If
        Next I
        
        Call MenuState(MENU_STATE_ADDCHAR)
    End If
End Sub

Private Sub picCancel_Click()
    frmChars.Visible = True
    Me.Visible = False
End Sub

Private Sub Timer1_Timer()
    If cmbClass.ListIndex < 0 Then Exit Sub
    
    Picpic.Width = SIZE_X
    Picpic.Height = SIZE_Y
    Picture4.Width = SIZE_X + 4
    Picture4.Height = SIZE_Y + 4
    
    If optMale.Value = True Then
        Call BitBlt(Picpic.hDC, 0, 0, SIZE_X, SIZE_Y, Picsprites.hDC, animi * SIZE_X, Int(Class(cmbClass.ListIndex + 1).MaleSprite) * SIZE_Y, SRCCOPY)
    Else
        Call BitBlt(Picpic.hDC, 0, 0, SIZE_X, SIZE_Y, Picsprites.hDC, animi * SIZE_X, Int(Class(cmbClass.ListIndex + 1).FemaleSprite) * SIZE_Y, SRCCOPY)
    End If
End Sub

Private Sub form_load()
Dim I As Long
Dim Ending As String
Dim sDc As Long
    For I = 1 To 3
        If I = 1 Then Ending = ".gif"
        If I = 2 Then Ending = ".jpg"
        If I = 3 Then Ending = ".png"
 
    Next I
    
    sDc = DD_SpriteSurf.GetDC
    With Picsprites
        .Width = DDSD_Sprite.lWidth
        .Height = DDSD_Sprite.lHeight
        .Cls
        Call BitBlt(.hDC, 0, 0, .Width, .Height, sDc, 0, 0, SRCCOPY)
    End With
    Call DD_SpriteSurf.ReleaseDC(sDc)
End Sub

Private Sub Timer2_Timer()
    animi = animi + 1
If animi > 4 Then
    animi = 3
End If
End Sub

Private Sub tmrRace_Timer()
If chkRace1.Value = Checked Then
chkRace2.Value = Unchecked
chkRace3.Value = Unchecked
chkRace2.Enabled = False
chkRace3.Enabled = False
ElseIf chkRace2.Value = Checked Then
chkRace1.Value = Unchecked
chkRace3.Value = Unchecked
chkRace1.Enabled = False
chkRace3.Enabled = False
ElseIf chkRace3.Value = Checked Then
chkRace1.Value = Unchecked
chkRace2.Value = Unchecked
chkRace1.Enabled = False
chkRace2.Enabled = False
End If
End Sub
