VERSION 5.00
Begin VB.Form frmNewChar 
   BackColor       =   &H00004080&
   BorderStyle     =   0  'None
   ClientHeight    =   4485
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   3825
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmNewChar.frx":0000
   ScaleHeight     =   4485
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H002F3336&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1215
      MaxLength       =   20
      TabIndex        =   0
      Top             =   945
      Width           =   2325
   End
   Begin VB.PictureBox picSprite 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   285
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   14
      Top             =   1875
      Width           =   480
      Begin VB.PictureBox picCurrentSprite 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   3285
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   15
         Top             =   2040
         Width           =   495
      End
   End
   Begin VB.OptionButton optMale 
      BackColor       =   &H002F3336&
      Caption         =   "Male"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1755
      TabIndex        =   2
      Top             =   2130
      Value           =   -1  'True
      Width           =   750
   End
   Begin VB.PictureBox picHeading 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      Picture         =   "frmNewChar.frx":38144
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   255
      TabIndex        =   12
      Top             =   0
      Width           =   3825
   End
   Begin VB.PictureBox picCancel 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   2415
      Picture         =   "frmNewChar.frx":44486
      ScaleHeight     =   420
      ScaleWidth      =   1395
      TabIndex        =   11
      Top             =   4065
      Width           =   1395
   End
   Begin VB.PictureBox picAddChar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   15
      Negotiate       =   -1  'True
      Picture         =   "frmNewChar.frx":46368
      ScaleHeight     =   420
      ScaleWidth      =   1395
      TabIndex        =   10
      Top             =   4065
      Width           =   1395
   End
   Begin VB.OptionButton optFemale 
      BackColor       =   &H002F3336&
      Caption         =   "Female"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2580
      TabIndex        =   3
      Top             =   2130
      Width           =   990
   End
   Begin VB.ComboBox cmbClass 
      BackColor       =   &H002F3336&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      IntegralHeight  =   0   'False
      Left            =   1215
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1350
      Width           =   2325
   End
   Begin VB.Label lblMP 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   240
      TabIndex        =   13
      Top             =   3232
      Width           =   1230
   End
   Begin VB.Label lblMAGI 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   2100
      TabIndex        =   9
      Top             =   3360
      Width           =   1380
   End
   Begin VB.Label lblDEF 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   2100
      TabIndex        =   8
      Top             =   3105
      Width           =   1380
   End
   Begin VB.Label lblSP 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   240
      TabIndex        =   7
      Top             =   3555
      Width           =   1230
   End
   Begin VB.Label lblSPEED 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   2100
      TabIndex        =   6
      Top             =   3615
      Width           =   1380
   End
   Begin VB.Label lblSTR 
      BackColor       =   &H002F3336&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   2100
      TabIndex        =   5
      Top             =   2850
      Width           =   1380
   End
   Begin VB.Label lblHP 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   240
      TabIndex        =   4
      Top             =   2910
      Width           =   1230
   End
End
Attribute VB_Name = "frmNewChar"
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
    PicArray(n).Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "PicFrame" & n & "Target"))
    PicArray(n).Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "PicFrame" & n & "X"))
    PicArray(n).top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "PicFrame" & n & "Y"))
    PicArray(n).Visible = True
End Sub

Public Sub SetArray(ByVal num As Long)
    ReDim PicArray(1 To num) As VB.PictureBox
End Sub

Private Sub cmbClass_Click()
    lblHP.Caption = STRING_HP + ": " + STR(Class(cmbClass.ListIndex).HP)
    lblMP.Caption = STRING_MP + ": " + STR(Class(cmbClass.ListIndex).MP)
    lblSP.Caption = STRING_SP + ": " + STR(Class(cmbClass.ListIndex).SP)
    lblSTR.Caption = STRING_STRENGTH + ": " + STR(Class(cmbClass.ListIndex).STR)
    lblDEF.Caption = STRING_DEFENSE + ": " + STR(Class(cmbClass.ListIndex).DEF)
    lblSPEED.Caption = STRING_SPEED + ": " + STR(Class(cmbClass.ListIndex).Speed)
    lblMAGI.Caption = STRING_MAGIC + ": " + STR(Class(cmbClass.ListIndex).MAGI)
    
    'Set the sprite
    picCurrentSprite.top = (((Class(cmbClass.ListIndex).Sprite * PIC_Y) * -1))
End Sub


Private Sub picAddChar_Click()
Dim Msg As String
Dim i As Long

    If Trim$(txtName.Text) <> "" Then
        Msg = Trim$(txtName.Text)
        
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

