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
   Begin VB.PictureBox picSprite 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   3285
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   14
      Top             =   2010
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
      BackColor       =   &H00F5763F&
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
      ForeColor       =   &H0009E7F2&
      Height          =   255
      Left            =   1665
      TabIndex        =   2
      Top             =   1260
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.PictureBox picHeading 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   30
      Picture         =   "frmNewChar.frx":38142
      ScaleHeight     =   52
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   251
      TabIndex        =   12
      Top             =   30
      Width           =   3765
   End
   Begin VB.PictureBox picCancel 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1545
      Picture         =   "frmNewChar.frx":41B14
      ScaleHeight     =   375
      ScaleWidth      =   1500
      TabIndex        =   11
      Top             =   4080
      Width           =   1500
   End
   Begin VB.PictureBox picAddChar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   30
      Negotiate       =   -1  'True
      Picture         =   "frmNewChar.frx":438A2
      ScaleHeight     =   375
      ScaleWidth      =   1500
      TabIndex        =   10
      Top             =   4080
      Width           =   1500
   End
   Begin VB.OptionButton optFemale 
      BackColor       =   &H00F5763F&
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
      ForeColor       =   &H0009E7F2&
      Height          =   255
      Left            =   2565
      TabIndex        =   3
      Top             =   1260
      Width           =   1215
   End
   Begin VB.ComboBox cmbClass 
      BackColor       =   &H00F5763F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0009E7F2&
      Height          =   405
      Left            =   1260
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1560
      Width           =   2520
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5763F&
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1260
      MaxLength       =   20
      TabIndex        =   0
      Top             =   840
      Width           =   2520
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
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   1425
      TabIndex        =   13
      Top             =   3405
      Width           =   375
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
      ForeColor       =   &H00C000C0&
      Height          =   375
      Left            =   2820
      TabIndex        =   9
      Top             =   3390
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
      ForeColor       =   &H00F5763F&
      Height          =   375
      Left            =   2970
      TabIndex        =   8
      Top             =   3075
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
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   1500
      TabIndex        =   7
      Top             =   3780
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
      ForeColor       =   &H000750BC&
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   3765
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
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   3180
      TabIndex        =   5
      Top             =   2745
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
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   1530
      TabIndex        =   4
      Top             =   3045
      Width           =   375
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
    lblHP.Caption = STR(Class(cmbClass.ListIndex).HP)
    lblMP.Caption = STR(Class(cmbClass.ListIndex).MP)
    lblSP.Caption = STR(Class(cmbClass.ListIndex).SP)
    lblSTR.Caption = STR(Class(cmbClass.ListIndex).STR)
    lblDEF.Caption = STR(Class(cmbClass.ListIndex).DEF)
    lblSPEED.Caption = STR(Class(cmbClass.ListIndex).SPEED)
    lblMAGI.Caption = STR(Class(cmbClass.ListIndex).MAGI)
    
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

