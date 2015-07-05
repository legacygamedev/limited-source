VERSION 5.00
Begin VB.Form frmNewChar 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Character"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4065
   ControlBox      =   0   'False
   Icon            =   "frmNewChar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   306
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   271
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicSprites 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5760
      Left            =   2685
      Picture         =   "frmNewChar.frx":2372
      ScaleHeight     =   5760
      ScaleWidth      =   5760
      TabIndex        =   31
      Top             =   3360
      Visible         =   0   'False
      Width           =   5760
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   3030
      Top             =   2160
   End
   Begin VB.Timer Timer2 
      Interval        =   250
      Left            =   2385
      Top             =   2175
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   2670
      ScaleHeight     =   540
      ScaleWidth      =   540
      TabIndex        =   29
      Top             =   1560
      Width           =   570
      Begin VB.PictureBox Picpic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   30
         ScaleHeight     =   450
         ScaleWidth      =   450
         TabIndex        =   30
         Top             =   30
         Width           =   480
      End
   End
   Begin VB.OptionButton optMale 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Male"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   150
      TabIndex        =   25
      Top             =   2610
      Value           =   -1  'True
      Width           =   210
   End
   Begin VB.OptionButton optFemale 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Male"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   975
      TabIndex        =   24
      Top             =   2610
      Width           =   210
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   960
      ScaleHeight     =   195
      ScaleWidth      =   210
      TabIndex        =   22
      Top             =   2580
      Width           =   240
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   15
         ScaleHeight     =   180
         ScaleWidth      =   180
         TabIndex        =   23
         Top             =   15
         Width           =   180
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   135
      ScaleHeight     =   195
      ScaleWidth      =   210
      TabIndex        =   19
      Top             =   2580
      Width           =   240
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   15
         ScaleHeight     =   180
         ScaleWidth      =   180
         TabIndex        =   20
         Top             =   15
         Width           =   180
      End
   End
   Begin VB.ComboBox cmbClass 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1950
      Width           =   2040
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   120
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1425
      Width           =   2025
   End
   Begin VB.Label picCancel 
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      Height          =   165
      Left            =   1575
      TabIndex        =   28
      Top             =   4035
      Width           =   525
   End
   Begin VB.Label picAddChar 
      BackStyle       =   0  'Transparent
      Caption         =   "Create Character"
      Height          =   210
      Left            =   1275
      TabIndex        =   27
      Top             =   3765
      Width           =   1215
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Female"
      Height          =   225
      Left            =   1230
      TabIndex        =   26
      Top             =   2595
      Width           =   585
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Male"
      Height          =   210
      Left            =   405
      TabIndex        =   21
      Top             =   2595
      Width           =   375
   End
   Begin VB.Label lblMAGI 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   3270
      TabIndex        =   18
      Top             =   3000
      Width           =   315
   End
   Begin VB.Label lblDEF 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   2235
      TabIndex        =   17
      Top             =   3405
      Width           =   375
   End
   Begin VB.Label lblSP 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   8760
      TabIndex        =   16
      Top             =   4380
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblSPEED 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   8760
      TabIndex        =   15
      Top             =   4740
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblMP 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   2235
      TabIndex        =   14
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label lblSTR 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   945
      TabIndex        =   13
      Top             =   3405
      Width           =   375
   End
   Begin VB.Label lblHP 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   945
      TabIndex        =   12
      Top             =   2985
      Width           =   375
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "SPEED"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   7920
      TabIndex        =   11
      Top             =   4740
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Magic:"
      ForeColor       =   &H00808080&
      Height          =   210
      Left            =   2685
      TabIndex        =   10
      Top             =   2985
      Width           =   480
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "SP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   7920
      TabIndex        =   9
      Top             =   4380
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Health:"
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   210
      TabIndex        =   8
      Top             =   2985
      Width           =   525
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Mana:"
      ForeColor       =   &H00808080&
      Height          =   210
      Left            =   1425
      TabIndex        =   7
      Top             =   2985
      Width           =   495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Defence:"
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   1425
      TabIndex        =   6
      Top             =   3405
      Width           =   660
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Strength:"
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   210
      TabIndex        =   5
      Top             =   3405
      Width           =   630
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Character Sex:"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   120
      TabIndex        =   4
      Top             =   2325
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label lbltype 
      BackStyle       =   0  'Transparent
      Caption         =   "Class Choice"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   135
      TabIndex        =   3
      Top             =   1725
      Width           =   1980
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Character Name:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   105
      TabIndex        =   1
      Top             =   1200
      Width           =   1200
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
    lblSPEED.Caption = STR(Class(cmbClass.ListIndex).SPEED)
    lblMAGI.Caption = STR(Class(cmbClass.ListIndex).MAGI)
End Sub

Private Sub Form_Load()

picSprites.Picture = LoadPicture(App.Path & "\core files\graphics\Sprites.bmp")
lbltype.Caption = ReadIniValue(App.Path & "\Core Files\Configuration.ini", "general", "choice")

Dim i As Long
Dim Ending As String
    For i = 1 To 3
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"

        If FileExist("\core files\interface\new character" & Ending) Then frmNewChar.Picture = LoadPicture(App.Path & "\core files\interface\new character" & Ending)
    Next i

End Sub

Private Sub Label2_Click()

End Sub

Private Sub picAddChar_Click()
Dim Msg As String
Dim i As Long

    If Trim(txtName.Text) <> "" Then
        Msg = Trim(txtName.Text)
        
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
    Call BitBlt(picPic.hdc, 0, 0, PIC_X, PIC_Y, picSprites.hdc, animi * PIC_X, Int(Class(cmbClass.ListIndex).MaleSprite) * PIC_Y, SRCCOPY)
Else
    Call BitBlt(picPic.hdc, 0, 0, PIC_X, PIC_Y, picSprites.hdc, animi * PIC_X, Int(Class(cmbClass.ListIndex).FemaleSprite) * PIC_Y, SRCCOPY)
End If
End Sub


Private Sub Timer2_Timer()
    animi = animi + 1
If animi > 4 Then
    animi = 3
End If
End Sub
