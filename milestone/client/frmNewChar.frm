VERSION 5.00
Begin VB.Form frmNewChar 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Character"
   ClientHeight    =   5985
   ClientLeft      =   135
   ClientTop       =   375
   ClientWidth     =   5985
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmNewChar.frx":0000
   ScaleHeight     =   399
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picsprites 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4755
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   18
      Top             =   4695
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
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00789298&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   840
      ScaleHeight     =   63
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   274
      TabIndex        =   3
      Top             =   3000
      Width           =   4140
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
         Left            =   3120
         TabIndex        =   27
         Top             =   675
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
         Left            =   3120
         TabIndex        =   26
         Top             =   450
         Value           =   -1  'True
         Width           =   735
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
         Left            =   240
         TabIndex        =   17
         Top             =   675
         Width           =   360
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
         Left            =   1230
         TabIndex        =   16
         Top             =   675
         Width           =   495
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
         Left            =   1230
         TabIndex        =   15
         Top             =   75
         Width           =   495
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
         Left            =   255
         TabIndex        =   14
         Top             =   75
         Width           =   240
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
         Left            =   2115
         TabIndex        =   13
         Top             =   90
         Width           =   300
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
         Left            =   3225
         TabIndex        =   12
         Top             =   105
         Width           =   615
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
         Left            =   2115
         TabIndex        =   11
         Top             =   675
         Width           =   540
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
         Left            =   510
         TabIndex        =   10
         Top             =   75
         Width           =   405
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
         Left            =   540
         TabIndex        =   9
         Top             =   675
         Width           =   375
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
         Left            =   1500
         TabIndex        =   8
         Top             =   90
         Width           =   375
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
         Left            =   2580
         TabIndex        =   7
         Top             =   690
         Width           =   375
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
         Left            =   2565
         TabIndex        =   6
         Top             =   105
         Width           =   375
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
         Left            =   1500
         TabIndex        =   5
         Top             =   675
         Width           =   375
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
         Left            =   3675
         TabIndex        =   4
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H0084ADB3&
      ForeColor       =   &H80000008&
      Height          =   1080
      Left            =   855
      ScaleHeight     =   1050
      ScaleWidth      =   4110
      TabIndex        =   22
      Top             =   1935
      Width           =   4140
      Begin VB.PictureBox Picpic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   3300
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   23
         Top             =   285
         Width           =   495
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   3270
         ScaleHeight     =   525
         ScaleWidth      =   525
         TabIndex        =   24
         Top             =   255
         Width           =   555
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Vocation"
         Height          =   210
         Left            =   45
         TabIndex        =   25
         Top             =   30
         Width           =   885
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Character Name"
      Height          =   210
      Left            =   405
      TabIndex        =   21
      Top             =   1230
      Width           =   1755
   End
   Begin VB.Label picAddChar 
      BackStyle       =   0  'Transparent
      Caption         =   "Create New Character"
      Height          =   240
      Left            =   2160
      TabIndex        =   20
      Top             =   4080
      Width           =   1665
   End
   Begin VB.Label picCancel 
      BackStyle       =   0  'Transparent
      Caption         =   "Back To Character Screen"
      Height          =   240
      Left            =   2235
      TabIndex        =   19
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
    lblSPEED.Caption = STR(Class(cmbClass.ListIndex).SPEED)
    lblMAGI.Caption = STR(Class(cmbClass.ListIndex).MAGI)
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
On Error Resume Next

If optMale.Value = True Then
    Call BitBlt(picPic.hdc, 0, 0, PIC_X, PIC_Y, picSprites.hdc, animi * PIC_X, Int(Class(cmbClass.ListIndex).MaleSprite) * PIC_Y, SRCCOPY)
Else
    Call BitBlt(picPic.hdc, 0, 0, PIC_X, PIC_Y, picSprites.hdc, animi * PIC_X, Int(Class(cmbClass.ListIndex).FemaleSprite) * PIC_Y, SRCCOPY)
End If
End Sub

Private Sub Form_Load()
    picSprites.Picture = LoadPicture(App.Path & "\GFX\sprites.bmp")
End Sub

Private Sub Timer2_Timer()
    animi = animi + 1
If animi > 4 Then
    animi = 3
End If
End Sub
