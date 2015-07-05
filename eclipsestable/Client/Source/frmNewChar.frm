VERSION 5.00
Begin VB.Form frmNewChar 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Character"
   ClientHeight    =   5985
   ClientLeft      =   135
   ClientTop       =   315
   ClientWidth     =   5985
   ControlBox      =   0   'False
   Icon            =   "frmNewChar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   399
   ScaleMode       =   0  'User
   ScaleWidth      =   396.022
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1140
      Left            =   4560
      ScaleHeight     =   74
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   32
      Top             =   1680
      Visible         =   0   'False
      Width           =   540
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   15
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   37
         Top             =   720
         Width           =   480
         Begin VB.PictureBox iconn 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   2
            Left            =   0
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   38
            Top             =   -120
            Width           =   480
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   15
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   35
         Top             =   360
         Width           =   480
         Begin VB.PictureBox iconn 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   1
            Left            =   0
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   36
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.PictureBox picEmoticon 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   15
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   33
         Top             =   0
         Width           =   480
         Begin VB.PictureBox iconn 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   0
            Left            =   0
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   34
            Top             =   0
            Width           =   480
         End
      End
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   255
      Left            =   4080
      Max             =   200
      TabIndex        =   30
      Top             =   4680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   4080
      Max             =   200
      TabIndex        =   29
      Top             =   4080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   4080
      Max             =   200
      TabIndex        =   26
      Top             =   3480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.OptionButton optFemale 
      BackColor       =   &H00FFFFFF&
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
      Height          =   240
      Left            =   240
      MaskColor       =   &H00808080&
      TabIndex        =   18
      Top             =   3360
      Width           =   2040
   End
   Begin VB.OptionButton optMale 
      BackColor       =   &H00FFFFFF&
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
      Height          =   240
      Left            =   240
      MaskColor       =   &H00808080&
      TabIndex        =   17
      Top             =   3000
      Value           =   -1  'True
      Width           =   2040
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   4560
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   35
      TabIndex        =   7
      Top             =   2280
      Width           =   555
      Begin VB.PictureBox Picpic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   15
         ScaleHeight     =   31
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   31
         TabIndex        =   8
         Top             =   15
         Width           =   495
         Begin VB.PictureBox Picsprites 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            DrawStyle       =   5  'Transparent
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   120
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   25
            Top             =   120
            Width           =   495
         End
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   250
      Left            =   5760
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   5880
      Top             =   -120
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
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   975
      Width           =   2775
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
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
      Height          =   270
      Left            =   2160
      MaxLength       =   20
      TabIndex        =   0
      Top             =   420
      Width           =   3375
   End
   Begin VB.Label lblClassDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label13"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   1440
      TabIndex        =   39
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Label Label12 
      BackColor       =   &H00808080&
      Caption         =   "Legs:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4080
      TabIndex        =   31
      Top             =   4440
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label Label11 
      BackColor       =   &H00808080&
      Caption         =   "Body:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4080
      TabIndex        =   28
      Top             =   3840
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label Label14 
      BackColor       =   &H00808080&
      Caption         =   "Head:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4080
      TabIndex        =   27
      Top             =   3240
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label lblSPEED 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2781
      TabIndex        =   24
      Top             =   1800
      Width           =   480
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2880
      TabIndex        =   23
      Top             =   4920
      Width           =   600
   End
   Begin VB.Label lblDEF 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3480
      TabIndex        =   22
      Top             =   1800
      Width           =   600
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1560
      TabIndex        =   21
      Top             =   4800
      Width           =   600
   End
   Begin VB.Label lblSTR 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   720
      TabIndex        =   20
      Top             =   2400
      Width           =   480
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   19
      Top             =   3960
      Width           =   600
   End
   Begin VB.Label lblMAGI 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3340
      TabIndex        =   16
      Top             =   2400
      Width           =   480
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   15
      Top             =   4680
      Width           =   600
   End
   Begin VB.Label lblSP 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2055
      TabIndex        =   14
      Top             =   2400
      Width           =   480
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1680
      TabIndex        =   13
      Top             =   4440
      Width           =   600
   End
   Begin VB.Label lblMP 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1800
      TabIndex        =   12
      Top             =   1800
      Width           =   480
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1680
      TabIndex        =   11
      Top             =   4200
      Width           =   600
   End
   Begin VB.Label lblHP 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   756
      TabIndex        =   10
      Top             =   1800
      Width           =   480
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1680
      TabIndex        =   9
      Top             =   3960
      Width           =   600
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   240
      TabIndex        =   6
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label picAddChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   272
      TabIndex        =   4
      Top             =   4080
      Width           =   1005
   End
   Begin VB.Label picCancel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1783
      TabIndex        =   3
      Top             =   5280
      Width           =   2445
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
      Left            =   6840
      TabIndex        =   2
      Top             =   2160
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
    Dim i As Byte
    
    For i = 0 To Max_Classes
        If Trim(Class(i).name) = cmbClass.List(cmbClass.ListIndex) Then
            Exit For
        End If
    Next
    
    lblHP.Caption = STR(Class(i).HP)
    lblMP.Caption = STR(Class(i).MP)
    lblSP.Caption = STR(Class(i).SP)

    lblSTR.Caption = STR(Class(i).STR)
    lblDEF.Caption = STR(Class(i).DEF)
    lblSPEED.Caption = STR(Class(i).speed)
    lblMAGI.Caption = STR(Class(i).MAGI)

    lblClassDesc.Caption = Class(i).desc
End Sub


Private Sub HScroll1_Change()
    If SpriteSize = 1 Then
        iconn(0).Top = -Val(HScroll1.Value * 64 + 15)
    Else
        iconn(0).Top = -Val(HScroll1.Value * PIC_Y)
    End If
End Sub

Private Sub HScroll2_Change()
    If SpriteSize = 1 Then
        iconn(1).Top = -Val(HScroll2.Value * 64 + 25)
    Else
        iconn(1).Top = -Val(HScroll2.Value * PIC_Y)
    End If
End Sub

Private Sub HScroll3_Change()
    If SpriteSize = 1 Then
        iconn(2).Top = -Val(HScroll3.Value * 64 + 35)
    Else
        iconn(2).Top = -Val(HScroll3.Value * PIC_Y)
    End If
End Sub


Private Sub picAddChar_Click()
    Dim Msg As String
    Dim i As Long

    If Trim$(txtName.Text) <> vbNullString Then
        Msg = Trim$(txtName.Text)

        If Len(Trim$(txtName.Text)) < 3 Then
            MsgBox "Character name must be at least three characters in length."
            Exit Sub
        End If

        ' Prevent high ascii chars
        For i = 1 To Len(Msg)
            If Asc(Mid$(Msg, i, 1)) < 32 Or Asc(Mid$(Msg, i, 1)) > 255 Then
                Call MsgBox("You cannot use high ascii chars in your name, please reenter.", vbOKOnly, GAME_NAME)
                txtName.Text = vbNullString
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

    If cmbClass.ListIndex < 0 Then
        Exit Sub
    End If
    If 0 + CustomPlayers = 0 Then
        If SpriteSize = 1 Then
            If optMale.Value = True Then
                frmNewChar.picSprites.Left = (animi * PIC_X) * -1
                frmNewChar.picSprites.Top = (Int(Class(cmbClass.ListIndex).MaleSprite) * 64) * -1
                Call BitBlt(picPic.hDC, 0, 0, PIC_X, 64, picSprites.hDC, animi * PIC_X, Int(Class(cmbClass.ListIndex).MaleSprite) * 64, SRCCOPY)
            Else
                frmNewChar.picSprites.Left = (animi * PIC_X) * -1
                frmNewChar.picSprites.Top = (Int(Class(cmbClass.ListIndex).FemaleSprite) * 64) * -1
                Call BitBlt(picPic.hDC, 0, 0, PIC_X, 64, picSprites.hDC, animi * PIC_X, Int(Class(cmbClass.ListIndex).FemaleSprite) * 64, SRCCOPY)
            End If
        Else
            If optMale.Value = True Then
                frmNewChar.picSprites.Left = (animi * PIC_X) * -1
                frmNewChar.picSprites.Top = (Int(Class(cmbClass.ListIndex).MaleSprite) * PIC_Y) * -1

                Call BitBlt(picPic.hDC, 0, 0, PIC_X, PIC_Y, picSprites.hDC, animi * PIC_X, Int(Class(cmbClass.ListIndex).MaleSprite) * PIC_Y, SRCCOPY)
            Else
                frmNewChar.picSprites.Left = (animi * PIC_X) * -1
                frmNewChar.picSprites.Top = (Int(Class(cmbClass.ListIndex).FemaleSprite) * PIC_Y) * -1
                Call BitBlt(picPic.hDC, 0, 0, PIC_X, PIC_Y, picSprites.hDC, animi * PIC_X, Int(Class(cmbClass.ListIndex).FemaleSprite) * PIC_Y, SRCCOPY)
            End If
        End If
    End If

End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim Ending As String

    For i = 1 To 3
        If i = 1 Then
            Ending = ".gif"
        End If
        If i = 2 Then
            Ending = ".jpg"
        End If
        If i = 3 Then
            Ending = ".png"
        End If

        If FileExists("GUI\NewCharacter" & Ending) Then
            frmNewChar.Picture = LoadPicture(App.Path & "\GUI\NewCharacter" & Ending)
        End If
    Next i

    If CustomPlayers = 1 Then
        If FileExists("GFX\Sprites.bmp") Then
            picSprites.Picture = LoadPicture(App.Path & "\GFX\Sprites.bmp")
        Else
            Call MsgBox("Error: Could not find Sprites.bmp.")
            End
        End If
    End If

' Set the size of the scrolling bars
' HScroll1.Max = LoadPicture(App.Path & "\GFX\Heads.bmp").Height / 64
' DOES NOT WORK, FIX LATER PLZ KTHX

End Sub

Private Sub Timer2_Timer()
    animi = animi + 1
    If animi > 4 Then
        animi = 3
    End If
End Sub
