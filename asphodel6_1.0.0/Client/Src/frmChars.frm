VERSION 5.00
Begin VB.Form frmChars 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asphodel Source (Characters)"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmChars.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   400
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   368
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrChars 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   5040
      Top             =   5520
   End
   Begin VB.PictureBox picCharBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   2
      Left            =   420
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   312
      TabIndex        =   6
      Top             =   2340
      Width           =   4680
      Begin VB.PictureBox picChar 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   2
         Left            =   360
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   8
         Top             =   540
         Width           =   480
      End
      Begin VB.Label lblCharDetails 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(empty)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   2
         Left            =   1080
         TabIndex        =   12
         Top             =   555
         Width           =   3135
      End
   End
   Begin VB.PictureBox picCharBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   3
      Left            =   420
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   312
      TabIndex        =   5
      Top             =   3825
      Width           =   4680
      Begin VB.PictureBox picChar 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   3
         Left            =   360
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   9
         Top             =   540
         Width           =   480
      End
      Begin VB.Label lblCharDetails 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(empty)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   3
         Left            =   1080
         TabIndex        =   11
         Top             =   555
         Width           =   3135
      End
   End
   Begin VB.PictureBox picCharBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   1
      Left            =   420
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   312
      TabIndex        =   4
      Top             =   855
      Width           =   4680
      Begin VB.PictureBox picChar 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   1
         Left            =   360
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   7
         Top             =   540
         Width           =   480
      End
      Begin VB.Label lblCharDetails 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(empty)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   1
         Left            =   1080
         TabIndex        =   10
         Top             =   555
         Width           =   3135
      End
   End
   Begin VB.Label lblCancel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   4155
      TabIndex        =   3
      Top             =   5355
      UseMnemonic     =   0   'False
      Width           =   945
   End
   Begin VB.Label lblDeleteChar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   2865
      TabIndex        =   2
      Top             =   5355
      Width           =   1095
   End
   Begin VB.Label lblNewChar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   1725
      TabIndex        =   1
      Top             =   5355
      Width           =   870
   End
   Begin VB.Label lblUseChar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   405
      TabIndex        =   0
      Top             =   5355
      Width           =   1035
   End
End
Attribute VB_Name = "frmChars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ------------------------------------------
' --               Asphodel               --
' ------------------------------------------

Private Anim As Long

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyUp
            If Char_Selected > 1 Then picCharBox_Click Char_Selected - 1
        Case vbKeyDown
            If Char_Selected < 3 Then picCharBox_Click Char_Selected + 1
        Case vbKeyReturn
            If lblCharDetails(Char_Selected).Caption <> "(blank)" Then
                lblUseChar_Click
            Else
                lblNewChar_Click
            End If
        Case vbKeyDelete
            lblDeleteChar_Click
        Case vbKeyEscape
            lblCancel_Click
    End Select
    
End Sub

Private Sub Form_Load()
    Me.Picture = LoadPicture(App.Path & GFX_PATH & "interface\selectcharwindow.bmp")
End Sub

Private Sub lblCharDetails_Click(Index As Integer)
    picCharBox_Click Index
End Sub

Private Sub lblUseChar_Click()
    MenuState Menu_State.UseChar_
End Sub

Private Sub lblNewChar_Click()

    If lblCharDetails(Char_Selected).Caption <> "(blank)" Then
        MsgBox "There is a character in that slot already!", vbOKOnly + vbCritical, Game_Name
        Exit Sub
    End If
    
    MenuState Menu_State.NewChar_
    
End Sub

Private Sub lblDeleteChar_Click()

    If lblCharDetails(Char_Selected).Caption = "(blank)" Then
        MsgBox "There isn't a character in that slot to delete!", vbOKOnly + vbCritical, Game_Name
        Exit Sub
    End If
    
    If MsgBox("Are you sure you wish the delete the selected character?", vbYesNo, Game_Name) = vbYes Then MenuState Menu_State.DelChar_
    
End Sub

Private Sub lblCancel_Click()

    SendData CLogout & END_CHAR
    
    CurrentWindow = Window_State.Main_Menu
    
    frmMainMenu.Visible = True
    Me.Visible = False
    
End Sub

Private Sub picChar_Click(Index As Integer)
    picCharBox_Click Index
End Sub

Private Sub picChar_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub picCharBox_Click(Index As Integer)

    If Index = Char_Selected Then Exit Sub
    
    DrawSelectedCharacter Char_Selected, GameConfig.StandFrame
    frmChars.picCharBox(Char_Selected).Picture = LoadPicture(App.Path & GFX_PATH & "Interface/char" & Char_Selected & "unselected.bmp")
    
    Anim = 0
    Char_Selected = Index
    frmChars.picCharBox(Char_Selected).Picture = LoadPicture(App.Path & GFX_PATH & "Interface/char" & Char_Selected & "selected.bmp")
    
End Sub

Private Sub picCharBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub tmrChars_Timer()

    Anim = Anim + 1
    If Anim > GameConfig.Total_WalkFrames Then Anim = 1
    
    DrawSelectedCharacter Char_Selected, GameConfig.WalkFrame(Anim)
    
End Sub
