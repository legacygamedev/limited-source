VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMenu 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7725
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   352
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   515
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock Socket 
      Left            =   0
      Top             =   1155
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrUpdateNews 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox picLogin 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3630
      Left            =   555
      ScaleHeight     =   242
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   442
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   180
      Visible         =   0   'False
      Width           =   6630
      Begin VB.CheckBox chkUsername 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Remember"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4080
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   2160
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.TextBox txtLPass 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         IMEMode         =   3  'DISABLE
         Left            =   2520
         PasswordChar    =   "•"
         TabIndex        =   2
         Top             =   1800
         Width           =   2775
      End
      Begin VB.TextBox txtLUser 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   2520
         TabIndex        =   1
         Top             =   1440
         Width           =   2775
      End
      Begin VB.Label lblLAccept 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Accept"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   3
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label lblBlank 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   1320
         TabIndex        =   6
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblBlank 
         BackStyle       =   0  'Transparent
         Caption         =   "Username:"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   5
         Top             =   1440
         Width           =   1215
      End
   End
   Begin VB.PictureBox picCredits 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   3645
      Left            =   540
      ScaleHeight     =   3645
      ScaleWidth      =   6630
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   180
      Width           =   6630
   End
   Begin VB.PictureBox picMain 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   3645
      Left            =   555
      ScaleHeight     =   3645
      ScaleWidth      =   6630
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   180
      Width           =   6630
      Begin VB.Label lblServerStatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Offline"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   0
         TabIndex        =   26
         Top             =   3320
         Visible         =   0   'False
         Width           =   6615
      End
      Begin VB.Label lblNews 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Connecting to server..."
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1575
         Left            =   1680
         TabIndex        =   24
         Top             =   1200
         Width           =   3135
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picCharacter 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   3645
      Left            =   555
      ScaleHeight     =   243
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   442
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   180
      Visible         =   0   'False
      Width           =   6630
      Begin VB.OptionButton optMale 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Male"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2280
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   2040
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.PictureBox picSprite 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Left            =   4680
         ScaleHeight     =   48
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1320
         Width           =   480
      End
      Begin VB.ComboBox cmbClass 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1560
         Width           =   2175
      End
      Begin VB.OptionButton optFemale 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Female"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3360
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox txtCUser 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2280
         MaxLength       =   12
         TabIndex        =   0
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label lblBlank 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Gender:"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   1080
         TabIndex        =   21
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lblBlank 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Class:"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   1440
         TabIndex        =   20
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label lblBlank 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   19
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblCAccept 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Accept"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2280
         TabIndex        =   16
         Top             =   2520
         Width           =   2175
      End
   End
   Begin VB.PictureBox picRegister 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   3645
      Left            =   555
      ScaleHeight     =   243
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   442
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   180
      Visible         =   0   'False
      Width           =   6630
      Begin VB.TextBox txtRPass2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         IMEMode         =   3  'DISABLE
         Left            =   2520
         PasswordChar    =   "•"
         TabIndex        =   13
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox txtRPass 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         IMEMode         =   3  'DISABLE
         Left            =   2520
         PasswordChar    =   "•"
         TabIndex        =   10
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox txtRUser 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2520
         TabIndex        =   8
         Top             =   1320
         Width           =   2775
      End
      Begin VB.Label lblBlank 
         BackStyle       =   0  'Transparent
         Caption         =   "Retype:"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   11
         Left            =   1320
         TabIndex        =   14
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label txtRAccept 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Accept"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   12
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label lblBlank 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   9
         Left            =   1320
         TabIndex        =   11
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label lblBlank 
         BackStyle       =   0  'Transparent
         Caption         =   "Username:"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   8
         Left            =   1320
         TabIndex        =   9
         Top             =   1320
         Width           =   1215
      End
   End
   Begin VB.PictureBox picChars 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3630
      Left            =   540
      ScaleHeight     =   242
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   442
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   180
      Visible         =   0   'False
      Width           =   6630
      Begin VB.Label lblDelChar 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Delete Character"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2640
         TabIndex        =   31
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label lblAddChar 
         BackStyle       =   0  'Transparent
         Caption         =   "Add Character"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   30
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label lblUseChar 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Use Character"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   29
         Top             =   2520
         Width           =   1455
      End
   End
   Begin VB.Image ImgButton 
      Height          =   435
      Index           =   4
      Left            =   5460
      Top             =   4305
      Width           =   1335
   End
   Begin VB.Image ImgButton 
      Height          =   435
      Index           =   3
      Left            =   3960
      Top             =   4305
      Width           =   1335
   End
   Begin VB.Image ImgButton 
      Height          =   435
      Index           =   2
      Left            =   2460
      Top             =   4305
      Width           =   1335
   End
   Begin VB.Image ImgButton 
      Height          =   435
      Index           =   1
      Left            =   960
      Top             =   4305
      Width           =   1335
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbClass_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    NewCharClass = ClassSelection(cmbClass.ListIndex + 1)
    Menu_DrawCharacter
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmbClass_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_Load()
    Dim TmpTxt As String, TmpArray() As String, I As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    ' General menu things
    Me.Caption = ""
    
    ' Set the maxes for the username and password
    txtLUser.MaxLength = NAME_LENGTH
    txtLPass.MaxLength = NAME_LENGTH
    txtCUser.MaxLength = NAME_LENGTH
    txtRUser.MaxLength = NAME_LENGTH
    txtRPass.MaxLength = NAME_LENGTH
    txtRPass2.MaxLength = NAME_LENGTH
    
    ' Load the username
    If Options.SaveUsername = 1 Then
        txtLUser.text = Trim$(Options.UserName)
        txtLUser.SelStart = Len(Trim$(Options.UserName))
        chkUsername.Value = Options.SaveUsername
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Form_Load", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    LastButton_Menu = 0
    ResetMenuButtons
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Form_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Not EnteringGame And Not gameDestroyed Then DestroyGame
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Form_Unload", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub ImgButton_Click(Index As Integer)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Not CurButton_Menu = Index Then
        Call Audio.PlaySound(ButtonClick)
        CurButton_Menu = Index
        MenuButton(Index).State = 2
        Call RenderButton_Menu(Index)
        Call ResetMenuButtons
            
        Select Case Index
            Case 1
                If Not picLogin.Visible Then
                    picCredits.Visible = False
                    picLogin.Visible = True
                    picRegister.Visible = False
                    picCharacter.Visible = False
                    picMain.Visible = False
                    If Options.SaveUsername = 0 Then
                        txtLUser.SetFocus
                    Else
                        txtLPass.SetFocus
                    End If
                End If
            Case 2
                If Not picRegister.Visible Then
                    picCredits.Visible = False
                    picLogin.Visible = False
                    picRegister.Visible = True
                    picCharacter.Visible = False
                    picMain.Visible = False
                    txtRUser.SetFocus
                End If
            Case 3
                If Not picCredits.Visible Then
                    picCredits.Visible = True
                    picLogin.Visible = False
                    picRegister.Visible = False
                    picCharacter.Visible = False
                    picMain.Visible = False
                End If
            Case 4
                Call DestroyGame
        End Select
    
        ' Reset all buttons
        ResetMenuButtons
        CurButton_Menu = Index
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ImgButton_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub ImgButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Not LastButton_Menu = Index And Not CurButton_Menu = Index Then
        ResetMenuButtons
        MenuButton(Index).State = 1
        Call RenderButton_Menu(Index)
        Call Audio.PlaySound(ButtonHover)
        LastButton_Menu = Index
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ImgButton_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lblLAccept_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If IsLoginLegal(txtLUser.text, txtLPass.text) Then
        Call MenuState(MENU_STATE_LOGIN)
    End If
    
    ' Save options
    Options.SaveUsername = frmMenu.chkUsername.Value
    Options.UserName = Trim$(frmMenu.txtLUser.text)
    
    If frmMenu.chkUsername.Value = 0 Then
        Options.UserName = vbNullString
    Else
        Options.UserName = Trim$(frmMenu.txtLUser.text)
    End If
    
    SaveOptions
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "lblLAccept_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub optFemale_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    NewCharClass = ClassSelection(cmbClass.ListIndex + 1)
    Menu_DrawCharacter
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "optFemale_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub optFemale_GotFocus()
    txtCUser.SetFocus
    
    If optFemale.Value = False Then
        optFemale.Value = True
    End If
End Sub

Private Sub optMale_GotFocus()
    txtCUser.SetFocus
    
    If optMale.Value = False Then
        optMale.Value = True
    End If
End Sub

Private Sub optMale_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    NewCharClass = ClassSelection(cmbClass.ListIndex + 1)
    Menu_DrawCharacter
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "optMale_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picCharacter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ResetMenuButtons
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picCharacter_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picCredits_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ResetMenuButtons
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picCredits_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picLogin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ResetMenuButtons
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picLogin_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ResetMenuButtons
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picMain_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picRegister_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ResetMenuButtons
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picRegister_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

' Winsock event
Private Sub Socket_DataArrival(ByVal bytesTotal As Long)

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If IsConnected Then
        Call IncomingData(bytesTotal)
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Socket_DataArrival", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub tmrUpdateNews_Timer()
    ' Don't update if menu is not visible
    If frmMenu.Visible = False Then Exit Sub
    
    ' If we're connected we don't need to update anything
    If IsConnected Then Exit Sub
    
    ' If the the timer is paused but we're not connected clear it
    If StopTimer = True And IsConnected = False Then
        StopTimer = False
    End If
    
    ' Check if the timer is disabled
    If StopTimer Then Exit Sub
    
    If ConnectToServer(1) Then
        Call UpdateData
        StopTimer = True
    End If
    
    If IsConnected = False Then
        frmMenu.lblServerStatus.Caption = "Offline"
        frmMenu.lblServerStatus.ForeColor = vbRed
        frmMenu.lblNews.Caption = "The server appears to be offline. Please try connecting again later."
        frmMenu.lblServerStatus.Visible = True
    End If
End Sub

' Register
Private Sub txtRAccept_Click()
    Dim Name As String
    Dim Password As String
    Dim PasswordAgain As String, RndCharacters As String
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Name = Trim$(txtRUser.text)
    Password = Trim$(txtRPass.text)
    PasswordAgain = Trim$(txtRPass2.text)

    If IsLoginLegal(Name, Password) Then
        If IsLoginLegal(Name, PasswordAgain) Then
            If Not Password = PasswordAgain Then
                Call AlertMsg("Passwords don't match.")
                Exit Sub
            End If
    
            If Not IsStringLegal(Name) Then Exit Sub
    
            Call MenuState(MENU_STATE_NEWACCOUNT)
        End If
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtRAccept_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

' New Character
Private Sub lblCAccept_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If IsNewCharLegal(txtCUser) Then
        Call MenuState(MENU_STATE_ADDCHAR)
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "lblCAccept_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If picLogin.Visible = True And lblLAccept.Enabled Then
            lblLAccept_Click
        ElseIf picCharacter.Visible And lblCAccept.Enabled Then
            lblCAccept_Click
        ElseIf picRegister.Visible And txtRAccept.Enabled Then
            txtRAccept_Click
        ElseIf picMain.Visible Then
            Call ImgButton_Click(1)
        End If
        KeyAscii = 0
    End If
    
    If KeyAscii = vbKeyEscape Then
        DestroyGame
        KeyAscii = 0
    End If
End Sub
