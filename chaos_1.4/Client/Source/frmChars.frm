VERSION 5.00
Begin VB.Form frmChars 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Characters"
   ClientHeight    =   6975
   ClientLeft      =   120
   ClientTop       =   -45
   ClientWidth     =   9000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmChars.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "frmChars.frx":0CCA
   ScaleHeight     =   6975
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstChars2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   570
      ItemData        =   "frmChars.frx":106B5
      Left            =   3000
      List            =   "frmChars.frx":106B7
      TabIndex        =   21
      Top             =   2280
      Width           =   3585
   End
   Begin VB.PictureBox picCharsel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   1800
      Picture         =   "frmChars.frx":106B9
      ScaleHeight     =   191
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   399
      TabIndex        =   10
      Top             =   3000
      Width           =   6015
      Begin VB.PictureBox picChar1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   240
         ScaleHeight     =   78.961
         ScaleMode       =   0  'User
         ScaleWidth      =   65.376
         TabIndex        =   14
         Top             =   480
         Width           =   1455
      End
      Begin VB.PictureBox picSel1 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   120
         ScaleHeight     =   92.26
         ScaleMode       =   0  'User
         ScaleWidth      =   76.387
         TabIndex        =   16
         Top             =   360
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ListBox lstChars 
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
         Height          =   210
         ItemData        =   "frmChars.frx":4BFFF
         Left            =   5160
         List            =   "frmChars.frx":4C001
         TabIndex        =   15
         Top             =   2640
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.PictureBox picChar2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   2280
         ScaleHeight     =   78.961
         ScaleMode       =   0  'User
         ScaleWidth      =   65.376
         TabIndex        =   13
         Top             =   480
         Width           =   1455
      End
      Begin VB.PictureBox picChar3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   4320
         ScaleHeight     =   78.961
         ScaleMode       =   0  'User
         ScaleWidth      =   65.376
         TabIndex        =   12
         Top             =   480
         Width           =   1455
      End
      Begin VB.Timer Timer3 
         Interval        =   50
         Left            =   600
         Top             =   2640
      End
      Begin VB.PictureBox picSpriteloader 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   120
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   11
         Top             =   2640
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox picSel2 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   2160
         ScaleHeight     =   92.26
         ScaleMode       =   0  'User
         ScaleWidth      =   76.387
         TabIndex        =   17
         Top             =   360
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.PictureBox picSel3 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   4200
         ScaleHeight     =   92.26
         ScaleMode       =   0  'User
         ScaleWidth      =   76.387
         TabIndex        =   18
         Top             =   360
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label picUseChar 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Use Character or Create New"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   1560
         TabIndex        =   20
         Top             =   2160
         Width           =   3105
      End
      Begin VB.Label picDelChar 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Delete Character"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   2160
         TabIndex        =   19
         Top             =   2520
         Width           =   1740
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   250
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox Picsprites 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   9
      Top             =   6840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   6480
   End
   Begin VB.CommandButton Command1 
      Caption         =   "General Game Information"
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton btnPreOptions 
      Caption         =   "Pre-Login Option Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "3rd Party Programs Are Prohibited.  Such Programs are Subject to Termination or Deletion !"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   840
      TabIndex        =   22
      Top             =   6000
      Width           =   7755
   End
   Begin VB.Label lblAccountPassword 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Account Password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   2145
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Account Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   300
      Left            =   1680
      TabIndex        =   7
      Top             =   240
      Width           =   2310
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000C0&
      X1              =   120
      X2              =   5160
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000C0&
      X1              =   120
      X2              =   5160
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000C0&
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   2040
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      X1              =   5160
      X2              =   5160
      Y1              =   120
      Y2              =   2040
   End
   Begin VB.Label lblAccountName 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Account Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label lblRegister 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   270
      Left            =   120
      TabIndex        =   4
      Top             =   6600
      Width           =   1860
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exit Login Screen"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   7318
      TabIndex        =   2
      Top             =   6602
      Width           =   1695
   End
   Begin VB.Label Label30 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Days Remaining: Free Play"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Left            =   3360
      TabIndex        =   1
      Top             =   6240
      Width           =   3030
   End
   Begin VB.Label picCancel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   8520
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmChars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public animi As Long

Private Sub Form_Load()
Dim i As Long
Dim Ending As String
Dim sDc As Long
    For i = 1 To 3
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"
        
        If FileExist("Data\CharacterSelect" & Ending) Then frmChars.Picture = LoadPicture(App.Path & "\Data\CharacterSelect" & Ending)
 
    Next i
    
End Sub

Private Sub Label1_Click()
    Call TcpDestroy
    frmLogin.Visible = True
    Me.Visible = False
End Sub

Private Sub picCancel_Click()
    Call TcpDestroy
    frmLogin.Visible = True
    Me.Visible = False
End Sub

Private Sub picNewChar_Click()
    If lstChars.List(lstChars.ListIndex) <> "Free Character Slot" Then
        MsgBox "There is already a character in this slot!"
        Exit Sub
    End If
    Call MenuState(MENU_STATE_NEWCHAR)
End Sub

Private Sub picUseChar_Click()

    If lstChars.List(lstChars.ListIndex) = "Free Character Slot" Then
        Call MenuState(MENU_STATE_NEWCHAR)
    Else
        Call MenuState(MENU_STATE_USECHAR)
    End If
End Sub

Private Sub picDelChar_Click()
Dim Value As Long

    If lstChars.List(lstChars.ListIndex) = "Free Character Slot" Then
        MsgBox "There is no character in this slot!"
        Exit Sub
    End If

    Value = MsgBox("Are you sure you wish to delete this character?", vbYesNo, GAME_NAME)
    If Value = vbYes Then
        Call MenuState(MENU_STATE_DELCHAR)
    End If
End Sub

Private Sub Timer1_Timer()
Dim ACN
Dim BCN

ACN = frmLogin.txtName.Text
BCN = frmLogin.txtPassword.Text
    
    lblAccountName.Caption = "Account Name: " & ACN
    lblAccountPassword.Caption = "Password: " & BCN
End Sub

Private Sub Timer2_Timer()

    animi = animi + 1
    If animi > 4 Then
        animi = 3
    End If
    
End Sub

Private Sub Timer3_Timer()
    
If charselsprite(1) <> 0 Then
    Call BitBlt(picChar1.hDC, 30, 20, SIZE_X, SIZE_Y, picSpriteloader.hDC, animi * SIZE_X, charselsprite(1) * SIZE_Y, SRCCOPY)
End If
If charselsprite(2) <> 0 Then
    Call BitBlt(picChar2.hDC, 30, 20, SIZE_X, SIZE_Y, picSpriteloader.hDC, animi * SIZE_X, charselsprite(2) * SIZE_Y, SRCCOPY)
End If
If charselsprite(3) <> 0 Then
    Call BitBlt(picChar3.hDC, 30, 20, SIZE_X, SIZE_Y, picSpriteloader.hDC, animi * SIZE_X, charselsprite(3) * SIZE_Y, SRCCOPY)
End If
End Sub

Private Sub picChar1_Click()
    lstChars.ListIndex = 0
    lstChars2.ListIndex = 0
    picSel1.Visible = True
    picSel2.Visible = False
    picSel3.Visible = False
End Sub

Private Sub picChar2_Click()
    lstChars.ListIndex = 1
    lstChars2.ListIndex = 1
    picSel1.Visible = False
    picSel2.Visible = True
    picSel3.Visible = False
End Sub

Private Sub picChar3_Click()
    lstChars.ListIndex = 2
    lstChars2.ListIndex = 2
    picSel1.Visible = False
    picSel2.Visible = False
    picSel3.Visible = True
End Sub

Private Sub picCharsel_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    SOffsetX = x
    SOffsetY = y
End Sub

Private Sub picCharsel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'        Call MovePicture(frmChars.picCharsel, Button, Shift, x, y)
End Sub

