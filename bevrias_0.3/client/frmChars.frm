VERSION 5.00
Begin VB.Form frmChars 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Characters"
   ClientHeight    =   5250
   ClientLeft      =   165
   ClientTop       =   390
   ClientWidth     =   7500
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmChars.frx":0000
   ScaleHeight     =   5250
   ScaleWidth      =   7500
   StartUpPosition =   2  'CenterScreen
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
      Height          =   1290
      ItemData        =   "frmChars.frx":2685
      Left            =   840
      List            =   "frmChars.frx":2687
      TabIndex        =   0
      Top             =   1320
      Width           =   4690
   End
   Begin VB.Label CharStat10 
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
      Height          =   255
      Left            =   4680
      TabIndex        =   16
      Top             =   4350
      Width           =   975
   End
   Begin VB.Label CharStat9 
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
      Height          =   255
      Left            =   1440
      TabIndex        =   15
      Top             =   3720
      Width           =   2535
   End
   Begin VB.Label CharStat8 
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
      Height          =   255
      Left            =   1440
      TabIndex        =   14
      Top             =   3480
      Width           =   2535
   End
   Begin VB.Label CharStat5 
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
      Height          =   255
      Left            =   1680
      TabIndex        =   13
      Top             =   4560
      Width           =   2295
   End
   Begin VB.Label CharStat2 
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
      Height          =   255
      Left            =   5040
      TabIndex        =   12
      Top             =   3700
      Width           =   735
   End
   Begin VB.Label CharStats 
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
      Height          =   255
      Left            =   1440
      TabIndex        =   11
      Top             =   3270
      Width           =   2535
   End
   Begin VB.Label CharStat 
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
      Height          =   255
      Left            =   5040
      TabIndex        =   10
      Top             =   3270
      Width           =   735
   End
   Begin VB.Label CharStat1 
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
      Height          =   255
      Left            =   5040
      TabIndex        =   9
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label CharStat3 
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
      Height          =   255
      Left            =   5040
      TabIndex        =   8
      Top             =   3920
      Width           =   735
   End
   Begin VB.Label CharStat4 
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
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   4125
      Width           =   2535
   End
   Begin VB.Label CharStat6 
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
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   4335
      Width           =   2415
   End
   Begin VB.Label CharStat7 
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
      Height          =   240
      Left            =   1440
      TabIndex        =   5
      Top             =   3930
      Width           =   1695
   End
   Begin VB.Label picUseChar 
      BackStyle       =   0  'Transparent
      Height          =   345
      Left            =   840
      TabIndex        =   4
      Top             =   2520
      Width           =   1425
   End
   Begin VB.Label picNewChar 
      BackStyle       =   0  'Transparent
      Height          =   360
      Left            =   3960
      TabIndex        =   3
      Top             =   2520
      Width           =   1605
   End
   Begin VB.Label picCancel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   6120
      TabIndex        =   2
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label picDelChar 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   2520
      Width           =   1740
   End
End
Attribute VB_Name = "frmChars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim i As Long
Dim Ending As String
    For i = 1 To 3
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"
 
        If FileExist("GUI\CharacterSelect" & Ending) Then frmChars.Picture = LoadPicture(App.Path & "\GUI\CharacterSelect" & Ending)
    Next i
End Sub

Private Sub lstChars_Click()
Dim Ind As Long
Dim MyInd As Long
Dim i As Long
CharStats.Caption = ""
CharStat.Caption = ""
CharStat1.Caption = ""
CharStat2.Caption = ""
CharStat3.Caption = ""
i = lstChars.ListIndex + 1
MyInd = lstChars.ListCount
Call CharStas(i)
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
    
    Dim Packet As String
    Packet = "charwindow" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
    Call MenuState(MENU_STATE_NEWCHAR)
End Sub

Private Sub picUseChar_Click()
    If lstChars.List(lstChars.ListIndex) = "Free Character Slot" Then
        MsgBox "There is no character in this slot!"
        Exit Sub
    End If
    Call MenuState(MENU_STATE_USECHAR)
If ReadINI("STORY", "DontShowAgain", App.Path & "\config.ini") = 0 Then
frmMirage.Picture29.Visible = True
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
