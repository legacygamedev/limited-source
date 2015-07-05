VERSION 5.00
Begin VB.Form frmParty 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Party"
   ClientHeight    =   4575
   ClientLeft      =   270
   ClientTop       =   75
   ClientWidth     =   1680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   305
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   112
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picRedCheck 
      Height          =   495
      Left            =   4320
      Picture         =   "frmParty.frx":0000
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   22
      Top             =   4680
      Width           =   1215
   End
   Begin VB.PictureBox picGreenCheck 
      Height          =   495
      Left            =   3120
      Picture         =   "frmParty.frx":0141
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   21
      Top             =   4680
      Width           =   1215
   End
   Begin VB.PictureBox picMPBar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   3
      Left            =   450
      Picture         =   "frmParty.frx":0271
      ScaleHeight     =   135
      ScaleWidth      =   630
      TabIndex        =   20
      Top             =   2505
      Width           =   630
   End
   Begin VB.PictureBox picMPBar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   2
      Left            =   450
      Picture         =   "frmParty.frx":03EE
      ScaleHeight     =   135
      ScaleWidth      =   630
      TabIndex        =   19
      Top             =   1710
      Width           =   630
   End
   Begin VB.PictureBox picMPBar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   1
      Left            =   450
      Picture         =   "frmParty.frx":056B
      ScaleHeight     =   135
      ScaleWidth      =   630
      TabIndex        =   18
      Top             =   870
      Width           =   630
   End
   Begin VB.PictureBox picMPBar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   4
      Left            =   450
      Picture         =   "frmParty.frx":06E8
      ScaleHeight     =   135
      ScaleWidth      =   630
      TabIndex        =   17
      Top             =   3345
      Width           =   630
   End
   Begin VB.PictureBox picHPBar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   4
      Left            =   450
      Picture         =   "frmParty.frx":0865
      ScaleHeight     =   135
      ScaleWidth      =   630
      TabIndex        =   16
      Top             =   3195
      Width           =   630
   End
   Begin VB.PictureBox picHPBar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   3
      Left            =   450
      Picture         =   "frmParty.frx":09EA
      ScaleHeight     =   135
      ScaleWidth      =   630
      TabIndex        =   15
      Top             =   2355
      Width           =   630
   End
   Begin VB.PictureBox picHPBar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   2
      Left            =   450
      Picture         =   "frmParty.frx":0B6F
      ScaleHeight     =   135
      ScaleWidth      =   630
      TabIndex        =   14
      Top             =   1560
      Width           =   630
   End
   Begin VB.PictureBox picHPBar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   1
      Left            =   450
      Picture         =   "frmParty.frx":0CF4
      ScaleHeight     =   135
      ScaleWidth      =   630
      TabIndex        =   13
      Top             =   720
      Width           =   630
   End
   Begin VB.PictureBox picSprites 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   3
      Left            =   4200
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   12
      Top             =   1.35000e5
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picSprites 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   2
      Left            =   3960
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   11
      Top             =   1.35000e5
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picSprites 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   1
      Left            =   3720
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   10
      Top             =   1.35000e5
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picItems 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   360
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   5
      Top             =   5400
      Width           =   240
   End
   Begin VB.PictureBox picSprites 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   4
      Left            =   3480
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   4
      Top             =   1.35000e5
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Timer tmrSprite 
      Interval        =   50
      Left            =   -240
      Top             =   0
   End
   Begin VB.Label lblLeader 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Party Led By:"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   240
      TabIndex        =   26
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "   Leave Party"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kick"
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
      Height          =   195
      Left            =   1200
      TabIndex        =   24
      Top             =   3960
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Invite "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   23
      Top             =   3960
      Width           =   555
   End
   Begin VB.Label Level 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   9
      Top             =   2925
      Width           =   585
   End
   Begin VB.Label Level 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   8
      Top             =   2085
      Width           =   585
   End
   Begin VB.Label Level 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   7
      Top             =   1290
      Width           =   585
   End
   Begin VB.Label Level 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   6
      Top             =   450
      Width           =   585
   End
   Begin VB.Label MemName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Blank Slot"
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   3
      Top             =   3540
      Width           =   1095
   End
   Begin VB.Label MemName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Blank Slot"
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   1905
      Width           =   1095
   End
   Begin VB.Label MemName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Blank Slot"
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   1
      Top             =   2700
      Width           =   1095
   End
   Begin VB.Label MemName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Blank Slot"
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   0
      Top             =   1065
      Width           =   1095
   End
End
Attribute VB_Name = "frmParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessage Lib "User32" _
                         Alias "SendMessageA" (ByVal hWnd As Long, _
                                               ByVal wMsg As Long, _
                                               ByVal wParam As Long, _
                                               lParam As Any) As Long

      Private Declare Sub ReleaseCapture Lib "User32" ()

      Const WM_NCLBUTTONDOWN = &HA1
      Const HTCAPTION = 2
Public Function Transparent(Form As Form, Layout As Byte) As Boolean
    SetWindowLong Form.hWnd, GWL_EXSTYLE, WS_EX_LAYERED
    SetLayeredWindowAttributes Form.hWnd, 0, Layout, LWA_ALPHA
    Transparent = Err.LastDllError = 0
End Function

Private Sub Form_Load()
    Me.Icon = frmMirage.Icon
    
    frmParty.picItems.Width = 480
    frmParty.picItems.Height = 720
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then

    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, _
                                 X As Single, Y As Single)
         Dim lngReturnValue As Long

         If Button = 1 Then
            Call ReleaseCapture
            lngReturnValue = SendMessage(frmParty.hWnd, WM_NCLBUTTONDOWN, _
                                         HTCAPTION, 0&)
         End If
      End Sub


Private Sub Label1_Click()
Call InvitePlayer
End Sub

Private Sub Label2_Click()
Call RemoveMember
End Sub

Private Sub Label3_Click()
Dim I As Byte
            Call SendLeaveParty
            If frmParty.Visible = True Then
            Unload frmParty
            For I = 1 To MAX_PARTY_INV_SLOTS
            'Player(MyIndex).Party.PartyItems(i).Num = 0
            Next I
            End If
            Unload Me
End Sub

Private Sub picItem_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Button = 1 Then
''Call SendRoll(Index)
'ElseIf Button = 2 Then
'Call SendNoRoll(Index)
'End If
'frmToW.Text1.SetFocus
End Sub

Private Sub picSprite_Click(Index As Integer)
'Call SendTarget(index)
'frmToW.Text1.SetFocus
End Sub

Private Sub tmrSprite_Timer()
'Call PartyBltSprite
End Sub

