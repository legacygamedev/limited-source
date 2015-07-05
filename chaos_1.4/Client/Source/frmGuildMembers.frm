VERSION 5.00
Begin VB.Form frmGuildMembers 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Guild Member List"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmGuildMembers.frx":0000
   ScaleHeight     =   3675
   ScaleWidth      =   3630
   Begin VB.CommandButton Command2 
      Caption         =   "Done"
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin VB.ListBox lstMembers 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   2790
      ItemData        =   "frmGuildMembers.frx":10134
      Left            =   120
      List            =   "frmGuildMembers.frx":10136
      TabIndex        =   0
      Top             =   480
      Width           =   3405
   End
   Begin VB.Label lblGuild 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Guild:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   2820
   End
End
Attribute VB_Name = "frmGuildMembers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim i As Long
Dim Playerguildname As String

Private Sub Command2_Click()
    Timer1.Enabled = False
    Unload Me
    Call PlaySound("clk.wav")
End Sub

Private Sub Form_Load()
    Timer1.Enabled = True
    lblGuild.Caption = "Guild: " & GetPlayerGuild(MyIndex)
End Sub

Private Sub Timer1_Timer()
    lstMembers.Clear
    lstMembers.AddItem ("-Guild Members Online-")

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) Then
            If GetPlayerGuild(i) = GetPlayerGuild(MyIndex) Then
                If GetPlayerGuildAccess(i) <= 22 Then
                    lstMembers.AddItem (GetPlayerName(i) & " - " & frmMirage.lblLevel.Caption & "" & " - HP: " & GetPlayerHP(i) & "/" & GetPlayerMaxHP(i) & " - Rank: " & GetPlayerGuildAccess(i))
                End If
            End If
        End If

    Next i

End Sub
