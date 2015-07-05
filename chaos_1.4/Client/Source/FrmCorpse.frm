VERSION 5.00
Begin VB.Form FrmCorpse 
   BorderStyle     =   0  'None
   Caption         =   "Corpse "
   ClientHeight    =   3705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3405
   ForeColor       =   &H000000C0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmCorpse.frx":0000
   ScaleHeight     =   3705
   ScaleWidth      =   3405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picItems 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   120
      ScaleHeight     =   480
      ScaleWidth      =   465
      TabIndex        =   5
      Top             =   3600
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Timer TmrBlt 
      Interval        =   50
      Left            =   120
      Top             =   3600
   End
   Begin VB.PictureBox PicLoot 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   3
      Left            =   360
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   3
      Top             =   2520
      Width           =   480
   End
   Begin VB.PictureBox PicLoot 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   2
      Left            =   360
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   2
      Top             =   1920
      Width           =   480
   End
   Begin VB.PictureBox PicLoot 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   0
      Left            =   360
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   1
      Top             =   720
      Width           =   480
   End
   Begin VB.PictureBox PicLoot 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   1
      Left            =   360
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   1320
      Width           =   480
   End
   Begin VB.Label pic 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Decaying Corpse"
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
      Left            =   720
      TabIndex        =   10
      Top             =   240
      Width           =   1965
   End
   Begin VB.Label LblItemName 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   3
      Left            =   960
      TabIndex        =   9
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label LblItemName 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   2
      Left            =   960
      TabIndex        =   8
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label LblItemName 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   7
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label LblItemName 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   6
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Lblexit 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "   Exit Corpse"
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   3240
      Width           =   1095
   End
End
Attribute VB_Name = "FrmCorpse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' XCORPSEX
Option Explicit
Private Sub Form_LostFocus()
Unload Me
End Sub

Private Sub Lblexit_Click()
Unload Me
End Sub

Private Sub PicLoot_Click(Index As Integer)
If IsPlaying(CorpseIndex) Then
  If Player(CorpseIndex).CorpseLoot(Index + 1).Num > 0 Then
  Call SendData("takecorpseloot" & SEP_CHAR & CorpseIndex & SEP_CHAR & (Index + 1) & SEP_CHAR & END_CHAR)
  End If
End If
End Sub

Private Sub TmrBlt_Timer()
Dim pic As Long
Dim i As Integer
Dim sDc As Long

    If CorpseIndex = 0 Then Exit Sub
    
    For i = 1 To 4
        If Player(CorpseIndex).CorpseLoot(i).Num > 0 Then
            pic = STR(Item(Player(CorpseIndex).CorpseLoot(i).Num).pic)
            'Call BitBlt(PicLoot(i - 1).hDC, 0, 0, PIC_X, PIC_Y, picItems.hDC, (pic - Int(pic / 6) * 6) * PIC_X, Int(pic / 6) * PIC_Y, SRCCOPY)
            PicLoot(i - 1).Cls
            sDc = DD_ItemSurf.GetDC
            Call BitBlt(PicLoot(i - 1).hDC, 0, 0, PIC_X, PIC_Y, sDc, (Item(Player(CorpseIndex).CorpseLoot(i).Num).pic - Int(Item(Player(CorpseIndex).CorpseLoot(i).Num).pic / 6) * 6) * PIC_X, Int(Item(Player(CorpseIndex).CorpseLoot(i).Num).pic / 6) * PIC_Y, SRCCOPY)
            Call DD_ItemSurf.ReleaseDC(sDc)
        Else
            PicLoot(i - 1).Cls
        End If
    Next i
End Sub
' XCORPSEX
