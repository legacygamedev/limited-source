VERSION 5.00
Begin VB.Form frmNewChar 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Character"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2805
   Icon            =   "frmNewChar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   230
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   187
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optFemale 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Male"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   255
      TabIndex        =   8
      Top             =   1590
      Width           =   210
   End
   Begin VB.OptionButton optMale 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Male"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   255
      TabIndex        =   7
      Top             =   1350
      Value           =   -1  'True
      Width           =   210
   End
   Begin VB.ListBox cmbclass 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   570
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   2400
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   1080
      ScaleHeight     =   540
      ScaleWidth      =   540
      TabIndex        =   3
      Top             =   1920
      Width           =   570
      Begin VB.PictureBox Picpic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   25
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   13
         Top             =   25
         Width           =   480
      End
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   240
      MaxLength       =   20
      TabIndex        =   0
      Top             =   360
      Width           =   2400
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   240
      ScaleHeight     =   195
      ScaleWidth      =   210
      TabIndex        =   10
      Top             =   1320
      Width           =   240
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   240
      ScaleHeight     =   195
      ScaleWidth      =   210
      TabIndex        =   9
      Top             =   1560
      Width           =   240
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Character Name"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Male"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   510
      TabIndex        =   12
      Top             =   1335
      Width           =   375
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Female"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   510
      TabIndex        =   11
      Top             =   1575
      Width           =   585
   End
   Begin VB.Label lblPrevious 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "<---"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1920
      Width           =   600
   End
   Begin VB.Label lblNext 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "--->"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   1920
      Width           =   600
   End
   Begin VB.Label picCancel 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Back"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label picAddChar 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Create Character"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   2640
      Width           =   1350
   End
End
Attribute VB_Name = "frmNewChar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    DrawSprite
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmChars.Visible = True
    Me.Visible = False
End Sub

Private Sub cmbClass_Click()
    
    spriteMale() = Split(Class(cmbclass.ListIndex).MaleSprite, ",")
    spriteFemale() = Split(Class(cmbclass.ListIndex).FemaleSprite, ",")
    
    currSpriteNum = 0
    
    DrawSprite
End Sub

Private Sub cmbclass_GotFocus()
    txtName.SetFocus
End Sub

Private Sub optFemale_Click()
    DrawSprite
    txtName.SetFocus
End Sub

Private Sub optMale_Click()
    DrawSprite
    txtName.SetFocus
End Sub

Private Sub picAddChar_Click()
Dim Msg As String
Dim i As Long

    If Trim$(txtName.Text) <> vbNullString Then
        Msg = Trim$(txtName.Text)
        
        ' Prevent high ascii chars
        For i = 1 To Len(Msg)
            If Asc(Mid$(Msg, i, 1)) < 32 Or Asc(Mid$(Msg, i, 1)) > 126 Then
                Call TcpDestroy
                frmEvent.Visible = True
                frmEvent.lblInformation.Caption = "Adventurer name contains invalid characters. (Err: #3)"
                txtName.Text = vbNullString
                Exit Sub
            End If
        Next
        
        frmNewChar.Visible = False
        If ConnectToServer = True Then
            Call SetStatus("Connected, sending adventurer addition...")
            If frmNewChar.optMale.Value = True Then
                Call SendAddChar(frmNewChar.txtName, 0, frmNewChar.cmbclass.ListIndex, frmChars.lstChars.ListIndex + 1, spriteMale(currSpriteNum))
            Else
                Call SendAddChar(frmNewChar.txtName, 1, frmNewChar.cmbclass.ListIndex, frmChars.lstChars.ListIndex + 1, spriteFemale(currSpriteNum))
            End If
        End If
    End If
End Sub

Private Sub picCancel_Click()
    frmChars.Visible = True
    Me.Visible = False
End Sub

Private Sub lblNext_Click()
    currSpriteNum = currSpriteNum + 1
    
    If optMale.Value = True Then
        If currSpriteNum > UBound(spriteMale) Then currSpriteNum = 0
    Else
        If currSpriteNum > UBound(spriteFemale) Then currSpriteNum = 0
    End If
    
    DrawSprite
    txtName.SetFocus
End Sub

Private Sub lblPrevious_Click()
    currSpriteNum = currSpriteNum - 1
    
    If optMale.Value = True Then
        If currSpriteNum < 0 Then currSpriteNum = UBound(spriteMale)
    Else
        If currSpriteNum < 0 Then currSpriteNum = UBound(spriteFemale)
    End If
    
    DrawSprite
    txtName.SetFocus
End Sub

Private Sub DrawSprite()
Dim rec As RECT
Dim rec_pos As RECT
    
    If cmbclass.ListIndex < 0 Then Exit Sub
    
    With rec
        If optMale.Value = True Then
            .Top = CLng(spriteMale(currSpriteNum)) * PIC_Y
        Else
            .Top = CLng(spriteFemale(currSpriteNum)) * PIC_Y
        End If
        .Bottom = .Top + PIC_Y
        .Left = 4 * PIC_X
        .Right = .Left + PIC_X
    End With
    
    With rec_pos
        .Top = 0
        .Bottom = PIC_Y
        .Left = 0
        .Right = PIC_X
    End With
    
    DD_SpriteSurf.BltToDC Picpic.hdc, rec, rec_pos
    Picpic.Refresh
End Sub
