VERSION 5.00
Begin VB.Form frmAnimationEditor 
   Caption         =   "Form1"
   ClientHeight    =   3315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   221
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   439
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   120
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   360
      Width           =   4815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   5520
      TabIndex        =   15
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   255
      Left            =   4440
      TabIndex        =   14
      Top             =   3000
      Width           =   975
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   5040
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   13
      Top             =   120
      Width           =   480
   End
   Begin VB.Timer tmrAnimation 
      Interval        =   5
      Left            =   6120
      Top             =   120
   End
   Begin VB.Frame Frame1 
      Caption         =   "Animations"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   6375
      Begin VB.OptionButton optAbove 
         Caption         =   "Above Npcs + Players"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   105
         TabIndex        =   19
         Top             =   1485
         Width           =   1935
      End
      Begin VB.OptionButton optBelow 
         Caption         =   "Below Npcs + Players"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   105
         TabIndex        =   18
         Top             =   1245
         Width           =   1935
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Left            =   2160
         Max             =   255
         TabIndex        =   6
         Top             =   360
         Width           =   3495
      End
      Begin VB.HScrollBar scrlFrames 
         Height          =   255
         Left            =   2160
         Max             =   25
         TabIndex        =   5
         Top             =   645
         Value           =   1
         Width           =   3495
      End
      Begin VB.HScrollBar scrlSpeed 
         Height          =   255
         Left            =   2160
         Max             =   255
         TabIndex        =   4
         Top             =   930
         Value           =   1
         Width           =   3495
      End
      Begin VB.Frame Frame2 
         Height          =   975
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1095
         Begin VB.OptionButton opt32 
            Caption         =   "32x32"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   105
            TabIndex        =   3
            Top             =   285
            Width           =   855
         End
         Begin VB.OptionButton opt64 
            Caption         =   "64x64"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   105
            TabIndex        =   2
            Top             =   525
            Width           =   855
         End
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Animation"
         Height          =   255
         Left            =   1330
         TabIndex        =   12
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblSprite 
         Caption         =   "0"
         Height          =   255
         Left            =   5760
         TabIndex        =   11
         Top             =   405
         Width           =   495
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Frames"
         Height          =   255
         Left            =   1330
         TabIndex        =   10
         Top             =   645
         Width           =   735
      End
      Begin VB.Label lblFrames 
         Caption         =   "0"
         Height          =   255
         Left            =   5760
         TabIndex        =   9
         Top             =   690
         Width           =   495
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Speed"
         Height          =   255
         Left            =   1330
         TabIndex        =   8
         Top             =   930
         Width           =   735
      End
      Begin VB.Label lblSpeed 
         Caption         =   "0"
         Height          =   255
         Left            =   5760
         TabIndex        =   7
         Top             =   930
         Width           =   495
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Animation Name:"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmAnimationEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private CurrFrame As Byte
Private LastUpdate As Long
Private Size As Long

Private Sub opt32_Click()
    Size = 1
    picSprite.Width = 32
    picSprite.Height = 32
    frmAnimationEditor.scrlSprite.Max = (DDSD_Animation.lHeight \ PIC_Y) - 1
End Sub

Private Sub opt64_Click()
    Size = 2
    picSprite.Width = 64
    picSprite.Height = 64
    frmAnimationEditor.scrlSprite.Max = (DDSD_Animation2.lHeight \ PIC_Y) - 1
End Sub

Private Sub cmdOk_Click()
    AnimationEditorOk
End Sub

Private Sub cmdCancel_Click()
    AnimationEditorCancel
End Sub

Private Sub scrlSprite_Change()
    lblSprite.Caption = scrlSprite.Value
    CurrFrame = 0
End Sub

Private Sub scrlframes_Change()
    lblFrames.Caption = scrlFrames.Value
    CurrFrame = 0
End Sub

Private Sub scrlSpeed_Change()
    lblSpeed.Caption = scrlSpeed.Value
    CurrFrame = 0
End Sub

Private Sub tmrAnimation_Timer()
Dim sRECT As RECT
Dim dRECT As RECT
Dim Anim As Byte, Frames As Byte, Speed As Byte

    Anim = scrlSprite.Value
    Frames = scrlFrames.Value
    Speed = scrlSpeed.Value
    
    If Size = 0 Then Exit Sub

    If Size = 1 Then
         With dRECT
            .Top = 0
            .Bottom = PIC_Y
            .Left = 0
            .Right = PIC_X
        End With
    
        With sRECT
            .Top = Anim * PIC_Y
            .Bottom = .Top + PIC_Y
            .Left = CurrFrame * PIC_X
            .Right = .Left + PIC_X
        End With
        
        DD_AnimationSurf.BltToDC picSprite.hdc, sRECT, dRECT
    ElseIf Size = 2 Then
        With dRECT
            .Top = 0
            .Bottom = (PIC_Y * 2)
            .Left = 0
            .Right = (PIC_X * 2)
        End With
        
        With sRECT
            .Top = Anim * (PIC_Y * 2)
            .Bottom = .Top + (PIC_Y * 2)
            .Left = CurrFrame * (PIC_X * 2)
            .Right = .Left + (PIC_X * 2)
        End With
        
        DD_AnimationSurf2.BltToDC picSprite.hdc, sRECT, dRECT
    End If
    
    picSprite.Refresh

    If GetTickCount > LastUpdate Then
    
        CurrFrame = CurrFrame + 1
        
        If CurrFrame > Frames Then
            CurrFrame = 0
        Else
            LastUpdate = GetTickCount + Speed
        End If
    End If
End Sub
