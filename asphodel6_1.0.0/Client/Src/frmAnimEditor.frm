VERSION 5.00
Begin VB.Form frmAnimEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editing Animation"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   4335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   15
      Top             =   3960
      Width           =   2055
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Timer tmrAnim 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3720
      Top             =   3360
   End
   Begin VB.Frame FraPicture 
      Caption         =   "Picture"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   4095
      Begin VB.PictureBox picPic 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   240
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   13
         Top             =   1320
         Width           =   480
      End
      Begin VB.HScrollBar scrlPic 
         Height          =   255
         Left            =   960
         Max             =   1
         Min             =   1
         TabIndex        =   10
         Top             =   720
         Value           =   1
         Width           =   2295
      End
      Begin VB.HScrollBar scrlDelay 
         Height          =   255
         Left            =   960
         Max             =   1000
         Min             =   1
         TabIndex        =   9
         Top             =   960
         Value           =   100
         Width           =   2295
      End
      Begin VB.TextBox txtSizeY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         TabIndex        =   6
         Text            =   "32"
         Top             =   330
         Width           =   1095
      End
      Begin VB.TextBox txtSizeX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   4
         Text            =   "32"
         Top             =   330
         Width           =   1095
      End
      Begin VB.Label lblDelay 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   12
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblPic 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   11
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblPicture 
         Caption         =   "Picture:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblDelay2 
         Caption         =   "Delay:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Size Y:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblSizeX 
         Caption         =   "Size X:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   840
      MaxLength       =   10
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmAnimEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    AnimEditorCancel
End Sub

Private Sub cmdOk_Click()

    If LenB(Trim$(txtName.Text)) < 1 Then
        MsgBox "You need to add a name for this animation!", , "Error"
        Exit Sub
    End If
    
    AnimEditorOk
    
End Sub

Private Sub Form_Load()
    tmrAnim.Enabled = True
    scrlPic.Max = TOTAL_ANIMGFX
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Editor = 0
End Sub

Private Sub scrlDelay_Change()
    tmrAnim.Interval = scrlDelay.Value
    lblDelay.Caption = scrlDelay.Value
End Sub

Private Sub scrlDelay_Scroll()
    scrlDelay_Change
End Sub

Private Sub scrlPic_Change()
    txtSizeY.Text = DDSD_Anim(scrlPic.Value).lHeight
    lblPic.Caption = scrlPic.Value
End Sub

Private Sub scrlPic_Scroll()
    scrlPic_Change
End Sub

Private Sub tmrAnim_Timer()
    AnimEditorAnim = AnimEditorAnim + 1
    If AnimEditorAnim > (DDSD_Anim(scrlPic.Value).lWidth \ Val(txtSizeX.Text)) - 1 Then AnimEditorAnim = 0
    AnimEditorDrawPic
End Sub

Private Sub txtSizeX_Change()
    If LenB(Trim$(txtSizeX.Text)) < 1 Then txtSizeX.Text = 1
End Sub

Private Sub txtSizeX_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then Exit Sub
    If Not IsNumeric(Chr$(KeyAscii)) Then KeyAscii = 0
End Sub

Private Sub txtSizeY_Change()
    If LenB(Trim$(txtSizeY.Text)) < 1 Then txtSizeY.Text = 1
End Sub

Private Sub txtSizeY_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then Exit Sub
    If Not IsNumeric(Chr$(KeyAscii)) Then KeyAscii = 0
End Sub
