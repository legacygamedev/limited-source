VERSION 5.00
Begin VB.Form frmMapEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Editor"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   7260
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMapEditor 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5115
      Left            =   0
      ScaleHeight     =   341
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   484
      TabIndex        =   0
      Top             =   0
      Width           =   7260
      Begin VB.Frame fraLayers 
         Caption         =   "Layers"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3090
         Left            =   3960
         TabIndex        =   35
         Top             =   120
         Width           =   3120
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   2040
            TabIndex        =   45
            Top             =   2640
            Width           =   855
         End
         Begin VB.OptionButton optFringe 
            Caption         =   "Fringe"
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
            TabIndex        =   44
            Top             =   1800
            Width           =   855
         End
         Begin VB.OptionButton optAnim 
            Caption         =   "Animation"
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
            TabIndex        =   43
            Top             =   840
            Width           =   1215
         End
         Begin VB.OptionButton optMask 
            Caption         =   "Mask"
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
            TabIndex        =   42
            Top             =   600
            Width           =   1215
         End
         Begin VB.OptionButton optGround 
            Caption         =   "Ground"
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
            TabIndex        =   41
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optMask2 
            Caption         =   "Mask 2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   40
            Top             =   1200
            Width           =   1005
         End
         Begin VB.OptionButton optM2Anim 
            Caption         =   "Animation"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   39
            Top             =   1440
            Width           =   1245
         End
         Begin VB.OptionButton optFAnim 
            Caption         =   "Animation"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   38
            Top             =   2040
            Width           =   1095
         End
         Begin VB.OptionButton optFringe2 
            Caption         =   "Fringe 2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   37
            Top             =   2400
            Width           =   1080
         End
         Begin VB.OptionButton optF2Anim 
            Caption         =   "Animation"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   36
            Top             =   2640
            Width           =   1080
         End
      End
      Begin VB.PictureBox picSelect 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   5040
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   34
         Top             =   3360
         Visible         =   0   'False
         Width           =   480
         Begin VB.PictureBox MouseSelected 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   46
            Top             =   120
            Width           =   255
         End
      End
      Begin VB.PictureBox picBack 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   4320
         Left            =   105
         ScaleHeight     =   288
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   224
         TabIndex        =   32
         Top             =   120
         Width           =   3360
         Begin VB.PictureBox picBackSelect 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   2520
            Left            =   0
            ScaleHeight     =   168
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   224
            TabIndex        =   33
            Top             =   0
            Width           =   3360
            Begin VB.Shape shpSelected 
               BorderColor     =   &H000000FF&
               Height          =   480
               Left            =   0
               Top             =   0
               Width           =   480
            End
         End
      End
      Begin VB.VScrollBar scrlPicture 
         Height          =   4305
         LargeChange     =   10
         Left            =   3465
         Max             =   512
         TabIndex        =   31
         Top             =   135
         Width           =   255
      End
      Begin VB.OptionButton optAttribs 
         Caption         =   "Attributes"
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
         Left            =   3960
         TabIndex        =   30
         Top             =   3600
         Width           =   1035
      End
      Begin VB.OptionButton optLayers 
         Caption         =   "Layers"
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
         Left            =   3960
         TabIndex        =   29
         Top             =   3360
         Value           =   -1  'True
         Width           =   780
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   90
         Max             =   7
         TabIndex        =   28
         Top             =   4440
         Width           =   3375
      End
      Begin VB.CheckBox optMapGrid 
         Caption         =   "Map Grid"
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
         Left            =   5640
         TabIndex        =   27
         Top             =   3360
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox optSM 
         Caption         =   "ScreenShot Mode"
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
         Left            =   5640
         TabIndex        =   26
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Frame fraAttribs 
         Caption         =   "Attributes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3090
         Left            =   3960
         TabIndex        =   5
         Top             =   120
         Visible         =   0   'False
         Width           =   3120
         Begin VB.OptionButton optBank 
            Caption         =   "Bank"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   50
            Top             =   2760
            Width           =   690
         End
         Begin VB.OptionButton optLevelBlock 
            Caption         =   "Level Block"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   48
            Top             =   2520
            Width           =   1170
         End
         Begin VB.OptionButton optSexBlock 
            Caption         =   "Sex Block"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1800
            TabIndex        =   47
            Top             =   2520
            Width           =   1170
         End
         Begin VB.OptionButton optKey 
            Caption         =   "Key"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   25
            Top             =   1320
            Width           =   1215
         End
         Begin VB.OptionButton optNpcAvoid 
            Caption         =   "Npc Avoid"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   24
            Top             =   1080
            Width           =   1215
         End
         Begin VB.OptionButton optItem 
            Caption         =   "Item"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   23
            Top             =   840
            Width           =   1215
         End
         Begin VB.CommandButton cmdClear2 
            Caption         =   "Clear"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1800
            TabIndex        =   22
            Top             =   2760
            Width           =   975
         End
         Begin VB.OptionButton optWarp 
            Caption         =   "Warp"
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
            Left            =   120
            TabIndex        =   21
            Top             =   600
            Width           =   1215
         End
         Begin VB.OptionButton optBlocked 
            Caption         =   "Blocked"
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
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optKeyOpen 
            Caption         =   "Key Open"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   19
            Top             =   1560
            Width           =   1215
         End
         Begin VB.OptionButton optHeal 
            Caption         =   "Heal"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   18
            Top             =   1800
            Width           =   1035
         End
         Begin VB.OptionButton optKill 
            Caption         =   "Kill"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   17
            Top             =   2040
            Width           =   810
         End
         Begin VB.OptionButton optShop 
            Caption         =   "Shop"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1800
            TabIndex        =   16
            Top             =   1800
            Width           =   810
         End
         Begin VB.OptionButton optCBlock 
            Caption         =   "Class Block"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1800
            TabIndex        =   15
            Top             =   2040
            Width           =   1170
         End
         Begin VB.OptionButton optArena 
            Caption         =   "Arena"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1800
            TabIndex        =   14
            Top             =   2280
            Width           =   1170
         End
         Begin VB.OptionButton optSound 
            Caption         =   "Play Sound"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   13
            Top             =   2280
            Width           =   1170
         End
         Begin VB.OptionButton optSprite 
            Caption         =   "Sprite Change"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1800
            TabIndex        =   12
            Top             =   1560
            Width           =   1200
         End
         Begin VB.OptionButton optSign 
            Caption         =   "Sign"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1800
            TabIndex        =   11
            Top             =   1320
            Width           =   1080
         End
         Begin VB.OptionButton optDoor 
            Caption         =   "Door"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1800
            TabIndex        =   10
            Top             =   1080
            Width           =   960
         End
         Begin VB.OptionButton optNotice 
            Caption         =   "Notice"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1800
            TabIndex        =   9
            Top             =   840
            Width           =   1155
         End
         Begin VB.OptionButton optChest 
            Caption         =   "Chest"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   3840
            TabIndex        =   8
            Top             =   3240
            Width           =   720
         End
         Begin VB.OptionButton optClassChange 
            Caption         =   "Class Change"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1800
            TabIndex        =   7
            Top             =   600
            Width           =   1200
         End
         Begin VB.OptionButton optScripted 
            Caption         =   "Scripted"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1800
            TabIndex        =   6
            Top             =   360
            Width           =   1050
         End
      End
      Begin VB.CommandButton cmdFill 
         Caption         =   "Fill"
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
         Left            =   3960
         TabIndex        =   4
         Top             =   4440
         Width           =   615
      End
      Begin VB.CommandButton cmdProperties 
         Caption         =   "Properties"
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
         Left            =   4560
         TabIndex        =   3
         Top             =   4440
         Width           =   1095
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send"
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
         Left            =   5640
         TabIndex        =   2
         Top             =   4440
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
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
         Left            =   6360
         TabIndex        =   1
         Top             =   4440
         Width           =   720
      End
      Begin VB.Label lblXY 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cursor - X:0 - Y:0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3960
         TabIndex        =   49
         Top             =   4800
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmMapEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' // MAP EDITOR STUFF //
Dim KeyShift As Boolean

Private Sub optLayers_Click()
    If optLayers.Value = True Then
        fraLayers.Visible = True
        fraAttribs.Visible = False
        'cmdFill.Caption = "Fill Map With Tile"
    End If
End Sub

Private Sub optAttribs_Click()
    If optAttribs.Value = True Then
        fraLayers.Visible = False
        fraAttribs.Visible = True
        'cmdFill.Caption = "Fill Map With Attribute"
    End If
End Sub

Private Sub optArena_Click()
    frmArena.Show vbModal
End Sub

Private Sub optCBlock_Click()
    frmBClass.scrlNum1.Max = Max_Classes
    frmBClass.scrlNum2.Max = Max_Classes
    frmBClass.scrlNum3.Max = Max_Classes
    frmBClass.Show vbModal
End Sub

Private Sub optClassChange_Click()
    frmClassChange.scrlClass.Max = Max_Classes
    frmClassChange.scrlReqClass.Max = Max_Classes
    frmClassChange.Show vbModal
End Sub

Private Sub optLevelBlock_Click()
    frmLevelBlock.Show vbModal
End Sub

Private Sub optSexBlock_Click()
    frmSexBlock.Show vbModal
End Sub

Private Sub picBackSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyShift Then
        KeyShift = True
    End If
End Sub

Private Sub picBackSelect_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyShift = False
End Sub

Private Sub picBackSelect_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If KeyShift = False Then
            Call EditorChooseTile(Button, Shift, x, y)
            shpSelected.Width = 32
            shpSelected.Height = 32
        Else
            EditorTileX = Int(x / PIC_X)
            EditorTileY = Int(y / PIC_Y)
            
            If Int(EditorTileX * PIC_X) >= shpSelected.Left + shpSelected.Width Then
                EditorTileX = Int(EditorTileX * PIC_X + PIC_X) - (shpSelected.Left + shpSelected.Width)
                shpSelected.Width = shpSelected.Width + Int(EditorTileX)
            Else
                If shpSelected.Width > PIC_X Then
                    If Int(EditorTileX * PIC_X) >= shpSelected.Left Then
                        EditorTileX = (EditorTileX * PIC_X + PIC_X) - (shpSelected.Left + shpSelected.Width)
                        shpSelected.Width = shpSelected.Width + Int(EditorTileX)
                    End If
                End If
            End If
            
            If Int(EditorTileY * PIC_Y) >= shpSelected.Top + shpSelected.Height Then
                EditorTileY = Int(EditorTileY * PIC_Y + PIC_Y) - (shpSelected.Top + shpSelected.Height)
                shpSelected.Height = shpSelected.Height + Int(EditorTileY)
            Else
                If shpSelected.Height > PIC_Y Then
                    If Int(EditorTileY * PIC_Y) >= shpSelected.Top Then
                        EditorTileY = (EditorTileY * PIC_Y + PIC_Y) - (shpSelected.Top + shpSelected.Height)
                        shpSelected.Height = shpSelected.Height + Int(EditorTileY)
                    End If
                End If
            End If
        End If
    End If
        
    EditorTileX = Int((shpSelected.Left + PIC_X) / PIC_X)
    EditorTileY = Int((shpSelected.Top + PIC_Y) / PIC_Y)
End Sub

Private Sub picBackSelect_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If KeyShift = False Then
            Call EditorChooseTile(Button, Shift, x, y)
            shpSelected.Width = 32
            shpSelected.Height = 32
        Else
            EditorTileX = Int(x / PIC_X)
            EditorTileY = Int(y / PIC_Y)
            
            If Int(EditorTileX * PIC_X) >= shpSelected.Left + shpSelected.Width Then
                EditorTileX = Int(EditorTileX * PIC_X + PIC_X) - (shpSelected.Left + shpSelected.Width)
                shpSelected.Width = shpSelected.Width + Int(EditorTileX)
            Else
                If shpSelected.Width > PIC_X Then
                    If Int(EditorTileX * PIC_X) >= shpSelected.Left Then
                        EditorTileX = (EditorTileX * PIC_X + PIC_X) - (shpSelected.Left + shpSelected.Width)
                        shpSelected.Width = shpSelected.Width + Int(EditorTileX)
                    End If
                End If
            End If
            
            If Int(EditorTileY * PIC_Y) >= shpSelected.Top + shpSelected.Height Then
                EditorTileY = Int(EditorTileY * PIC_Y + PIC_Y) - (shpSelected.Top + shpSelected.Height)
                shpSelected.Height = shpSelected.Height + Int(EditorTileY)
            Else
                If shpSelected.Height > PIC_Y Then
                    If Int(EditorTileY * PIC_Y) >= shpSelected.Top Then
                        EditorTileY = (EditorTileY * PIC_Y + PIC_Y) - (shpSelected.Top + shpSelected.Height)
                        shpSelected.Height = shpSelected.Height + Int(EditorTileY)
                    End If
                End If
            End If
        End If
    End If
        
    EditorTileX = Int(shpSelected.Left / PIC_X)
    EditorTileY = Int(shpSelected.Top / PIC_Y)
End Sub

Private Sub cmdSend_Click()
Dim x As Long

    x = MsgBox("Are you sure you want to make these changes?", vbYesNo)
    If x = vbNo Then
        Exit Sub
    End If
    
    ScreenMode = 0
    Call EditorSend
End Sub

Private Sub cmdCancel_Click()
Dim x As Long

    x = MsgBox("Are you sure you want to discard your changes?", vbYesNo)
    If x = vbNo Then
        Exit Sub
    End If
    
    ScreenMode = 0
    CheckMap(GetPlayerMap(MyIndex)) = SaveMap
    Call EditorCancel
End Sub

Private Sub cmdProperties_Click()
    frmMapProperties.Show vbModal
End Sub

Private Sub optWarp_Click()
    frmMapWarp.Show vbModal
End Sub

Private Sub optItem_Click()
    frmMapItem.scrlItem.Value = 1
    frmMapItem.Show vbModal
End Sub

Private Sub optKey_Click()
    frmMapKey.Show vbModal
End Sub

Private Sub optKeyOpen_Click()
    frmKeyOpen.Show vbModal
End Sub

Private Sub scrlPicture_Change()
    Call EditorTileScroll
End Sub

Private Sub optMapGrid_Click()
    WriteINI "CONFIG", "MapGrid", optMapGrid.Value, App.Path & "\config.ini"
End Sub

Private Sub optNotice_Click()
    frmNotice.Show vbModal
End Sub

Private Sub optScripted_Click()
    frmScript.Show vbModal
End Sub

Private Sub optShop_Click()
    frmShop.scrlNum.Max = MAX_SHOPS
    frmShop.Show vbModal
End Sub

Private Sub optSign_Click()
    frmSign.Show vbModal
End Sub

Private Sub optSM_Click()
If optSM.Value = 0 Then
    ScreenMode = 0
Else
    ScreenMode = 1
End If
End Sub

Private Sub optSound_Click()
    frmSound.Show vbModal
End Sub

Private Sub optSprite_Click()
    frmSpriteChange.scrlItem.Max = MAX_ITEMS
    frmSpriteChange.Show vbModal
End Sub

Private Sub HScroll1_Change()
    picBackSelect.Left = (HScroll1.Value * PIC_X) * -1
End Sub

Private Sub cmdFill_Click()
Dim y As Long
Dim x As Long

x = MsgBox("Are you sure you want to fill the map?", vbYesNo)
If x = vbNo Then
    Exit Sub
End If

If optAttribs.Value = False Then
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            With CheckMap(GetPlayerMap(MyIndex)).Tile(x, y)
                If optGround.Value = True Then .Ground = EditorTileY * 14 + EditorTileX
                If optMask.Value = True Then .Mask = EditorTileY * 14 + EditorTileX
                If optAnim.Value = True Then .Anim = EditorTileY * 14 + EditorTileX
                If optMask2.Value = True Then .Mask2 = EditorTileY * 14 + EditorTileX
                If optM2Anim.Value = True Then .M2Anim = EditorTileY * 14 + EditorTileX
                If optFringe.Value = True Then .Fringe = EditorTileY * 14 + EditorTileX
                If optFAnim.Value = True Then .FAnim = EditorTileY * 14 + EditorTileX
                If optFringe2.Value = True Then .Fringe2 = EditorTileY * 14 + EditorTileX
                If optF2Anim.Value = True Then .F2Anim = EditorTileY * 14 + EditorTileX
            End With
        Next x
    Next y
Else
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            With CheckMap(GetPlayerMap(MyIndex)).Tile(x, y)
                If optBlocked.Value = True Then .Type = TILE_TYPE_BLOCKED
                If optWarp.Value = True Then
                    .Type = TILE_TYPE_WARP
                    .Data1 = EditorWarpMap
                    .Data2 = EditorWarpX
                    .Data3 = EditorWarpY
                    .String1 = ""
                    .String2 = ""
                    .String3 = ""
                End If

                If optHeal.Value = True Then
                    .Type = TILE_TYPE_HEAL
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                    .String1 = ""
                    .String2 = ""
                    .String3 = ""
                End If

                If optKill.Value = True Then
                    .Type = TILE_TYPE_KILL
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                    .String1 = ""
                    .String2 = ""
                    .String3 = ""
                End If

                If optItem.Value = True Then
                    .Type = TILE_TYPE_ITEM
                    .Data1 = ItemEditorNum
                    .Data2 = ItemEditorValue
                    .Data3 = 0
                    .String1 = ""
                    .String2 = ""
                    .String3 = ""
                End If
                If optNpcAvoid.Value = True Then
                    .Type = TILE_TYPE_NPCAVOID
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                    .String1 = ""
                    .String2 = ""
                    .String3 = ""
                End If
                If optKey.Value = True Then
                    .Type = TILE_TYPE_KEY
                    .Data1 = KeyEditorNum
                    .Data2 = KeyEditorTake
                    .Data3 = 0
                    .String1 = ""
                    .String2 = ""
                    .String3 = ""
                End If
                If optKeyOpen.Value = True Then
                    .Type = TILE_TYPE_KEYOPEN
                    .Data1 = KeyOpenEditorX
                    .Data2 = KeyOpenEditorY
                    .Data3 = 0
                    .String1 = KeyOpenEditorMsg
                    .String2 = ""
                    .String3 = ""
                End If
                If optShop.Value = True Then
                    .Type = TILE_TYPE_SHOP
                    .Data1 = EditorShopNum
                    .Data2 = 0
                    .Data3 = 0
                    .String1 = ""
                    .String2 = ""
                    .String3 = ""
                End If
                If optCBlock.Value = True Then
                    .Type = TILE_TYPE_CBLOCK
                    .Data1 = EditorItemNum1
                    .Data2 = EditorItemNum2
                    .Data3 = EditorItemNum3
                    .String1 = ""
                    .String2 = ""
                    .String3 = ""
                End If
                If optArena.Value = True Then
                    .Type = TILE_TYPE_ARENA
                    .Data1 = Arena1
                    .Data2 = Arena2
                    .Data3 = Arena3
                    .String1 = ""
                    .String2 = ""
                    .String3 = ""
                End If
                If optSound.Value = True Then
                    .Type = TILE_TYPE_SOUND
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                    .String1 = SoundFileName
                    .String2 = ""
                    .String3 = ""
                End If
                If optSprite.Value = True Then
                    .Type = TILE_TYPE_SPRITE_CHANGE
                    .Data1 = SpritePic
                    .Data2 = SpriteItem
                    .Data3 = SpritePrice
                    .String1 = ""
                    .String2 = ""
                    .String3 = ""
                End If
                If optSign.Value = True Then
                    .Type = TILE_TYPE_SIGN
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                    .String1 = SignLine1
                    .String2 = SignLine2
                    .String3 = SignLine3
                End If
                If optDoor.Value = True Then
                    .Type = TILE_TYPE_DOOR
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                    .String1 = ""
                    .String2 = ""
                    .String3 = ""
                End If
                If optNotice.Value = True Then
                    .Type = TILE_TYPE_NOTICE
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                    .String1 = NoticeTitle
                    .String2 = NoticeText
                    .String3 = NoticeSound
                End If
                If optChest.Value = True Then
                    .Type = TILE_TYPE_CHEST
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                    .String1 = ""
                    .String2 = ""
                    .String3 = ""
                End If
                If optClassChange.Value = True Then
                    .Type = TILE_TYPE_CLASS_CHANGE
                    .Data1 = ClassChange
                    .Data2 = ClassChangeReq
                    .Data3 = 0
                    .String1 = ""
                    .String2 = ""
                    .String3 = ""
                End If
                If optScripted.Value = True Then
                    .Type = TILE_TYPE_SCRIPTED
                    .Data1 = ScriptNum
                    .Data2 = 0
                    .Data3 = 0
                    .String1 = ""
                    .String2 = ""
                    .String3 = ""
                End If
                If frmMapEditor.optSexBlock.Value = True Then
                    .Type = TILE_TYPE_SEX_BLOCK
                    .Data1 = SexBlockNum
                    .Data2 = 0
                    .Data3 = 0
                    .String1 = ""
                    .String2 = ""
                    .String3 = ""
                End If
                If frmMapEditor.optLevelBlock.Value = True Then
                    .Type = TILE_TYPE_LEVEL_BLOCK
                    .Data1 = LevelBlockLow
                    .Data2 = LevelBlockHigh
                    .Data3 = 0
                    .String1 = ""
                    .String2 = ""
                    .String3 = ""
                End If
                If frmMapEditor.optBank.Value = True Then
                    .Type = TILE_TYPE_BANK
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                    .String1 = ""
                    .String2 = ""
                    .String3 = ""
                End If
            End With
        Next x
    Next y
End If
End Sub

Private Sub cmdClear_Click()
    Call EditorClearLayer
End Sub

Private Sub cmdClear2_Click()
    Call EditorClearAttribs
End Sub
