VERSION 5.00
Begin VB.Form frmMapEditor 
   Caption         =   "Map Editor"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMapEditor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillColor       =   &H00404000&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00404000&
      Height          =   3645
      Left            =   0
      ScaleHeight     =   241
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   479
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      Begin VB.Frame fraLayers 
         BackColor       =   &H80000005&
         Caption         =   "Layers"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   3840
         TabIndex        =   18
         Top             =   120
         Width           =   1575
         Begin VB.CommandButton cmdClear 
            BackColor       =   &H80000007&
            Caption         =   "Clear"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   29
            Top             =   2520
            Width           =   1215
         End
         Begin VB.OptionButton optFringe 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Fringe"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   1080
            Width           =   1215
         End
         Begin VB.OptionButton optAnim 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Animation"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   840
            Width           =   1215
         End
         Begin VB.OptionButton optMask 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Mask"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   600
            Width           =   1215
         End
         Begin VB.OptionButton optGround 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Ground"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   360
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Fill"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   2880
            Width           =   1215
         End
         Begin VB.OptionButton optMask2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Mask2"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            TabIndex        =   23
            Top             =   1320
            Width           =   975
         End
         Begin VB.OptionButton optM2Anim 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Animation 2"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            TabIndex        =   22
            Top             =   1560
            Width           =   1335
         End
         Begin VB.OptionButton optFAnim 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Animation 3"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            TabIndex        =   21
            Top             =   1800
            Width           =   1335
         End
         Begin VB.OptionButton optFringe2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Fringe2"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            TabIndex        =   20
            Top             =   2040
            Width           =   1335
         End
         Begin VB.OptionButton optF2Anim 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Fringe 2 Anim."
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            TabIndex        =   19
            Top             =   2280
            Width           =   1335
         End
      End
      Begin VB.PictureBox picSelect 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   5520
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   17
         Top             =   1080
         Width           =   480
      End
      Begin VB.PictureBox picBack 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   3360
         Left            =   120
         ScaleHeight     =   224
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   224
         TabIndex        =   15
         Top             =   120
         Width           =   3360
         Begin VB.PictureBox picBackSelect 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   960
            Left            =   0
            ScaleHeight     =   64
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   64
            TabIndex        =   16
            Top             =   0
            Width           =   960
         End
      End
      Begin VB.VScrollBar scrlPicture 
         Height          =   3375
         Left            =   3480
         Max             =   255
         TabIndex        =   14
         Top             =   120
         Width           =   255
      End
      Begin VB.CommandButton cmdProperties 
         BackColor       =   &H00404000&
         Caption         =   "Properties"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         MaskColor       =   &H00404000&
         TabIndex        =   13
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   12
         Top             =   2520
         Width           =   1575
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   11
         Top             =   3000
         Width           =   1575
      End
      Begin VB.OptionButton optAttribs 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Attributes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5520
         TabIndex        =   10
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton optLayers 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Layers"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5520
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.Frame fraAttribs 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Caption         =   "Attributes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   3840
         TabIndex        =   1
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
         Begin VB.OptionButton optKey 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Key"
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   120
            TabIndex        =   8
            Top             =   1200
            Width           =   1215
         End
         Begin VB.OptionButton optNpcAvoid 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Npc Avoid"
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   120
            TabIndex        =   7
            Top             =   960
            Width           =   1215
         End
         Begin VB.OptionButton optItem 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Item"
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   120
            TabIndex        =   6
            Top             =   720
            Width           =   1215
         End
         Begin VB.CommandButton cmdClear2 
            Caption         =   "Clear"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   5
            Top             =   1680
            Width           =   1215
         End
         Begin VB.OptionButton optWarp 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Warp"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton optBlocked 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Blocked"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optKeyOpen 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Key Open"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   120
            TabIndex        =   2
            Top             =   1440
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "frmMapEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' // MAP EDITOR STUFF //

Private Sub optLayers_Click()
    If optLayers.Value = True Then
        fraLayers.Visible = True
        fraAttribs.Visible = False
    End If
End Sub

Private Sub optAttribs_Click()
    If optAttribs.Value = True Then
        fraLayers.Visible = False
        fraAttribs.Visible = True
    End If
End Sub

Private Sub picBackSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call EditorChooseTile(Button, Shift, X, Y)
End Sub

Private Sub picBackSelect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call EditorChooseTile(Button, Shift, X, Y)
End Sub

Private Sub cmdSend_Click()
    Call EditorSend

End Sub

Private Sub cmdCancel_Click()
    Call EditorCancel

End Sub

Private Sub cmdProperties_Click()
    frmMapProperties.Show vbModal
End Sub

Private Sub optWarp_Click()
    frmMapWarp.Show vbModal
End Sub

Private Sub optItem_Click()
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

Private Sub cmdClear_Click()
    Call EditorClearLayer
End Sub

Private Sub cmdClear2_Click()
    Call EditorClearAttribs
End Sub




