VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMirage 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14115
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMirage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   528
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   941
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picSpellsList 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2685
      Left            =   2640
      ScaleHeight     =   177
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   167
      TabIndex        =   42
      Top             =   120
      Visible         =   0   'False
      Width           =   2535
      Begin VB.ListBox lstSpells 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1920
         ItemData        =   "frmMirage.frx":08CA
         Left            =   120
         List            =   "frmMirage.frx":08CC
         TabIndex        =   43
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label lblSpellSelected 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   58
         Top             =   2400
         UseMnemonic     =   0   'False
         Width           =   975
      End
      Begin VB.Label lblSpells 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Spells"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   960
         TabIndex        =   46
         Top             =   0
         UseMnemonic     =   0   'False
         Width           =   735
      End
      Begin VB.Label lblCast 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cast"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   840
         TabIndex        =   45
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label lblSpellsCancel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Hide"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   840
         TabIndex        =   44
         Top             =   2400
         Width           =   855
      End
   End
   Begin VB.PictureBox picInvList 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2925
      Left            =   120
      ScaleHeight     =   193
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   167
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   2535
      Begin VB.PictureBox picInvSelected 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   1920
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   59
         Top             =   2280
         Width           =   480
      End
      Begin VB.ListBox lstInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1920
         ItemData        =   "frmMirage.frx":08CE
         Left            =   120
         List            =   "frmMirage.frx":08D0
         TabIndex        =   38
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label lblInvSelected 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   57
         Top             =   2640
         UseMnemonic     =   0   'False
         Width           =   855
      End
      Begin VB.Label lblCancel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Hide"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   840
         TabIndex        =   41
         Top             =   2640
         UseMnemonic     =   0   'False
         Width           =   735
      End
      Begin VB.Label lblDropItem 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Drop Item"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   720
         TabIndex        =   40
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label lblUseItem 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Use Item"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   720
         TabIndex        =   39
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label lblInventory 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Inventory"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   37
         Top             =   0
         UseMnemonic     =   0   'False
         Width           =   1215
      End
   End
   Begin VB.PictureBox picMapEditor 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7485
      Left            =   9960
      ScaleHeight     =   497
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   263
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   3975
      Begin VB.Frame frmTileSet 
         Caption         =   "TileSet"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   52
         Top             =   3600
         Width           =   3495
         Begin VB.HScrollBar scrlTileSet 
            Height          =   255
            Left            =   120
            Max             =   10
            Min             =   1
            TabIndex        =   53
            Top             =   240
            Value           =   1
            Width           =   2775
         End
         Begin VB.Label lblTileset 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   3000
            TabIndex        =   54
            Top             =   240
            Width           =   375
         End
      End
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
         Height          =   2415
         Left            =   2160
         TabIndex        =   31
         Top             =   4320
         Width           =   1455
         Begin VB.OptionButton optMask2 
            Caption         =   "Mask2"
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
            TabIndex        =   61
            Top             =   960
            Width           =   1215
         End
         Begin VB.OptionButton optFringe2 
            Caption         =   "Fringe2"
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
            TabIndex        =   60
            Top             =   1440
            Width           =   1215
         End
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
            Left            =   120
            TabIndex        =   32
            Top             =   2040
            Width           =   1215
         End
         Begin VB.CommandButton cmdFill 
            Caption         =   "Fill"
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
            Left            =   120
            TabIndex        =   56
            Top             =   1800
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
            TabIndex        =   36
            Top             =   240
            Value           =   -1  'True
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
            TabIndex        =   35
            Top             =   480
            Width           =   1215
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
            TabIndex        =   34
            Top             =   720
            Width           =   1215
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
            TabIndex        =   33
            Top             =   1200
            Width           =   1215
         End
      End
      Begin VB.Frame fraAttribs 
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
         Height          =   2175
         Left            =   2160
         TabIndex        =   24
         Top             =   4320
         Visible         =   0   'False
         Width           =   1455
         Begin VB.OptionButton optKeyOpen 
            Caption         =   "Key Open"
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
            TabIndex        =   47
            Top             =   1440
            Width           =   1215
         End
         Begin VB.OptionButton optBlocked 
            Caption         =   "Blocked"
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
            TabIndex        =   30
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optWarp 
            Caption         =   "Warp"
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
            TabIndex        =   29
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton cmdClear2 
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
            Left            =   120
            TabIndex        =   28
            Top             =   1800
            Width           =   1215
         End
         Begin VB.OptionButton optItem 
            Caption         =   "Item"
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
            Left            =   120
            TabIndex        =   27
            Top             =   720
            Width           =   1215
         End
         Begin VB.OptionButton optNpcAvoid 
            Caption         =   "Npc Avoid"
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
            Left            =   120
            TabIndex        =   26
            Top             =   960
            Width           =   1215
         End
         Begin VB.OptionButton optKey 
            Caption         =   "Key"
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
            Left            =   120
            TabIndex        =   25
            Top             =   1200
            Width           =   1215
         End
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
         Left            =   360
         TabIndex        =   23
         Top             =   4920
         Value           =   -1  'True
         Width           =   1575
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
         Left            =   360
         TabIndex        =   22
         Top             =   5160
         Width           =   1575
      End
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
         Left            =   240
         TabIndex        =   21
         Top             =   6960
         Width           =   1215
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send"
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
         Left            =   2520
         TabIndex        =   20
         Top             =   6960
         Width           =   1215
      End
      Begin VB.CommandButton cmdProperties 
         Caption         =   "Properties"
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
         TabIndex        =   19
         Top             =   5640
         Width           =   1215
      End
      Begin VB.VScrollBar scrlPicture 
         Height          =   3375
         Left            =   3480
         Max             =   255
         TabIndex        =   18
         Top             =   120
         Width           =   255
      End
      Begin VB.PictureBox picBack 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3360
         Left            =   120
         ScaleHeight     =   224
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   224
         TabIndex        =   16
         Top             =   120
         Width           =   3360
         Begin VB.PictureBox picBackSelect 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   960
            Left            =   0
            ScaleHeight     =   64
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   64
            TabIndex        =   17
            Top             =   0
            Width           =   960
            Begin VB.Shape shpLoc 
               BorderColor     =   &H00FF0000&
               Height          =   480
               Left            =   0
               Top             =   0
               Width           =   480
            End
            Begin VB.Shape shpSelected 
               BorderColor     =   &H000000FF&
               Height          =   480
               Left            =   0
               Top             =   0
               Width           =   480
            End
         End
      End
      Begin VB.PictureBox picSelect 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   1440
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   15
         Top             =   4320
         Width           =   480
      End
      Begin VB.Label lblPreview 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tile Preview: "
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
         TabIndex        =   55
         Top             =   4440
         UseMnemonic     =   0   'False
         Width           =   1095
      End
   End
   Begin VB.PictureBox picGUI 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5775
      Left            =   7920
      Picture         =   "frmMirage.frx":08D2
      ScaleHeight     =   385
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   130
      TabIndex        =   3
      Top             =   105
      Width           =   1950
      Begin VB.PictureBox picInventoryButton 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   690
         Left            =   240
         Picture         =   "frmMirage.frx":2569C
         ScaleHeight     =   660
         ScaleWidth      =   660
         TabIndex        =   9
         ToolTipText     =   "Inventory"
         Top             =   3000
         Width           =   690
      End
      Begin VB.PictureBox picSpellsButton 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   690
         Left            =   1080
         Picture         =   "frmMirage.frx":26D8E
         ScaleHeight     =   660
         ScaleWidth      =   660
         TabIndex        =   8
         ToolTipText     =   "Spells"
         Top             =   3000
         Width           =   690
      End
      Begin VB.PictureBox picStatsButton 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   690
         Left            =   240
         Picture         =   "frmMirage.frx":28480
         ScaleHeight     =   660
         ScaleWidth      =   660
         TabIndex        =   7
         ToolTipText     =   "Stats"
         Top             =   3840
         Width           =   690
      End
      Begin VB.PictureBox picTrainButton 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   690
         Left            =   1080
         Picture         =   "frmMirage.frx":29B72
         ScaleHeight     =   660
         ScaleWidth      =   660
         TabIndex        =   6
         ToolTipText     =   "Train"
         Top             =   3840
         Width           =   690
      End
      Begin VB.PictureBox picTradeButton 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   690
         Left            =   240
         Picture         =   "frmMirage.frx":2B264
         ScaleHeight     =   660
         ScaleWidth      =   660
         TabIndex        =   5
         ToolTipText     =   "Trade"
         Top             =   4680
         Width           =   690
      End
      Begin VB.PictureBox picQuitButton 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   690
         Left            =   1080
         Picture         =   "frmMirage.frx":2C956
         ScaleHeight     =   660
         ScaleWidth      =   660
         TabIndex        =   4
         ToolTipText     =   "Quit"
         Top             =   4680
         Width           =   690
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "SP:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   50
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "MP:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   49
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "HP:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   48
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Mirage Source"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   0
         TabIndex        =   13
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblHP 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   600
         TabIndex        =   12
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblMP 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   600
         TabIndex        =   11
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblSP 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   255
         Left            =   600
         TabIndex        =   10
         Top             =   1680
         Width           =   1095
      End
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   6240
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   2778
      _Version        =   393217
      BackColor       =   4210752
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMirage.frx":2E048
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5760
      Left            =   120
      ScaleHeight     =   384
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   512
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   7680
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   9360
      Top             =   7200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtMyChat 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   120
      TabIndex        =   51
      Top             =   5985
      Width           =   7695
   End
   Begin VB.Line Line4 
      X1              =   7
      X2              =   7
      Y1              =   7
      Y2              =   393
   End
   Begin VB.Line Line3 
      X1              =   7
      X2              =   520
      Y1              =   392
      Y2              =   392
   End
   Begin VB.Line Line2 
      X1              =   7
      X2              =   520
      Y1              =   7
      Y2              =   7
   End
   Begin VB.Line Line1 
      X1              =   520
      X2              =   520
      Y1              =   7
      Y2              =   393
   End
End
Attribute VB_Name = "frmMirage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ******************************************
' **            Mirage Source 4           **
' ******************************************

' ************
' ** Events **
' ************

Private Sub Form_Load()
    frmMirage.Width = 10080
End Sub

Private Sub Form_Unload(Cancel As Integer)
    InGame = False
End Sub

' Winsock event
Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
    If IsConnected Then
        Call IncomingData(bytesTotal)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Call HandleKeypresses(KeyAscii)
    
    ' prevents textbox on error ding sound
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
    
        Case vbKeyEscape
            frmMirage.picInvList.Visible = False
            frmMirage.picSpellsList.Visible = False
    
        Case vbKeyF1
            Call PlayerSearch(GetPlayerX(MyIndex), GetPlayerY(MyIndex))
    
        Case vbKeyF3
            Call CastSpell
    
        Case vbKeyF4
            Call UseItem
    
    End Select

End Sub

Private Sub txtMyChat_Change()
    MyText = txtMyChat
End Sub

Private Sub txtChat_GotFocus()
    SetFocusOnChat
End Sub

Private Sub picInvList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SOffsetX = X
    SOffsetY = Y
End Sub

Private Sub picSpellsList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SOffsetX = X
    SOffsetY = Y
End Sub

Private Sub picInvList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MovePicture(frmMirage.picInvList, Button, Shift, X, Y)
End Sub

Private Sub picSpellsList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MovePicture(frmMirage.picSpellsList, Button, Shift, X, Y)
End Sub

' ***************
' ** Inventory **
' ***************

Private Sub lblUseItem_Click()
    Call UseItem
End Sub

Private Sub lblDropItem_Click()
Dim InvNum As Long

    InvNum = frmMirage.lstInv.ListIndex + 1

    If GetPlayerInvItemNum(MyIndex, InvNum) > 0 Then
        If GetPlayerInvItemNum(MyIndex, InvNum) <= MAX_ITEMS Then
            If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_CURRENCY Then
                ' Show them the drop dialog
                frmDrop.Show vbModal
            Else
                Call SendDropItem(frmMirage.lstInv.ListIndex + 1, 0)
                
                ' clear inventory graphic
                frmMirage.picInvSelected.Cls
            End If
        End If
    End If
End Sub

Private Sub lstInv_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InventoryItemSelected = frmMirage.lstInv.ListIndex + 1
    lblInvSelected.Caption = "<slot " & InventoryItemSelected & ">"
    Call BltInventory(GetPlayerInvItemNum(MyIndex, InventoryItemSelected))
End Sub

Private Sub lblCancel_Click()
    picInvList.Visible = False
End Sub

Private Sub lstInv_GotFocus()
On Error Resume Next
    frmMirage.picScreen.SetFocus
End Sub

' ************
' ** Spells **
' ************

Private Sub lblCast_Click()
    Call CastSpell
End Sub

Private Sub lstSpells_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SpellSelected = frmMirage.lstSpells.ListIndex + 1
    lblSpellSelected.Caption = "<slot " & SpellSelected & ">"
End Sub

Private Sub lblSpellsCancel_Click()
    picSpellsList.Visible = False
End Sub

Private Sub lstSpells_GotFocus()
On Error Resume Next
    frmMirage.picScreen.SetFocus
End Sub
' *****************
' ** GUI Buttons **
' *****************

Private Sub picInventoryButton_Click()
    Call UpdateInventory
    picInvList.Visible = (Not picInvList.Visible)
End Sub

Private Sub picSpellsButton_Click()
Dim Buffer As clsBuffer

    If picSpellsList.Visible Then
        picSpellsList.Visible = False
    Else
        Set Buffer = New clsBuffer
        
        Buffer.WriteLong CSpells
        
        SendData Buffer.ToArray()
        
        Set Buffer = Nothing
    End If
End Sub

Private Sub picStatsButton_Click()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong CGetStats
    
    SendData Buffer.ToArray()
    
    Set Buffer = Nothing
    
End Sub

Private Sub picTrainButton_Click()
    frmTraining.Show vbModal
End Sub

Private Sub picTradeButton_Click()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong CTrade
    
    SendData Buffer.ToArray()
    
    Set Buffer = Nothing

End Sub

Private Sub picQuitButton_Click()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    isLogging = True
    InGame = False
    DirectMusic_StopMidi
    
    Buffer.WriteLong CQuit
    
    SendData Buffer.ToArray()
    
    Call DestroyTCP
    
    Set Buffer = Nothing
    
End Sub

' **********************
' ** MAP EDITOR STUFF **
' **********************

Private Sub picScreen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If InMapEditor Then
        Call MapEditorMouseDown(Button)
    Else
        Call PlayerSearch(CurX, CurY)
    End If
    
    Call SetFocusOnChat
End Sub

Private Sub picScreen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CurX = TileView.Left + ((X + Camera.Left) \ PIC_X)
    CurY = TileView.Top + ((Y + Camera.Top) \ PIC_Y)
    
    If InMapEditor Then
        shpLoc.Visible = False
        
        If Button = vbLeftButton Or Button = vbRightButton Then
            Call MapEditorMouseDown(Button)
        End If
    End If

End Sub

Private Sub optLayers_Click()
    If optLayers.Value Then
        fraLayers.Visible = True
        fraAttribs.Visible = False
    End If
End Sub

Private Sub optAttribs_Click()
    If optAttribs.Value Then
        fraLayers.Visible = False
        fraAttribs.Visible = True
    End If
End Sub

Private Sub picBackSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MapEditorChooseTile(Button, X, Y)
End Sub
 
Private Sub picBackSelect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    shpLoc.Top = (Y \ PIC_Y) * PIC_Y
    shpLoc.Left = (X \ PIC_X) * PIC_X
    
    shpLoc.Visible = True
End Sub

Private Sub cmdSend_Click()
    Call MapEditorSend
End Sub

Private Sub cmdCancel_Click()
    Call MapEditorCancel
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
    Call MapEditorTileScroll
End Sub

Private Sub cmdFill_Click()
    MapEditorFillLayer
End Sub

Private Sub cmdClear_Click()
    Call MapEditorClearLayer
End Sub

Private Sub cmdClear2_Click()
    Call MapEditorClearAttribs
End Sub

Private Sub scrlTileSet_Change()

    frmMirage.scrlPicture.Max = (frmMirage.picBackSelect.Height \ PIC_Y) - (frmMirage.picBack.Height \ PIC_Y)

    Map.TileSet = scrlTileSet.Value
    lblTileset = scrlTileSet.Value
    
    Call InitTileSurf(scrlTileSet)
    
    Call BltMapEditor
    Call BltMapEditorTilePreview
End Sub
