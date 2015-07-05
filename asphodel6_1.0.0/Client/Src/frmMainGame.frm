VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMainGame 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asphodel Source "
   ClientHeight    =   10200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17355
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMainGame.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   680
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1157
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picSpellDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1605
      Left            =   4800
      ScaleHeight     =   107
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   148
      TabIndex        =   34
      Top             =   0
      Visible         =   0   'False
      Width           =   2220
      Begin VB.PictureBox picSpellDescBottom 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   120
         Left            =   0
         ScaleHeight     =   120
         ScaleWidth      =   2220
         TabIndex        =   117
         Top             =   1440
         Width           =   2220
      End
      Begin VB.Label lblSpellDescription 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   945
         TabIndex        =   116
         Top             =   480
         UseMnemonic     =   0   'False
         Width           =   405
      End
      Begin VB.Label lblSpell 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
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
         Height          =   375
         Left            =   120
         TabIndex        =   115
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   2055
      End
   End
   Begin VB.PictureBox picItemDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1605
      Left            =   2400
      ScaleHeight     =   107
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   148
      TabIndex        =   36
      Top             =   0
      Visible         =   0   'False
      Width           =   2220
      Begin VB.PictureBox picItemDescBottom 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   120
         Left            =   0
         ScaleHeight     =   120
         ScaleWidth      =   2220
         TabIndex        =   114
         Top             =   1440
         Width           =   2220
      End
      Begin VB.Label lblDescription 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   945
         TabIndex        =   95
         Top             =   480
         UseMnemonic     =   0   'False
         Width           =   405
      End
      Begin VB.Label lblItemName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
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
         Height          =   375
         Left            =   120
         TabIndex        =   37
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   2055
      End
   End
   Begin VB.PictureBox picShop 
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
      Height          =   7200
      Left            =   345
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   104
      Top             =   585
      Visible         =   0   'False
      Width           =   9600
      Begin VB.PictureBox picShopList 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H009F8369&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3975
         Left            =   4395
         ScaleHeight     =   265
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   265
         TabIndex        =   109
         Top             =   1665
         Width           =   3975
         Begin VB.Shape shpSelect 
            BorderColor     =   &H00BBA896&
            Height          =   510
            Left            =   15
            Top             =   15
            Visible         =   0   'False
            Width           =   510
         End
      End
      Begin VB.Label lblSell 
         BackStyle       =   0  'Transparent
         Height          =   360
         Left            =   1305
         TabIndex        =   113
         Top             =   5475
         Width           =   1155
      End
      Begin VB.Label lblRepair 
         BackStyle       =   0  'Transparent
         Height          =   360
         Left            =   1305
         TabIndex        =   112
         Top             =   5070
         Width           =   1155
      End
      Begin VB.Label lblPurchaseItem 
         BackStyle       =   0  'Transparent
         Height          =   360
         Left            =   1305
         TabIndex        =   111
         Top             =   4650
         Width           =   1155
      End
      Begin VB.Label lblWelcome 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   480
         TabIndex        =   110
         Top             =   240
         Width           =   8775
      End
      Begin VB.Label lblShopItem 
         BackStyle       =   0  'Transparent
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   900
         TabIndex        =   108
         Top             =   1350
         Width           =   2055
      End
      Begin VB.Label lblShopCost 
         BackStyle       =   0  'Transparent
         Caption         =   "Nothing"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   900
         TabIndex        =   107
         Top             =   2100
         Width           =   2055
      End
      Begin VB.Label lblShopDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "No item selected."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   1575
         Left            =   900
         TabIndex        =   106
         Top             =   2880
         Width           =   2055
      End
      Begin VB.Label lblQuitShop 
         BackStyle       =   0  'Transparent
         Height          =   360
         Left            =   1305
         TabIndex        =   105
         Top             =   5865
         Width           =   1155
      End
   End
   Begin VB.PictureBox picSign 
      BackColor       =   &H007D6853&
      BorderStyle     =   0  'None
      Height          =   1950
      Left            =   120
      ScaleHeight     =   1950
      ScaleWidth      =   8535
      TabIndex        =   59
      Top             =   8160
      Visible         =   0   'False
      Width           =   8535
      Begin VB.Label lblPressEnter 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "[Press Enter]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   240
         TabIndex        =   61
         Top             =   1440
         Visible         =   0   'False
         Width           =   8055
      End
      Begin VB.Label lblSignText 
         BackStyle       =   0  'Transparent
         Caption         =   "[empty]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   1335
         Left            =   240
         TabIndex        =   60
         Top             =   120
         Width           =   8055
      End
   End
   Begin VB.PictureBox picMP 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   10485
      ScaleHeight     =   195
      ScaleWidth      =   2400
      TabIndex        =   94
      Top             =   1830
      Width           =   2400
   End
   Begin VB.PictureBox picHP 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   10485
      ScaleHeight     =   195
      ScaleWidth      =   2400
      TabIndex        =   93
      Top             =   1320
      Width           =   2400
   End
   Begin VB.PictureBox picTNL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   315
      ScaleHeight     =   330
      ScaleWidth      =   9660
      TabIndex        =   92
      Top             =   7890
      Width           =   9660
   End
   Begin VB.PictureBox picCheckSize 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   1.50000e5
      ScaleHeight     =   15
      ScaleWidth      =   15
      TabIndex        =   83
      Top             =   6360
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.PictureBox picMapEditor 
      Appearance      =   0  'Flat
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
      Height          =   10200
      Left            =   13500
      ScaleHeight     =   680
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   257
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   3855
      Begin VB.HScrollBar scrlRight 
         Height          =   255
         Left            =   120
         Max             =   0
         TabIndex        =   70
         Top             =   5760
         Width           =   3615
      End
      Begin VB.OptionButton optAttribs 
         Caption         =   "Attributes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   10
         Top             =   9480
         Width           =   1575
      End
      Begin VB.Frame frmTileSet 
         Caption         =   "TileSet"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   21
         Top             =   6120
         Width           =   3615
         Begin VB.HScrollBar scrlTileSet 
            Height          =   255
            Left            =   120
            Max             =   0
            TabIndex        =   22
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label lblTileset 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3120
            TabIndex        =   23
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.OptionButton optLayers 
         Caption         =   "Layers"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   9480
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   9
         Top             =   9840
         Width           =   855
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   8
         Top             =   9840
         Width           =   735
      End
      Begin VB.CommandButton cmdProperties 
         Caption         =   "Properties"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   9840
         Width           =   1455
      End
      Begin VB.VScrollBar scrlPicture 
         Height          =   5655
         Left            =   3480
         Max             =   0
         TabIndex        =   6
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
         Height          =   5640
         Left            =   120
         ScaleHeight     =   376
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   224
         TabIndex        =   4
         Top             =   120
         Width           =   3360
         Begin VB.PictureBox picBackSelect 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
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
            TabIndex        =   5
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
      Begin VB.Frame FraLayers 
         Caption         =   "Layers"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   120
         TabIndex        =   24
         Top             =   6840
         Width           =   3615
         Begin VB.OptionButton optLayer 
            Caption         =   "Fringe (3)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   30
            Top             =   1440
            Width           =   1575
         End
         Begin VB.OptionButton optLayer 
            Caption         =   "Animation (2)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   29
            Top             =   1200
            Width           =   1935
         End
         Begin VB.OptionButton optLayer 
            Caption         =   "Mask (1)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   28
            Top             =   960
            Width           =   1455
         End
         Begin VB.OptionButton optLayer 
            Caption         =   "Ground (0)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   27
            Top             =   720
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.CommandButton cmdFill 
            Caption         =   "Fill"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   2040
            TabIndex        =   26
            Top             =   2160
            Width           =   735
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   2760
            TabIndex        =   25
            Top             =   2160
            Width           =   735
         End
      End
      Begin VB.Frame fraAttribs 
         Caption         =   "Attributes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   120
         TabIndex        =   12
         Top             =   6840
         Visible         =   0   'False
         Width           =   3615
         Begin VB.OptionButton optAttrib 
            Caption         =   "DOT (11)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   11
            Left            =   1920
            TabIndex        =   119
            Top             =   960
            Width           =   1455
         End
         Begin VB.OptionButton optAttrib 
            Caption         =   "Heal (10)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   10
            Left            =   1920
            TabIndex        =   118
            Top             =   720
            Width           =   1455
         End
         Begin VB.OptionButton optAttrib 
            Caption         =   "None (0)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   82
            Top             =   240
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton optAttrib 
            Caption         =   "Guild (9)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   9
            Left            =   1920
            TabIndex        =   81
            Top             =   480
            Width           =   1455
         End
         Begin VB.OptionButton optAttrib 
            Caption         =   "Sign (8)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   8
            Left            =   1920
            TabIndex        =   58
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdClear2 
            Caption         =   "Clear"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   2760
            TabIndex        =   16
            Top             =   2160
            Width           =   735
         End
         Begin VB.CommandButton cmdFill2 
            Caption         =   "Fill"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   2040
            TabIndex        =   32
            Top             =   2160
            Width           =   735
         End
         Begin VB.OptionButton optAttrib 
            Caption         =   "Shop (7)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   7
            Left            =   120
            TabIndex        =   31
            Top             =   1920
            Width           =   1335
         End
         Begin VB.OptionButton optAttrib 
            Caption         =   "Key Open (6)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   6
            Left            =   120
            TabIndex        =   20
            Top             =   1680
            Width           =   1935
         End
         Begin VB.OptionButton optAttrib 
            Caption         =   "Blocked (1)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   18
            Top             =   480
            Width           =   1695
         End
         Begin VB.OptionButton optAttrib 
            Caption         =   "Warp (2)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   17
            Top             =   720
            Width           =   1455
         End
         Begin VB.OptionButton optAttrib 
            Caption         =   "Item (3)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   3
            Left            =   120
            TabIndex        =   15
            Top             =   960
            Width           =   1455
         End
         Begin VB.OptionButton optAttrib 
            Caption         =   "Npc Avoid (4)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   4
            Left            =   120
            TabIndex        =   14
            Top             =   1200
            Width           =   1935
         End
         Begin VB.OptionButton optAttrib 
            Caption         =   "Key (5)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   5
            Left            =   120
            TabIndex        =   13
            Top             =   1440
            Width           =   1215
         End
      End
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   1140
      Left            =   360
      TabIndex        =   2
      Top             =   8685
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   2011
      _Version        =   393217
      BackColor       =   4210752
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMainGame.frx":08CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      BackColor       =   &H00FFFFFF&
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
      Height          =   7200
      Left            =   345
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   0
      Top             =   585
      Width           =   9600
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   13080
      Top             =   9840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox txtMyChat 
      Height          =   300
      Left            =   360
      TabIndex        =   91
      Top             =   8310
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   529
      _Version        =   393217
      BackColor       =   4210752
      MultiLine       =   0   'False
      ScrollBars      =   2
      MaxLength       =   100
      Appearance      =   0
      TextRTF         =   $"frmMainGame.frx":0945
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picStatWindow 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5265
      Left            =   10215
      ScaleHeight     =   351
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   199
      TabIndex        =   38
      Top             =   2730
      Visible         =   0   'False
      Width           =   2985
      Begin VB.PictureBox picEquipment 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   4
         Left            =   2160
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   53
         Top             =   4440
         Width           =   480
      End
      Begin VB.PictureBox picEquipment 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   2
         Left            =   960
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   52
         Top             =   4440
         Width           =   480
      End
      Begin VB.PictureBox picEquipment 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   3
         Left            =   1560
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   51
         Top             =   4440
         Width           =   480
      End
      Begin VB.PictureBox picEquipment 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   1
         Left            =   360
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   50
         Top             =   4440
         Width           =   480
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "shield"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   2160
         TabIndex        =   57
         Top             =   4320
         Width           =   375
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "helmet"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   1560
         TabIndex        =   56
         Top             =   4320
         Width           =   375
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "plate"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   960
         TabIndex        =   55
         Top             =   4320
         Width           =   375
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "sword"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   360
         TabIndex        =   54
         Top             =   4320
         Width           =   375
      End
      Begin VB.Label lblPoints 
         BackStyle       =   0  'Transparent
         Caption         =   "Points: 0"
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
         TabIndex        =   49
         Top             =   1560
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblStatAdd 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "[+]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   360
         TabIndex        =   48
         Top             =   2520
         Width           =   255
      End
      Begin VB.Label lblStatAdd 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "[+]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   360
         TabIndex        =   47
         Top             =   2880
         Width           =   255
      End
      Begin VB.Label lblStatAdd 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "[+]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   360
         TabIndex        =   46
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label lblStatAdd 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "[+]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   45
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label lblStat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stat: (0/0)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   840
         TabIndex        =   44
         Top             =   2550
         Width           =   945
      End
      Begin VB.Label lblStat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stat: (0/0)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   840
         TabIndex        =   43
         Top             =   2910
         Width           =   945
      End
      Begin VB.Label lblStat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stat: (0/0)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   840
         TabIndex        =   42
         Top             =   2190
         Width           =   945
      End
      Begin VB.Label lblStat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stat: (0/0)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   840
         TabIndex        =   41
         Top             =   1830
         Width           =   945
      End
      Begin VB.Label lblClassLevel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Level 0 (blank)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   0
         TabIndex        =   40
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label lblPlayerName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "[ (none) ]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   39
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.PictureBox picPlayerSpells 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   5265
      Left            =   10215
      ScaleHeight     =   351
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   199
      TabIndex        =   19
      Top             =   2730
      Visible         =   0   'False
      Width           =   2985
      Begin VB.PictureBox picSpellWaiting 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   9
         Left            =   1080
         ScaleHeight     =   480
         ScaleWidth      =   45
         TabIndex        =   80
         Top             =   4170
         Width           =   45
      End
      Begin VB.PictureBox picSpellWaiting 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   8
         Left            =   1080
         ScaleHeight     =   480
         ScaleWidth      =   45
         TabIndex        =   79
         Top             =   3660
         Width           =   45
      End
      Begin VB.PictureBox picSpellWaiting 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   7
         Left            =   1080
         ScaleHeight     =   480
         ScaleWidth      =   45
         TabIndex        =   78
         Top             =   3150
         Width           =   45
      End
      Begin VB.PictureBox picSpellWaiting 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   6
         Left            =   1080
         ScaleHeight     =   480
         ScaleWidth      =   45
         TabIndex        =   77
         Top             =   2640
         Width           =   45
      End
      Begin VB.PictureBox picSpellWaiting 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   5
         Left            =   1080
         ScaleHeight     =   480
         ScaleWidth      =   45
         TabIndex        =   76
         Top             =   2130
         Width           =   45
      End
      Begin VB.PictureBox picSpellWaiting 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   4
         Left            =   1080
         ScaleHeight     =   480
         ScaleWidth      =   45
         TabIndex        =   75
         Top             =   1620
         Width           =   45
      End
      Begin VB.PictureBox picSpellWaiting 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   3
         Left            =   1080
         ScaleHeight     =   480
         ScaleWidth      =   45
         TabIndex        =   74
         Top             =   1110
         Width           =   45
      End
      Begin VB.PictureBox picSpellWaiting 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   2
         Left            =   1080
         ScaleHeight     =   480
         ScaleWidth      =   45
         TabIndex        =   73
         Top             =   600
         Width           =   45
      End
      Begin VB.PictureBox picSpellWaiting 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   10
         Left            =   1080
         ScaleHeight     =   480
         ScaleWidth      =   45
         TabIndex        =   72
         Top             =   4680
         Width           =   45
      End
      Begin VB.PictureBox picSpellWaiting 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   1
         Left            =   1080
         ScaleHeight     =   480
         ScaleWidth      =   45
         TabIndex        =   71
         Top             =   90
         Width           =   45
      End
      Begin VB.PictureBox picSpellList 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5130
         Left            =   1200
         ScaleHeight     =   5130
         ScaleWidth      =   540
         TabIndex        =   33
         Top             =   60
         Width           =   540
      End
   End
   Begin VB.PictureBox picOptions 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
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
      Height          =   5265
      Left            =   10215
      ScaleHeight     =   351
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   199
      TabIndex        =   96
      Top             =   2730
      Visible         =   0   'False
      Width           =   2985
      Begin VB.CheckBox chkNPCNames 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "NPC Names"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   720
         TabIndex        =   103
         Top             =   2640
         Width           =   1455
      End
      Begin VB.CheckBox chkPlayerNames 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Player Names"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   720
         TabIndex        =   102
         Top             =   2280
         Width           =   1815
      End
      Begin VB.CheckBox chkPing 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Ping"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   720
         TabIndex        =   101
         Top             =   1920
         Width           =   975
      End
      Begin VB.CheckBox chkFPS 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "FPS"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   720
         TabIndex        =   100
         Top             =   1560
         Width           =   975
      End
      Begin VB.CheckBox chkSound 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Sound"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   720
         TabIndex        =   99
         Top             =   1200
         Width           =   975
      End
      Begin VB.CheckBox chkMusic 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Music"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   720
         TabIndex        =   98
         Top             =   840
         Width           =   975
      End
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
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
      Height          =   5265
      Left            =   10215
      ScaleHeight     =   351
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   199
      TabIndex        =   1
      Top             =   2730
      Visible         =   0   'False
      Width           =   2985
      Begin VB.PictureBox picInventoryList 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2985
         Left            =   480
         ScaleHeight     =   2985
         ScaleWidth      =   1995
         TabIndex        =   35
         Top             =   1080
         Width           =   1995
      End
   End
   Begin VB.PictureBox picGuildCP 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5265
      Left            =   10215
      ScaleHeight     =   351
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   199
      TabIndex        =   62
      Top             =   2730
      Visible         =   0   'False
      Width           =   2985
      Begin VB.TextBox txtInvite 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   480
         TabIndex        =   65
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label lblKick 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "[kick]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   69
         Top             =   3360
         Width           =   3015
      End
      Begin VB.Label lblDemote 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "[demote]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   68
         Top             =   3000
         Width           =   3015
      End
      Begin VB.Label lblPromote 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "[promote]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   67
         Top             =   2640
         Width           =   3015
      End
      Begin VB.Label lblInvite 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "[invite]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   66
         Top             =   2280
         Width           =   3015
      End
      Begin VB.Label lblDisband 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "disband"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   64
         Top             =   4680
         Width           =   3015
      End
      Begin VB.Label lblGuildName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "[ (none) ]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   63
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.Label lblOptions 
      BackStyle       =   0  'Transparent
      Height          =   360
      Left            =   11820
      TabIndex        =   97
      Top             =   8955
      Width           =   1155
   End
   Begin VB.Label picGuild 
      BackStyle       =   0  'Transparent
      Height          =   360
      Left            =   11820
      TabIndex        =   90
      Top             =   8550
      Width           =   1155
   End
   Begin VB.Label picQuit 
      BackStyle       =   0  'Transparent
      Height          =   360
      Left            =   11820
      TabIndex        =   89
      Top             =   9360
      Width           =   1155
   End
   Begin VB.Label picTrain 
      BackStyle       =   0  'Transparent
      Height          =   360
      Left            =   10455
      TabIndex        =   88
      Top             =   9360
      Width           =   1155
   End
   Begin VB.Label picSpells 
      BackStyle       =   0  'Transparent
      Height          =   360
      Left            =   10455
      TabIndex        =   87
      Top             =   8955
      Width           =   1155
   End
   Begin VB.Label picInventory 
      BackStyle       =   0  'Transparent
      Height          =   360
      Left            =   10455
      TabIndex        =   86
      Top             =   8550
      Width           =   1155
   End
   Begin VB.Label lblMP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   10920
      TabIndex        =   85
      Top             =   1590
      Width           =   1935
   End
   Begin VB.Label lblHP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   10920
      TabIndex        =   84
      Top             =   1050
      Width           =   1935
   End
End
Attribute VB_Name = "frmMainGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ------------------------------------------
' --               Asphodel               --
' ------------------------------------------

Private Sub Form_Load()

    SetWindowLong txtMyChat.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT
    SetWindowLong txtChat.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT
    
    frmMainGame.Width = Default_MenuWidth
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ResetWindows
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DestroyGame
End Sub

Private Sub chkFPS_Click()
    BFPS = chkFPS.Value
End Sub

Private Sub chkPing_Click()
    PingEnabled = chkPing.Value
    PutVar App.Path & "/info.ini", "BASIC", "Ping", CStr(PingEnabled)
End Sub

Private Sub chkPlayerNames_Click()

    ShowPNames = chkPlayerNames.Value
    
    If ShowPNames Then
        PutVar App.Path & "\info.ini", "BASIC", "ShowPlayerNames", CStr(1)
    Else
        PutVar App.Path & "\info.ini", "BASIC", "ShowPlayerNames", CStr(0)
    End If
    
End Sub

Private Sub chkNPCNames_Click()

    ShowNNames = chkNPCNames.Value
    
    If ShowNNames Then
        PutVar App.Path & "\info.ini", "BASIC", "ShowNPCNames", CStr(1)
    Else
        PutVar App.Path & "\info.ini", "BASIC", "ShowNPCNames", CStr(0)
    End If
    
End Sub

Private Sub chkMusic_Click()

    Music_On = chkMusic.Value
    
    If Music_On Then
        InitDirectMusic
        DirectMusic_StopMidi
        DirectMusic_PlayMidi Trim$(Map.Music) & MUSIC_EXT
    Else
        DestroyDirectMusic
    End If
    
    If Music_On Then
        PutVar App.Path & "\info.ini", "BASIC", "Music_On", CStr(1)
    Else
        PutVar App.Path & "\info.ini", "BASIC", "Music_On", CStr(0)
    End If
    
End Sub

Private Sub chkSound_Click()

    Sound_On = chkSound.Value
    
    If Sound_On Then InitDirectSound Else DestroyDirectSound
    
    If Sound_On Then
        PutVar App.Path & "\info.ini", "BASIC", "Sound_On", CStr(1)
    Else
        PutVar App.Path & "\info.ini", "BASIC", "Sound_On", CStr(0)
    End If
    
End Sub

Private Sub lblDisband_Click()
    If Player(MyIndex).GuildRank = 4 Then
        If MsgBox("Are you sure you want to disband your guild?", vbCritical + vbYesNo, "Confirm") = vbYes Then SendData CGuildDisband & END_CHAR
    Else
        If MsgBox("Are you sure you want to leave your guild?", vbCritical + vbYesNo, "Confirm") = vbYes Then SendData CGuildDisband & END_CHAR
    End If
End Sub

Private Sub lblInvite_Click()
    SendData CGuildInvite & SEP_CHAR & txtInvite.Text & END_CHAR
End Sub

Private Sub lblPlayerName_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MovePicture(picStatWindow, Button, X, Y)
End Sub

Private Sub lblPlayerName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SOffsetX = X
    SOffsetY = Y
End Sub

Private Sub lblPromote_Click()
    SendData CGuildPromoteDemote & SEP_CHAR & 1 & SEP_CHAR & txtInvite.Text & END_CHAR
End Sub

Private Sub lblDemote_Click()
    SendData CGuildPromoteDemote & SEP_CHAR & 2 & SEP_CHAR & txtInvite.Text & END_CHAR
End Sub

Private Sub lblKick_Click()
    SendData CGuildPromoteDemote & SEP_CHAR & 3 & SEP_CHAR & txtInvite.Text & END_CHAR
End Sub

Private Sub lblPurchaseItem_Click()
    If ShopTrade.BuyItem > 0 Then SendData CTradeRequest & SEP_CHAR & ShopTrade.BuyItem & END_CHAR
End Sub

Private Sub lblQuitShop_Click()
    GetRidOfShop
End Sub

Private Sub lblRepair_Click()

    If ReadyToSell Then ReadyToSell = False
    
    ReadyToRepair = True
    AddText "Click an item in your inventory you wish to repair!", Color.Black
    
End Sub

Private Sub lblSell_Click()

    If ReadyToRepair Then ReadyToRepair = False
    
    ReadyToSell = True
    AddText "Click an item in your inventory you wish to sell!", Color.Black
    
End Sub

Private Sub lblStatAdd_Click(Index As Integer)
    SendData CUseStatPoint & SEP_CHAR & Index & END_CHAR
End Sub

Private Sub picEquipment_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    InvPosX = X / Screen.TwipsPerPixelX
    InvPosY = Y / Screen.TwipsPerPixelY
    
    If GetPlayerEquipmentSlot(MyIndex, Index) > 0 Then
    
        lblItemName.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, GetPlayerEquipmentSlot(MyIndex, Index))).Name)
        
        UpdateItemDescription GetPlayerInvItemNum(MyIndex, GetPlayerEquipmentSlot(MyIndex, Index)), 0
        
        picItemDesc.Top = picStatWindow.Top
        picItemDesc.Left = picStatWindow.Left - picItemDesc.Width
        
        picItemDesc.Visible = True
        Exit Sub
        
    End If
    
    picItemDesc.Visible = False
    
End Sub

Private Sub picGuild_Click()

    If LenB(Trim$(Player(MyIndex).GuildName)) < 1 Then
        AddText "You aren't in a guild!", AlertColor
        Exit Sub
    End If
    
    frmMainGame.picGuildCP.Visible = Not frmMainGame.picGuildCP.Visible
    
    If picGuildCP.Visible Then
        If picPlayerSpells.Visible Then picPlayerSpells.Visible = False
        If picInv.Visible Then picInv.Visible = False
        If picStatWindow.Visible Then picStatWindow.Visible = False
        If picOptions.Visible Then picOptions.Visible = False
    End If
    
End Sub

Private Sub lblOptions_Click()

    frmMainGame.picOptions.Visible = Not frmMainGame.picOptions.Visible
    
    If picOptions.Visible Then
        If picPlayerSpells.Visible Then picPlayerSpells.Visible = False
        If picInv.Visible Then picInv.Visible = False
        If picStatWindow.Visible Then picStatWindow.Visible = False
        If picGuildCP.Visible Then picGuildCP.Visible = False
    End If
    
End Sub

Private Sub picShop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ResetWindows
End Sub

Private Sub picShopList_Click()
Dim SetItem As Long
Dim rec_pos As DxVBLib.RECT

    SetItem = IsShopItem(ShopPosX, ShopPosY)
    
    If SetItem > 0 Then
        lblShopItem.Caption = ShopTrade.TradeItem(SetItem).GetValue & " " & Trim$(Item(ShopTrade.TradeItem(SetItem).GetItem).Name)
        
        If ShopTrade.TradeItem(SetItem).GiveValue = 0 Then
            lblShopCost.Caption = "Nothing"
        Else
            lblShopCost.Caption = ShopTrade.TradeItem(SetItem).GiveValue & " " & Trim$(Item(ShopTrade.TradeItem(SetItem).GiveItem).Name)
        End If
        
        ShopTrade.BuyItem = SetItem
        
        rec_pos = Get_RECT(ShopIconY + ((ShopOffsetY + PIC_Y) * ((SetItem - 1) \ ShopIconsInRow)), ShopIconY + ((ShopOffsetX + PIC_X) * (((SetItem - 1) Mod ShopIconsInRow))))
        
        shpSelect.Top = rec_pos.Top - 1
        shpSelect.Left = rec_pos.Left - 1
        shpSelect.Visible = True
        
    Else
        lblShopItem.Caption = "None"
        lblShopCost.Caption = "Nothing"
        ShopTrade.BuyItem = 0
        
        shpSelect.Top = 1
        shpSelect.Left = 1
        shpSelect.Visible = False
        
    End If
    
End Sub

Private Sub picShopList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ItemSlot As Long

    ShopPosX = X
    ShopPosY = Y
    
    ItemSlot = IsShopItem(ShopPosX, ShopPosY)
    
    If ItemSlot > 0 Then
        lblItemName.Caption = Trim$(Item(ShopTrade.TradeItem(ItemSlot).GetItem).Name)
        
        UpdateItemDescription ShopTrade.TradeItem(ItemSlot).GetItem, ShopTrade.TradeItem(ItemSlot).GetValue
        
        picItemDesc.Top = picShopList.Top + picShop.Top
        picItemDesc.Left = ((picShopList.Left + picShop.Left) - picItemDesc.Width) - 6
        
        picItemDesc.Visible = True
        Exit Sub
    End If
    
    picItemDesc.Visible = False
    
End Sub

Private Sub picStatWindow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ResetWindows
End Sub

Private Sub picTrain_Click()

    picStatWindow.Visible = Not picStatWindow.Visible
    
    If picStatWindow.Visible Then
        If picPlayerSpells.Visible Then picPlayerSpells.Visible = False
        If picGuildCP.Visible Then picGuildCP.Visible = False
        If picInv.Visible Then picInv.Visible = False
        If picOptions.Visible Then picOptions.Visible = False
    End If
    
End Sub

' Winsock event
Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
    If IsConnected Then IncomingData bytesTotal
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If picSign.Visible Then
        If lblSignText.Caption = ScrollText.Text Then
            If KeyAscii = vbKeyReturn Then
                SendData CPressReturn & SEP_CHAR & ScrollText.KeyValue & SEP_CHAR & ScrollText.CurKey & END_CHAR
                KeyAscii = 0
                Exit Sub
            End If
        ElseIf ScrollText.Running Then
            If KeyAscii = vbKeyReturn Then
                lblSignText.Caption = ScrollText.Text
                ScrollText.Running = False
                KeyAscii = 0
                Exit Sub
            End If
        End If
    End If
    
    HandleKeypresses KeyAscii
    
    ' prevents textbox on error ding sound
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    CheckInput 1, KeyCode
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    CheckInput 0, KeyCode
End Sub

Private Sub picScreen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If InEditor Then
        If SettingSpawn Then MapEditorSetSpawn: Exit Sub
        MapEditorMouseDown Button
        Exit Sub
    Else
        If Button = vbLeftButton Then PlayerSearch
    End If
    
    If Button = vbRightButton Then
        If GetPlayerAccess(MyIndex) >= StaffType.Mapper Then
            If frmAdmin.chkRightClick.Value = 1 Then
                SendData CRCWarp & SEP_CHAR & CurX & SEP_CHAR & CurY & END_CHAR
            End If
        End If
    End If
    
    SetFocusOnChat
    
End Sub

Private Sub picScreen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    CurX = X \ PIC_X
    CurY = Y \ PIC_Y
    
    If InEditor Then
        shpLoc.Visible = False
        
        If Button = vbLeftButton Or Button = vbRightButton Then
            If SettingSpawn Then Exit Sub
            MapEditorMouseDown (Button)
        End If
    End If
    
    ResetWindows
    
End Sub

Private Sub picSpellList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim SetSpell As Long

    If Shift <> 1 Then Exit Sub
    
    SetSpell = IsSpell(IconPosX, IconPosY)
    
    If SetSpell > 0 Then
        If SetSpell <> Selected_Spell Then
            Selected_Spell = SetSpell
            If Player(MyIndex).Spell(Selected_Spell) > 0 Then
                AddText Trim$(Spell(Player(MyIndex).Spell(Selected_Spell)).Name) & " has been memorized. Hit insert to cast!", Color.blue
            End If
        End If
    End If
    
End Sub

Private Sub picInventoryList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim SetItem As Long

    SetItem = IsItem(InvPosX, InvPosY)
    
    If SetItem < 1 Then Exit Sub
    
    If Button = vbRightButton Then
        If GetPlayerInvItemNum(MyIndex, SetItem) > 0 Then
            If GetPlayerInvItemNum(MyIndex, SetItem) <= MAX_ITEMS Then
                If Item(GetPlayerInvItemNum(MyIndex, SetItem)).Type = ItemType.Currency_ Then
                    ' Show them the drop dialog
                    frmDrop.Show vbModal
                Else
                    SendDropItem SetItem, 0
                End If
            End If
        End If
    End If
    
End Sub

Private Sub picInventoryList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ItemSlot As Long

    InvPosX = X / Screen.TwipsPerPixelX
    InvPosY = Y / Screen.TwipsPerPixelY
    
    ItemSlot = IsItem(InvPosX, InvPosY)
    
    If ItemSlot > 0 Then
        If GetPlayerInvItemNum(MyIndex, ItemSlot) > 0 Then
        
            lblItemName.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, ItemSlot)).Name)
            
            UpdateItemDescription GetPlayerInvItemNum(MyIndex, ItemSlot), GetPlayerInvItemValue(MyIndex, ItemSlot)
            
            picItemDesc.Top = picInv.Top + picInventoryList.Top
            picItemDesc.Left = picInv.Left - picItemDesc.Width
            
            picItemDesc.Visible = True
            
        End If
        Exit Sub
    End If
    
    picItemDesc.Visible = False
    
End Sub

Private Sub picInventoryList_DblClick()
Dim SetItem As Long

    SetItem = IsItem(InvPosX, InvPosY)
    
    If SetItem > 0 Then
        SendUseItem SetItem
        Exit Sub
    End If
    
End Sub

Private Sub picInventoryList_Click()
Dim SetItem As Long
Dim LoopI As Long

    SetItem = IsItem(InvPosX, InvPosY)
    
    If SetItem > 0 Then
        If ReadyToRepair Then
            If Not ItemIsEquipment(GetPlayerInvItemNum(MyIndex, SetItem)) Then
                AddText "Only equipment is allowed to be repaired!", Color.BrightRed
                ReadyToRepair = False
                Exit Sub
            End If
            
            SendData CFixItem & SEP_CHAR & SetItem & END_CHAR
            ReadyToRepair = False
            Exit Sub
        End If
        
        If ReadyToSell Then
            SendData CSellItem & SEP_CHAR & SetItem & END_CHAR
            ReadyToSell = False
            Exit Sub
        End If
    End If
    
    If ReadyToRepair Or ReadyToSell Then AddText "No item selected! Try again.", Color.BrightRed
    
End Sub

Private Sub picSpellList_DblClick()
Dim SpellSlot As Long

    SpellSlot = IsSpell(IconPosX, IconPosY)
    
    If SpellSlot <> 0 Then
    
    If Player(MyIndex).Spell(SpellSlot) < 1 Then Exit Sub
        If Player(MyIndex).CastTimer(SpellSlot) < GetTickCountNew Then
            If Player(MyIndex).Moving = 0 Then
                SendData CCast & SEP_CHAR & SpellSlot & END_CHAR
            Else
                AddText "Cannot cast while walking!", Color.BrightRed
            End If
        End If
        Exit Sub
    End If
    
End Sub

Private Sub picSpellList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim SpellSlot As Long
Dim SpellNum As Long

    IconPosX = X / Screen.TwipsPerPixelX
    IconPosY = Y / Screen.TwipsPerPixelY
    
    SpellSlot = IsSpell(IconPosX, IconPosY)

    If SpellSlot > 0 Then
    
        SpellNum = Player(MyIndex).Spell(SpellSlot)
        
        If SpellNum > 0 Then
        
            UpdateSpellDescription SpellNum
            
            picSpellDesc.Top = picPlayerSpells.Top
            picSpellDesc.Left = picPlayerSpells.Left - picSpellDesc.Width
            
            picSpellDesc.Visible = True
        End If
        Exit Sub
    End If
    
    picSpellDesc.Visible = False
    
End Sub

Private Sub txtMyChat_Change()
    MyText = txtMyChat.Text
End Sub

Private Sub txtChat_GotFocus()
    SetFocusOnChat
End Sub

Public Sub picInventory_Click()

    picInv.Visible = Not picInv.Visible
    
    If picInv.Visible Then
        If picPlayerSpells.Visible Then picPlayerSpells.Visible = False
        If picGuildCP.Visible Then picGuildCP.Visible = False
        If picStatWindow.Visible Then picStatWindow.Visible = False
        If picOptions.Visible Then picOptions.Visible = False
        UpdateInventory
    End If
    
End Sub

Private Sub picSpells_Click()

    frmMainGame.picPlayerSpells.Visible = Not frmMainGame.picPlayerSpells.Visible
    
    If frmMainGame.picPlayerSpells.Visible Then
        If picInv.Visible Then picInv.Visible = False
        If picGuildCP.Visible Then picGuildCP.Visible = False
        If picStatWindow.Visible Then picStatWindow.Visible = False
        If picOptions.Visible Then picOptions.Visible = False
        SendData CSpells & END_CHAR
    End If
    
End Sub

Private Sub picQuit_Click()

    isLogging = True
    InGame = False
    
    SendData CQuit & END_CHAR
    
End Sub

Private Function IsSpell(ByVal X As Single, ByVal Y As Single) As Long
Dim tempRec As RECT
Dim Spell As Long

    For Spell = 1 To MAX_PLAYER_SPELLS
        If Player(MyIndex).Spell(Spell) > 0 Then
            tempRec = Get_RECT(IconY + ((IconOffsetY + PIC_Y) * ((Spell - IconsInRow) \ IconsInRow)), IconX + ((IconOffsetX + PIC_X) * (((Spell - IconsInRow) Mod IconsInRow))))
            
            If X >= tempRec.Left Then
                If X <= tempRec.Right Then
                    If Y >= tempRec.Top Then
                        If Y <= tempRec.Bottom Then
                            IsSpell = Spell
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next
    
End Function

Private Sub txtMyChat_KeyDown(KeyCode As Integer, Shift As Integer)

    Form_KeyDown KeyCode, Shift
    
    If KeyCode >= vbKeyLeft And KeyCode <= vbKeyDown Then KeyCode = 0
    If KeyCode >= 32 Or KeyCode <= 126 Or KeyCode = vbKeyBack Or KeyCode = vbKeyReturn Then Exit Sub
    
    KeyCode = 0
    
End Sub

' // MAP EDITOR STUFF //
Private Sub optLayers_Click()
    If optLayers.Value Then
        FraLayers.Visible = True
        fraAttribs.Visible = False
        picBack.Enabled = True
    End If
End Sub

Private Sub optAttribs_Click()
    If optAttribs.Value Then
        FraLayers.Visible = False
        fraAttribs.Visible = True
        picBack.Enabled = False
    End If
End Sub

Private Sub optAttrib_Click(Index As Integer)

    ClearMapAttribs
    
    Select Case Index
        Case Tile_Type.Blocked_
        Case Tile_Type.NpcAvoid_
        Case Tile_Type.Guild_
        Case Tile_Type.Heal_
        Case Tile_Type.Item_
            MapAttribFormTitle = "Editing Map Item"
            SetMapAttrib 1, "Item:", 1, MAX_ITEMS
            SetMapAttrib 2, "Value:", 1, 32767
            SetMapAttrib 3, "Anim:", 0, MAX_ANIMS
            Load frmAttrib
        Case Tile_Type.Key_
            MapAttribFormTitle = "Set Map Key"
            SetMapAttrib 1, "Item:", 1, MAX_ITEMS
            SetMapAttrib 2, "Take Key:", 0, 1
            Load frmAttrib
        Case Tile_Type.KeyOpen_
            MapAttribFormTitle = "Set Key Open"
            SetMapAttrib 1, "X:", 0, MAX_MAPX
            SetMapAttrib 2, "Y:", 0, MAX_MAPY
            Load frmAttrib
        Case Tile_Type.Shop_
            MapAttribFormTitle = "Editing Map Shop"
            SetMapAttrib 1, "Shop:", 1, MAX_SHOPS
            Load frmAttrib
        Case Tile_Type.Sign_
            MapAttribFormTitle = "Editing Map Sign"
            SetMapAttrib 1, "Sign:", 1, MAX_SIGNS
            Load frmAttrib
        Case Tile_Type.Warp_
            MapAttribFormTitle = "Editing Map Warp"
            SetMapAttrib 1, "Map:", 1, MAX_MAPS
            SetMapAttrib 2, "X:", 0, MAX_MAPX
            SetMapAttrib 3, "Y:", 0, MAX_MAPY
            Load frmAttrib
        Case Tile_Type.Damage_
            MapAttribFormTitle = "Editing Damage Over Time"
            SetMapAttrib 1, "Damage %:", 1, 100
            SetMapAttrib 2, "Time (ms):", 1, 10000
            Load frmAttrib
        Case Else
            Exit Sub
    End Select
    
    MapAttribType = Index
    
End Sub

Private Sub picBackSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MapEditorChooseTile Button, X, Y, Shift
End Sub

Private Sub picBackSelect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    shpLoc.Top = (Y \ PIC_Y) * PIC_Y
    shpLoc.Left = (X \ PIC_X) * PIC_X
    
    shpLoc.Visible = True
End Sub

Private Sub cmdSend_Click()
    MapEditorSend
End Sub

Private Sub cmdCancel_Click()
    MapEditorCancel
End Sub

Private Sub cmdProperties_Click()
On Error GoTo ErrHandler

    frmMainGame.Hide
    frmMapProperties.Show vbModal
    Exit Sub
    
ErrHandler:
    frmMapProperties.ZOrder 0
    
End Sub

Private Sub scrlPicture_Change()
    MapEditorTileScroll
End Sub

Private Sub scrlRight_Change()
    MapEditorTileScrollRight
End Sub

Private Sub scrlPicture_Scroll()
    scrlPicture_Change
End Sub

Private Sub cmdFill_Click()
    MapEditorFillLayer
End Sub

Private Sub cmdFill2_Click()
    MapEditorFillAttribs
End Sub

Private Sub cmdClear_Click()
    MapEditorClearLayer
End Sub

Private Sub cmdClear2_Click()
    MapEditorClearAttribs
End Sub

Public Sub scrlTileSet_Change()

    scrlRight.Value = 0
    lblTileset = scrlTileSet.Value
    
    InitTileSurf scrlTileSet
    
    BltMapEditor
    
    If frmMainGame.picBackSelect.Height > frmMainGame.picBack.Height Then
        frmMainGame.scrlPicture.Max = ((frmMainGame.picBackSelect.Height - frmMainGame.picBack.Height) \ PIC_Y)
    Else
        frmMainGame.scrlPicture.Max = 0
    End If
    
    If TILESHEET_WIDTH(scrlTileSet.Value) > (picBack.Width \ PIC_X) Then
        frmMainGame.scrlRight.Max = TILESHEET_WIDTH(scrlTileSet.Value) - (picBack.Width \ PIC_X)
    Else
        frmMainGame.scrlRight.Max = 0
    End If
    
    picBackSelect_MouseDown vbLeftButton, 0, 0, 0
    frmMainGame.scrlPicture.Value = 0
    
End Sub

Private Sub scrlTileSet_Scroll()
    scrlTileSet_Change
End Sub

Private Sub scrlRight_Scroll()
    scrlRight_Change
End Sub
