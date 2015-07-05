VERSION 5.00
Begin VB.Form frmEditor_Map 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Layers & Modes Picker"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14985
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditor_Map.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   303
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   999
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picAttributes 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7695
      Left            =   7320
      ScaleHeight     =   513
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   259
      TabIndex        =   25
      Top             =   120
      Visible         =   0   'False
      Width           =   3885
      Begin VB.CommandButton cmdCancel2 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   1680
         TabIndex        =   63
         Top             =   2835
         Width           =   1215
      End
      Begin VB.Frame fraSlide 
         Caption         =   "Slide"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   15
         TabIndex        =   60
         Top             =   1.11810e5
         Visible         =   0   'False
         Width           =   3000
         Begin VB.CommandButton cmdSlide 
            Caption         =   "Accept"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   61
            Top             =   2040
            Width           =   1215
         End
         Begin VB.ComboBox cmbSlide 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmEditor_Map.frx":038A
            Left            =   240
            List            =   "frmEditor_Map.frx":039A
            Style           =   2  'Dropdown List
            TabIndex        =   62
            Top             =   360
            Width           =   2895
         End
      End
      Begin VB.Frame fraSoundEffect 
         Caption         =   "Sound Effect"
         Height          =   2655
         Left            =   15
         TabIndex        =   90
         Top             =   810
         Visible         =   0   'False
         Width           =   3000
         Begin VB.CommandButton cmdSoundEffect 
            Caption         =   "Accept"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   92
            Top             =   2040
            Width           =   1215
         End
         Begin VB.ComboBox cmbSoundEffect 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmEditor_Map.frx":03B5
            Left            =   240
            List            =   "frmEditor_Map.frx":03B7
            Style           =   2  'Dropdown List
            TabIndex        =   91
            Top             =   360
            Width           =   2565
         End
      End
      Begin VB.Frame fraTrap 
         Caption         =   "Trap"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   15
         TabIndex        =   56
         Top             =   810
         Visible         =   0   'False
         Width           =   3000
         Begin VB.ComboBox cmbTrap 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmEditor_Map.frx":03B9
            Left            =   240
            List            =   "frmEditor_Map.frx":03C3
            Style           =   2  'Dropdown List
            TabIndex        =   64
            Top             =   360
            Width           =   2565
         End
         Begin VB.HScrollBar scrlDamage 
            Height          =   255
            Left            =   240
            Min             =   1
            TabIndex        =   58
            Top             =   960
            Value           =   1
            Width           =   2550
         End
         Begin VB.CommandButton cmdTrap 
            Caption         =   "Accept"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   57
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblDamage 
            Caption         =   "Amount: 1"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   59
            Top             =   720
            Width           =   2535
         End
      End
      Begin VB.Frame fraHeal 
         Caption         =   "Heal"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   15
         TabIndex        =   51
         Top             =   810
         Visible         =   0   'False
         Width           =   3000
         Begin VB.ComboBox cmbHeal 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmEditor_Map.frx":03D5
            Left            =   240
            List            =   "frmEditor_Map.frx":03DF
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   360
            Width           =   2535
         End
         Begin VB.CommandButton cmdHeal 
            Caption         =   "Accept"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   53
            Top             =   2040
            Width           =   1215
         End
         Begin VB.HScrollBar scrlHeal 
            Height          =   255
            Left            =   240
            Min             =   1
            TabIndex        =   52
            Top             =   960
            Value           =   1
            Width           =   2490
         End
         Begin VB.Label lblHeal 
            Caption         =   "Amount: 1"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   54
            Top             =   720
            Width           =   2535
         End
      End
      Begin VB.Frame fraMapItem 
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   15
         TabIndex        =   35
         Top             =   810
         Visible         =   0   'False
         Width           =   3000
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   540
            Left            =   1215
            ScaleHeight     =   34
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   34
            TabIndex        =   65
            Top             =   1290
            Width           =   540
            Begin VB.PictureBox Picture3 
               BackColor       =   &H00404040&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   480
               Left            =   15
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   66
               Top             =   15
               Width           =   480
               Begin VB.PictureBox picMapItem 
                  Appearance      =   0  'Flat
                  AutoRedraw      =   -1  'True
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   238
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   480
                  Left            =   0
                  ScaleHeight     =   32
                  ScaleMode       =   3  'Pixel
                  ScaleWidth      =   32
                  TabIndex        =   67
                  Top             =   0
                  Width           =   480
               End
            End
         End
         Begin VB.CommandButton cmdMapItem 
            Caption         =   "Accept"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   39
            Top             =   2040
            Width           =   1215
         End
         Begin VB.HScrollBar scrlMapItemValue 
            Height          =   255
            Left            =   240
            Min             =   1
            TabIndex        =   38
            Top             =   840
            Value           =   1
            Width           =   2445
         End
         Begin VB.HScrollBar scrlMapItem 
            Height          =   255
            Left            =   240
            Max             =   10
            Min             =   1
            TabIndex        =   37
            Top             =   480
            Value           =   1
            Width           =   2445
         End
         Begin VB.Label lblMapItem 
            Caption         =   "None"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   36
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.Frame fraMapWarp 
         Caption         =   "Warp"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   15
         TabIndex        =   40
         Top             =   810
         Visible         =   0   'False
         Width           =   3000
         Begin VB.CommandButton cmdMapWarp 
            Caption         =   "Accept"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   47
            Top             =   2040
            Width           =   1215
         End
         Begin VB.HScrollBar scrlMapWarpY 
            Height          =   255
            Left            =   240
            Max             =   255
            TabIndex        =   46
            Top             =   1680
            Width           =   2520
         End
         Begin VB.HScrollBar scrlMapWarpX 
            Height          =   255
            Left            =   240
            Max             =   255
            TabIndex        =   44
            Top             =   1080
            Width           =   2505
         End
         Begin VB.HScrollBar scrlMapWarp 
            Height          =   255
            Left            =   240
            Min             =   1
            TabIndex        =   42
            Top             =   480
            Value           =   1
            Width           =   2505
         End
         Begin VB.Label lblMapWarpY 
            Caption         =   "Y: 0"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   45
            Top             =   1440
            Width           =   2895
         End
         Begin VB.Label lblMapWarpX 
            Caption         =   "X: 0"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   43
            Top             =   840
            Width           =   2895
         End
         Begin VB.Label lblMapWarp 
            Caption         =   "Map: 1"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   41
            Top             =   240
            Width           =   2895
         End
      End
      Begin VB.Frame fraNpcSpawn 
         Caption         =   "Npc Spawn"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   15
         TabIndex        =   30
         Top             =   810
         Visible         =   0   'False
         Width           =   3000
         Begin VB.ListBox lstNpc 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   240
            TabIndex        =   34
            Top             =   360
            Width           =   2535
         End
         Begin VB.HScrollBar scrlNpcDir 
            Height          =   255
            Left            =   240
            Max             =   3
            TabIndex        =   32
            Top             =   1560
            Width           =   2505
         End
         Begin VB.CommandButton cmdNpcSpawn 
            Caption         =   "Accept"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   31
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblNpcDir 
            Caption         =   "Direction: Up"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   1320
            Width           =   2535
         End
      End
      Begin VB.Frame fraResource 
         Caption         =   "Resource"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   15
         TabIndex        =   26
         Top             =   810
         Visible         =   0   'False
         Width           =   3000
         Begin VB.CommandButton cmdResourceOk 
            Caption         =   "Accept"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   29
            Top             =   2040
            Width           =   1215
         End
         Begin VB.HScrollBar scrlResource 
            Height          =   255
            Left            =   240
            Max             =   100
            Min             =   1
            TabIndex        =   28
            Top             =   480
            Value           =   1
            Width           =   2505
         End
         Begin VB.Label lblResource 
            Caption         =   "None"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame fraShop 
         Caption         =   "Shop"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   15
         TabIndex        =   48
         Top             =   810
         Visible         =   0   'False
         Width           =   3000
         Begin VB.CommandButton cmdShop 
            Caption         =   "Accept"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   50
            Top             =   2040
            Width           =   1215
         End
         Begin VB.ComboBox cmbShop 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmEditor_Map.frx":03F1
            Left            =   240
            List            =   "frmEditor_Map.frx":03F3
            Style           =   2  'Dropdown List
            TabIndex        =   49
            Top             =   360
            Width           =   2550
         End
      End
   End
   Begin VB.Frame fraTileSet 
      Caption         =   "Tileset: 1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   1620
      TabIndex        =   22
      Top             =   0
      Width           =   1365
      Begin VB.CheckBox chkRandom 
         Caption         =   "Random"
         Height          =   255
         Left            =   240
         TabIndex        =   95
         ToolTipText     =   "Will place tiles you select randomly."
         Top             =   240
         Width           =   915
      End
      Begin VB.HScrollBar scrlTileSet 
         Height          =   255
         Left            =   240
         Max             =   10
         Min             =   1
         TabIndex        =   0
         Top             =   840
         Value           =   1
         Width           =   960
      End
      Begin VB.Label lblRevision 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Revision:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   68
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Height          =   405
      Left            =   1620
      TabIndex        =   93
      Top             =   1200
      Width           =   555
   End
   Begin VB.CommandButton cmdPaste 
      Caption         =   "Paste"
      Height          =   390
      Left            =   2400
      TabIndex        =   94
      Top             =   1200
      Width           =   615
   End
   Begin VB.Frame fraType 
      Caption         =   "Type"
      Height          =   1590
      Left            =   1620
      TabIndex        =   79
      Top             =   2880
      Width           =   1455
      Begin VB.OptionButton OptLayers 
         Alignment       =   1  'Right Justify
         Caption         =   "Layers"
         Height          =   255
         Left            =   360
         TabIndex        =   88
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton OptAttributes 
         Alignment       =   1  'Right Justify
         Caption         =   "Attributes"
         Height          =   255
         Left            =   120
         TabIndex        =   87
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton OptBlock 
         Alignment       =   1  'Right Justify
         Caption         =   "Block"
         Height          =   255
         Left            =   480
         TabIndex        =   86
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton OptEvents 
         Alignment       =   1  'Right Justify
         Caption         =   "Events"
         Height          =   255
         Left            =   360
         TabIndex        =   85
         Top             =   960
         Width           =   855
      End
   End
   Begin VB.Frame fraRandom 
      Caption         =   "Random"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   1620
      TabIndex        =   69
      Top             =   2955
      Visible         =   0   'False
      Width           =   1455
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   800
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   34
         TabIndex        =   80
         Top             =   240
         Width           =   540
         Begin VB.PictureBox Picture8 
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   15
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   81
            Top             =   15
            Width           =   480
            Begin VB.PictureBox picRandomTile 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   480
               Index           =   1
               Left            =   0
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   82
               Top             =   0
               Width           =   480
            End
         End
      End
      Begin VB.PictureBox Picture11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   800
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   34
         TabIndex        =   76
         Top             =   900
         Width           =   540
         Begin VB.PictureBox Picture12 
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   15
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   77
            Top             =   15
            Width           =   480
            Begin VB.PictureBox picRandomTile 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   480
               Index           =   3
               Left            =   0
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   78
               Top             =   0
               Width           =   480
            End
         End
      End
      Begin VB.PictureBox Picture9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   120
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   34
         TabIndex        =   73
         Top             =   900
         Width           =   540
         Begin VB.PictureBox Picture10 
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   15
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   74
            Top             =   15
            Width           =   480
            Begin VB.PictureBox picRandomTile 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   480
               Index           =   2
               Left            =   0
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   75
               Top             =   0
               Width           =   480
            End
         End
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   120
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   34
         TabIndex        =   70
         Top             =   240
         Width           =   540
         Begin VB.PictureBox Picture6 
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   15
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   71
            Top             =   15
            Width           =   480
            Begin VB.PictureBox picRandomTile 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   480
               Index           =   0
               Left            =   0
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   72
               Top             =   0
               Width           =   480
            End
         End
      End
   End
   Begin VB.Frame fraLayers 
      Caption         =   "Layers"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4470
      Left            =   120
      TabIndex        =   24
      Top             =   0
      Width           =   1455
      Begin VB.HScrollBar scrlAutotile 
         Height          =   255
         Left            =   255
         Max             =   5
         TabIndex        =   83
         Top             =   3105
         Width           =   975
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Fringe"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Mask"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Ground"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Roof"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         Caption         =   "Cover"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   3990
         Width           =   975
      End
      Begin VB.CommandButton cmdFill 
         Caption         =   "Fill"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   3510
         Width           =   975
      End
      Begin VB.Label lblAutoTile 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Normal"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   135
         TabIndex        =   84
         Top             =   2865
         Width           =   1215
      End
   End
   Begin VB.Frame fraAttribs 
      Caption         =   "Attributes"
      Height          =   4470
      Left            =   120
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
      Begin VB.OptionButton optSound 
         Caption         =   "Sound"
         Height          =   270
         Left            =   120
         TabIndex        =   89
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CommandButton cmdAttributeFill 
         Caption         =   "Fill"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   3540
         Width           =   975
      End
      Begin VB.OptionButton optCheckpoint 
         Caption         =   "Checkpoint"
         Height          =   270
         Left            =   120
         TabIndex        =   12
         Top             =   2880
         Width           =   1215
      End
      Begin VB.OptionButton optSlide 
         Caption         =   "Slide"
         Height          =   270
         Left            =   120
         TabIndex        =   11
         Top             =   2640
         Width           =   1215
      End
      Begin VB.OptionButton optTrap 
         Caption         =   "Trap"
         Height          =   270
         Left            =   120
         TabIndex        =   10
         Top             =   2400
         Width           =   1215
      End
      Begin VB.OptionButton optHeal 
         Caption         =   "Heal"
         Height          =   270
         Left            =   120
         TabIndex        =   9
         Top             =   2160
         Width           =   1215
      End
      Begin VB.OptionButton optBank 
         Caption         =   "Bank"
         Height          =   270
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   1215
      End
      Begin VB.OptionButton optShop 
         Caption         =   "Shop"
         Height          =   270
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   1215
      End
      Begin VB.OptionButton optNPCSpawn 
         Caption         =   "NPC Spawn"
         Height          =   270
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton optResource 
         Caption         =   "Resource"
         Height          =   240
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton optBlocked 
         Caption         =   "Blocked"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optWarp 
         Caption         =   "Warp"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdAttributeClear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   3960
         Width           =   975
      End
      Begin VB.OptionButton optItem 
         Caption         =   "Item"
         Height          =   270
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optNPCAvoid 
         Caption         =   "NPC Avoid"
         Height          =   270
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmEditor_Map"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public UnloadStarted As Boolean

Private Sub cmbHeal_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    MapEditorVitalType = cmbHeal.ListIndex + 1
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmbHeal_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmbSoundEffect_Click()
    If cmbSoundEffect.ListIndex < 0 Then Exit Sub
    Audio.StopSounds
    Audio.PlaySound cmbSoundEffect.List(cmbSoundEffect.ListIndex), -1, -1, True
End Sub

Private Sub cmdSoundEffect_Click()
   ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If cmbSoundEffect.ListIndex = 0 Then Exit Sub
    
    MapEditorSound = SoundCache(cmbSoundEffect.ListIndex)
    picAttributes.Visible = False
    fraSoundEffect.Visible = False
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdSoundEffect_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not UnloadStarted Then
        UnloadStarted = True
        LeaveMapEditorMode True
    End If
End Sub

Private Sub OptEvents_Click()
   ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    frmMain.chkGrid.Enabled = True
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "OptEvents_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub optSound_Click()
   ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ClearAttributeFrames
    picAttributes.Visible = True
    fraSoundEffect.Visible = True
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "optSound_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlAutotile_Change()
    Select Case scrlAutotile.Value
        Case 0 ' Normal
            lblAutoTile.Caption = "Normal"
        Case 1 ' Autotile
            lblAutoTile.Caption = "Autotile"
        Case 2 ' Fake autotile
            lblAutoTile.Caption = "Fake"
        Case 3 ' Animated
            lblAutoTile.Caption = "Animated"
        Case 4 ' Cliff
            lblAutoTile.Caption = "Cliff"
        Case 5 ' Waterfall
            lblAutoTile.Caption = "Waterfall"
    End Select
    
    SetMapAutotileScrollbar
End Sub

Private Sub cmbShop_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
        
    EditorShop = cmbShop.ListIndex
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmbShop_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmbTrap_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    MapEditorVitalType = cmbTrap.ListIndex + 1
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmbTrap_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdCancel2_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ClearAttributeFrames
    picAttributes.Visible = False
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdCancel2_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdHeal_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    MapEditorVitalType = cmbHeal.ListIndex + 1
    MapEditorVitalAmount = scrlHeal.Value
    picAttributes.Visible = False
    fraHeal.Visible = False
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdHeal_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdMapItem_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    ItemEditorNum = scrlMapItem.Value
    ItemEditorValue = scrlMapItemValue.Value
    picAttributes.Visible = False
    fraMapItem.Visible = False
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdMapItem_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdMapWarp_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    EditorWarpMap = scrlMapWarp.Value
    EditorWarpX = scrlMapWarpX.Value
    EditorWarpY = scrlMapWarpY.Value
    picAttributes.Visible = False
    fraMapWarp.Visible = False
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdMapWarp_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdNPCSpawn_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    SpawnNPCNum = lstNpc.ListIndex + 1
    SpawnNPCDir = scrlNpcDir.Value
    picAttributes.Visible = False
    fraNpcSpawn.Visible = False
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdNPCSpawn_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdResourceOk_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    ResourceEditorNum = scrlResource.Value
    picAttributes.Visible = False
    fraResource.Visible = False
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdResourceOk_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdShop_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    EditorShop = cmbShop.ListIndex
    picAttributes.Visible = False
    fraShop.Visible = False
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdShop_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdSlide_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    MapEditorSlideDir = cmbSlide.ListIndex
    picAttributes.Visible = False
    fraSlide.Visible = False
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdSlide_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdTrap_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    MapEditorVitalType = cmbTrap.ListIndex + 1
    MapEditorVitalAmount = scrlDamage.Value
    picAttributes.Visible = False
    fraTrap.Visible = False
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdTrap_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdAttributeFill_Click()
    Dim Button As Integer
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call MapEditorFillAttributes(Button)
    redrawMapCache = True
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdAttributeFill_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_Load()
    Dim I As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    frmEditor_Map.UnloadStarted = False
    
    ' Move the entire attributes box on screen
    picAttributes.Left = 0
    picAttributes.Top = 0
    
    ' Set maxes for attribute forms
    scrlMapItem.max = MAX_ITEMS
    scrlResource.max = MAX_RESOURCES
    scrlMapWarp.max = MAX_MAPS
    
    ' Set the width of the form
    Me.Width = 3250
    
    ' Set the max scrollbar to the number of tilesets
    frmEditor_Map.scrlTileSet.max = NumTileSets
    
    ' Populate the cache if we need to
    If Not HasPopulated Then PopulateLists

    ' Add the array to the combo
    frmEditor_Map.cmbSoundEffect.Clear
    frmEditor_Map.cmbSoundEffect.AddItem "None"

    For I = 1 To UBound(SoundCache)
        frmEditor_Map.cmbSoundEffect.AddItem SoundCache(I)
    Next
    
    frmEditor_Map.cmbSoundEffect.ListIndex = 0
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Form_Load", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub optBlock_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    frmMain.chkGrid.Enabled = False
    frmMain.chkEyeDropper.Enabled = True
    frmEditor_Map.chkRandom.Enabled = True
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "optBlock_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub optHeal_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    cmbHeal.ListIndex = 0
    ClearAttributeFrames
    picAttributes.Visible = True
    fraHeal.Visible = True
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "optHeal_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub optLayer_Click(Index As Integer)
    Dim I As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Set which layer we're on
    CurrentLayer = Index
    If optLayer(1).Value = 0 And displayTilesets Then
        displayTilesets = False
        frmMain.chkTilesets.Value = 0
    End If

    If chkRandom = 1 Then
        EditorTileX = 1
        EditorTileY = 1
    End If
    redrawMapCache = True
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "optLayer_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub optLayers_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If OptLayers.Value Then
        fraLayers.Visible = True
        fraAttribs.Visible = False
    End If
    
    frmMain.chkEyeDropper.Enabled = True
    frmEditor_Map.chkRandom.Enabled = True
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "optLayers_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub optAttributes_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If OptAttributes.Value Then
        fraLayers.Visible = False
        fraAttribs.Visible = True
    End If
    
    frmMain.chkGrid.Enabled = True
    frmMain.chkEyeDropper.Enabled = True
    frmEditor_Map.chkRandom.Enabled = True
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "optAttribs_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub optNPCSpawn_Click()
    Dim n As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lstNpc.Clear
    
    For n = 1 To MAX_MAP_NPCS
        If Map.NPC(n) > 0 Then
            lstNpc.AddItem n & ": " & NPC(Map.NPC(n)).Name
        Else
            lstNpc.AddItem n & ": No NPC"
        End If
    Next n
    
    scrlNpcDir.Value = 0
    lstNpc.ListIndex = 0
    
    ClearAttributeFrames
    picAttributes.Visible = True
    fraNpcSpawn.Visible = True
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "optNPCSpawn_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub chkRandom_Click()
    Dim I As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    frmEditor_Map.fraRandom.Visible = Not frmEditor_Map.fraRandom.Visible
    frmEditor_Map.fraType.Visible = Not frmEditor_Map.fraType.Visible
    fraLayers.Visible = True
    fraAttribs.Visible = False
    frmEditor_Map.OptLayers.Value = True
    
    If frmEditor_Map.chkRandom = 1 Then
        EditorTileX = 1
        EditorTileY = 1
        frmEditor_Map.optLayer(MapLayer.Ground).Value = 1
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "chkRandom_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub optResource_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ClearAttributeFrames
    If Not Trim$(Resource(scrlResource.Value).Name) = vbNullString Then
        lblResource.Caption = Trim$(Resource(scrlResource.Value).Name)
    End If
    picAttributes.Visible = True
    fraResource.Visible = True
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "optResource_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub optShop_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ClearAttributeFrames
    picAttributes.Visible = True
    fraShop.Visible = True
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "optShop_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub optSlide_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    cmbSlide.ListIndex = 0
    ClearAttributeFrames
    picAttributes.Visible = True
    fraSlide.Visible = True
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "optSlide_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub optSprite_Click()
  ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
   
    ClearAttributeFrames
    picAttributes.Visible = True
    Exit Sub
   
' Error handler
ErrorHandler:
    HandleError "optSprite_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub optTrap_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    cmbTrap.ListIndex = 0
    ClearAttributeFrames
    picAttributes.Visible = True
    fraTrap.Visible = True
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "optTrap_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub MapEditorDrag(Button As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Button = vbLeftButton Then
        ' Convert the pixel number to tile number
        X = (X \ PIC_X) + 1
        Y = (Y \ PIC_Y) + 1
        
        ' Check it's not out of bounds
        If X < 0 Then X = 0
        If X > Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Width / PIC_X Then X = Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Width / PIC_X
        If Y < 0 Then Y = 0
        If Y > Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Height / PIC_Y Then Y = Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Height / PIC_Y
        
        ' Find out what to set the width + height of map editor to
        If X > EditorTileX Then ' Drag right
            EditorTileWidth = X - EditorTileX
        Else ' Drag left
            ' TO DO
        End If
        If Y > EditorTileY Then ' Drag down
            EditorTileHeight = Y - EditorTileY
        Else ' Drag up
            ' TO DO
        End If
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "MapEditorDrag", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub optWarp_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ClearAttributeFrames
    picAttributes.Visible = True
    fraMapWarp.Visible = True
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "optWarp_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub optItem_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ClearAttributeFrames
    picAttributes.Visible = True
    fraMapItem.Visible = True
    
    If Not Trim$(Item(scrlMapItem.Value).Name) = vbNullString Then
        lblMapItem.Caption = Trim$(Item(scrlMapItem.Value).Name) & " x" & scrlMapItemValue.Value
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "optItem_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdFill_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    MapEditorFillLayer
    redrawMapCache = True
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdFill_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdClear_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call MapEditorClearLayer
    redrawMapCache = True
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdClear_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdAttributeClear_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call MapEditorClearAttributes
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdAttributeClear_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picRandomTile_Click(Index As Integer)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    RandomTileSelected = Index
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picRandomTile_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlHeal_Change()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    MapEditorVitalAmount = scrlHeal.Value
    lblHeal.Caption = "Amount: " & scrlHeal.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlHeal_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlDamage_Change()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    MapEditorVitalAmount = scrlDamage.Value
    lblDamage.Caption = "Amount: " & scrlDamage.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlDamage_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlMapItem_Change()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Item(scrlMapItem.Value).stackable = 1 Then
        scrlMapItemValue.Enabled = True
    Else
        scrlMapItemValue.Value = 1
        scrlMapItemValue.Enabled = False
    End If
    
    If Not Trim$(Item(scrlMapItem.Value).Name) = vbNullString Then
        lblMapItem.Caption = Trim$(Item(scrlMapItem.Value).Name) & " x" & scrlMapItemValue.Value
    Else
        lblMapItem.Caption = "None"
        frmEditor_Map.picMapItem.Cls
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlMapItem_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlMapItem_Scroll()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    scrlMapItem_Change
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlMapItem_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlMapItemValue_Change()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblMapItem.Caption = Trim$(Item(scrlMapItem.Value).Name) & " x" & scrlMapItemValue.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlMapItemValue_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlMapItemValue_Scroll()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    scrlMapItemValue_Change
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlMapItemValue_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlMapWarp_Change()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblMapWarp.Caption = "Map: " & scrlMapWarp.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlMapWarp_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlMapWarp_Scroll()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    scrlMapWarp_Change
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlMapWarp_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlMapWarpX_Change()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblMapWarpX.Caption = "X: " & scrlMapWarpX.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlMapWarpX_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlMapWarpX_Scroll()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    scrlMapWarpX_Change
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlMapWarpX_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlMapWarpY_Change()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblMapWarpY.Caption = "Y: " & scrlMapWarpY.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlMapWarpY_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlMapWarpY_Scroll()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    scrlMapWarpY_Change
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlMapWarpY_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlNPCDir_Change()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Select Case scrlNpcDir.Value
        Case DIR_UP
            lblNpcDir = "Direction: Up"
        Case DIR_DOWN
            lblNpcDir = "Direction: Down"
        Case DIR_LEFT
            lblNpcDir = "Direction: Left"
        Case DIR_RIGHT
            lblNpcDir = "Direction: Right"
    End Select
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlNPCDir_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlNPCDir_Scroll()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    scrlNPCDir_Change
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlNPCDir_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlResource_Change()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Not Trim$(Resource(scrlResource.Value).Name) = vbNullString Then
        lblResource.Caption = Resource(scrlResource.Value).Name
    Else
        lblResource.Caption = "None"
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlResource_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlResource_Scroll()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    scrlResource_Change
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlResource_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlTileSet_Change()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    fraTileSet.Caption = "Tileset: " & scrlTileSet.Value
    
    EditorTileX = 0
    EditorTileY = 0
    EditorTileWidth = 1
    EditorTileHeight = 1
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlTileSet_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlTileSet_Scroll()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    scrlTileSet_Change
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlTileSet_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdCopy_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call CopyMemory(ByVal VarPtr(TempMap), ByVal VarPtr(Map), LenB(Map))
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdCopy_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdPaste_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Don't paste if nothing has been copied
    If TempMap.MaxX = 0 Or TempMap.MaxY = 0 Then Exit Sub
    
    Call CopyMemory(ByVal VarPtr(Map), ByVal VarPtr(TempMap), LenB(TempMap))
    InitAutotiles
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdPaste_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub
