VERSION 5.00
Begin VB.Form frmMapEditor 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Map Editor"
   ClientHeight    =   7065
   ClientLeft      =   150
   ClientTop       =   420
   ClientWidth     =   9495
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   471
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   633
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraAttribs 
      Appearance      =   0  'Flat
      Caption         =   "Attributes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6315
      Left            =   7560
      TabIndex        =   32
      Top             =   720
      Visible         =   0   'False
      Width           =   1680
      Begin VB.OptionButton optSpawnGate 
         Caption         =   "Spawn Gate"
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
         TabIndex        =   62
         Top             =   5640
         Width           =   1095
      End
      Begin VB.OptionButton optFurniture 
         Caption         =   "Furniture"
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
         TabIndex        =   61
         Top             =   4680
         Width           =   1170
      End
      Begin VB.OptionButton optHouse 
         Caption         =   "Player House"
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
         TabIndex        =   60
         Top             =   4440
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
         Left            =   120
         TabIndex        =   59
         Top             =   4200
         Width           =   1200
      End
      Begin VB.OptionButton optHouseBuy 
         Caption         =   "Buy House"
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
         TabIndex        =   58
         Top             =   3960
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
         Left            =   120
         TabIndex        =   57
         Top             =   4920
         Width           =   810
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
         TabIndex        =   56
         Top             =   840
         Width           =   1050
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
         Left            =   120
         TabIndex        =   55
         Top             =   2280
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
         Left            =   120
         TabIndex        =   54
         Top             =   3480
         Width           =   1050
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
         Left            =   120
         TabIndex        =   53
         Top             =   3720
         Width           =   735
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
         Left            =   120
         TabIndex        =   52
         Top             =   3240
         Width           =   1170
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
         Left            =   120
         TabIndex        =   51
         Top             =   3000
         Width           =   600
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
         Left            =   960
         TabIndex        =   50
         Top             =   2760
         Width           =   600
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
         TabIndex        =   49
         Top             =   1080
         Width           =   975
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
         Left            =   120
         TabIndex        =   48
         Top             =   1320
         Width           =   690
      End
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
         Left            =   960
         TabIndex        =   47
         Top             =   1320
         Width           =   690
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
         TabIndex        =   46
         Top             =   2040
         Width           =   570
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
         Left            =   960
         TabIndex        =   45
         Top             =   2040
         Width           =   675
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
         TabIndex        =   44
         Top             =   2760
         Width           =   615
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
         TabIndex        =   43
         Top             =   2520
         Width           =   1095
      End
      Begin VB.OptionButton optItem 
         Caption         =   "Item"
         Enabled         =   0   'False
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
         TabIndex        =   42
         Top             =   600
         Visible         =   0   'False
         Width           =   855
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
         TabIndex        =   41
         Top             =   360
         Width           =   735
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
         TabIndex        =   40
         Top             =   120
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optFish 
         Caption         =   "Fishing"
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
         TabIndex        =   39
         Top             =   1560
         Width           =   810
      End
      Begin VB.OptionButton optMine 
         Caption         =   "Mining"
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
         Left            =   960
         TabIndex        =   38
         Top             =   1560
         Width           =   690
      End
      Begin VB.OptionButton optLJacking 
         Caption         =   "LJ Skill"
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
         TabIndex        =   37
         Top             =   1800
         Width           =   810
      End
      Begin VB.OptionButton optRoofBlock 
         Caption         =   "Roof/Block"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   5400
         Width           =   1095
      End
      Begin VB.OptionButton optRoof 
         Caption         =   "Roof"
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
         TabIndex        =   35
         Top             =   5160
         Width           =   735
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
         Left            =   2280
         TabIndex        =   34
         Top             =   4680
         Width           =   720
      End
      Begin VB.CommandButton cmdClear2 
         Caption         =   "Clear All"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   33
         Top             =   6000
         Width           =   1155
      End
   End
   Begin VB.HScrollBar sclTileset 
      Height          =   255
      Left            =   600
      Max             =   6
      TabIndex        =   27
      Top             =   6360
      Width           =   6675
   End
   Begin VB.VScrollBar scrlPicture 
      Height          =   5385
      LargeChange     =   10
      Left            =   240
      Max             =   512
      TabIndex        =   2
      Top             =   840
      Width           =   255
   End
   Begin VB.PictureBox picBack 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   5385
      Left            =   600
      ScaleHeight     =   359
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   445
      TabIndex        =   0
      Top             =   840
      Width           =   6675
      Begin VB.PictureBox picBackSelect 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   5400
         Left            =   0
         ScaleHeight     =   360
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   445
         TabIndex        =   1
         Top             =   0
         Width           =   6675
         Begin VB.Shape shpSelected 
            BorderColor     =   &H000000FF&
            Height          =   480
            Left            =   0
            Top             =   0
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   10770
      TabIndex        =   14
      Top             =   0
      Width           =   10800
      Begin VB.CommandButton cmdDelete 
         Height          =   615
         Left            =   3840
         Picture         =   "frmMapEditor.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Fill"
         Top             =   0
         Width           =   600
      End
      Begin VB.CommandButton cmdDayNight 
         Height          =   615
         Left            =   5880
         Picture         =   "frmMapEditor.frx":063B
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Change from night to day or day to night"
         Top             =   0
         Width           =   600
      End
      Begin VB.CheckBox chkGrid 
         Height          =   615
         Left            =   5280
         Picture         =   "frmMapEditor.frx":0C62
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Turn on/off the map grid."
         Top             =   0
         Width           =   600
      End
      Begin VB.CommandButton cmdFill 
         Height          =   615
         Left            =   3240
         Picture         =   "frmMapEditor.frx":131C
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Fill"
         Top             =   0
         Width           =   600
      End
      Begin VB.CommandButton cmdEyeDropper 
         Height          =   615
         Left            =   2640
         Picture         =   "frmMapEditor.frx":1F60
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Eyedropper"
         Top             =   0
         Width           =   600
      End
      Begin VB.CommandButton cmdprops 
         Height          =   615
         Left            =   2040
         Picture         =   "frmMapEditor.frx":297C
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Map properties"
         Top             =   0
         Width           =   600
      End
      Begin VB.OptionButton optlight 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   8760
         Picture         =   "frmMapEditor.frx":35C0
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Light"
         Top             =   0
         Width           =   600
      End
      Begin VB.OptionButton optAttributes 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   8160
         Picture         =   "frmMapEditor.frx":4204
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Attributes"
         Top             =   0
         Width           =   600
      End
      Begin VB.OptionButton optTiles 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7560
         Picture         =   "frmMapEditor.frx":4E48
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Tiles"
         Top             =   0
         Value           =   -1  'True
         Width           =   600
      End
      Begin VB.CommandButton cmdSave 
         Height          =   615
         Left            =   120
         Picture         =   "frmMapEditor.frx":5A8C
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Save and Exit"
         Top             =   0
         Width           =   600
      End
      Begin VB.CommandButton cmdExit 
         Height          =   615
         Left            =   720
         Picture         =   "frmMapEditor.frx":6088
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Exit and Don't Save"
         Top             =   0
         Width           =   600
      End
      Begin VB.CommandButton cmdMinnim 
         Height          =   615
         Left            =   1320
         Picture         =   "frmMapEditor.frx":6604
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Minimize"
         Top             =   0
         Width           =   600
      End
      Begin VB.CheckBox chkScreenshot 
         Height          =   615
         Left            =   4680
         Picture         =   "frmMapEditor.frx":6A70
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Screenshot Mode"
         Top             =   0
         Width           =   600
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Tileset:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   30
         Top             =   75
         Width           =   975
      End
      Begin VB.Label lblTileset 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6720
         TabIndex        =   28
         Top             =   240
         Width           =   450
      End
   End
   Begin VB.CheckBox chkDayNight 
      Caption         =   "Check1"
      Height          =   255
      Left            =   3420
      TabIndex        =   29
      Top             =   2280
      Width           =   255
   End
   Begin VB.Frame fraLayers 
      Appearance      =   0  'Flat
      Caption         =   "Layers"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5835
      Left            =   7560
      TabIndex        =   3
      Top             =   720
      Width           =   1680
      Begin VB.OptionButton optAnim 
         Caption         =   "Animation"
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
         Left            =   285
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   960
         Width           =   1080
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H80000002&
         Caption         =   "Clear Layer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   13
         Top             =   5400
         Width           =   1155
      End
      Begin VB.OptionButton optFringe 
         Caption         =   "Fringe"
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
         Left            =   285
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1920
         Width           =   1080
      End
      Begin VB.OptionButton optGround 
         Caption         =   "Ground"
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
         Left            =   285
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Value           =   -1  'True
         Width           =   1080
      End
      Begin VB.OptionButton optMask2 
         Caption         =   "Mask 2"
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
         Left            =   285
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1320
         Width           =   1080
      End
      Begin VB.OptionButton optM2Anim 
         Caption         =   "Animation"
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
         Left            =   285
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1560
         Width           =   1080
      End
      Begin VB.OptionButton optFAnim 
         Caption         =   "Animation"
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
         Left            =   285
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2160
         Width           =   1080
      End
      Begin VB.OptionButton optFringe2 
         Caption         =   "Fringe 2"
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
         Left            =   285
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2520
         Width           =   1080
      End
      Begin VB.OptionButton optF2Anim 
         Caption         =   "Animation"
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
         Left            =   285
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2760
         Width           =   1080
      End
      Begin VB.OptionButton optMask 
         Caption         =   "Mask"
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
         Left            =   285
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   720
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmMapEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Copyright (c) 2006 Chaos Engine Source. All rights reserved.
' This code is licensed under the Chaos Engine General License.

Dim KeyShift As Boolean

Private Sub chkGrid_Click()
   If chkGrid.Value = 0 Then
        Call PutVar(App.Path & "\Main\Config\config.ini", "CONFIG", "MapGrid", 0)
    Else
        Call PutVar(App.Path & "\Main\Config\config.ini", "CONFIG", "MapGrid", 1)
    End If
End Sub

Private Sub chkScreenshot_Click()
   If chkScreenshot.Value = 0 Then
        ScreenMode = 0
    Else
        ScreenMode = 1
    End If
End Sub

Private Sub cmdDayNight_Click()
If chkDayNight.Value = 1 Then
    chkDayNight.Value = 0
ElseIf chkDayNight.Value = 0 Then
    chkDayNight.Value = 1
End If
End Sub

Private Sub cmdDelete_Click()
    Call EditorClearMap
End Sub

Private Sub cmdEyeDropper_Click()
    'If frmMapEditor.MousePointer = 2 Or frmMirage.MousePointer = 2 Then
    '    frmMapEditor.MousePointer = 1
    '    frmMirage.MousePointer = 1
    'Else
    '    frmMapEditor.MousePointer = 2
    '    frmMirage.MousePointer = 2
    'End If
   
Dim x As Integer
Dim y As Integer
Dim scripts As String
Dim I As Byte, itemp As Byte, Perm As Byte, Mapy As Byte

    scripts = "' Edit out all of the tiles that you don't want, this converts the whole map to SS."

    Perm = MsgBox("Would you like the changes to be permanent?", vbYesNo)
    If Perm = vbNo Then
        Perm = 0
        Mapy = MsgBox("Would you like to show the changes to all users on the map?", vbYesNo)
        If Mapy = vbNo Then
            Mapy = 0
        Else
            Mapy = 1
        End If
    Else
        Perm = 1
        Mapy = 0
    End If
    itemp = MsgBox("Would you like existing tiles to be erased (if a tile exists at that location)?", vbYesNo)
    If itemp = vbNo Then
        itemp = 0
    Else
        itemp = 1
    End If
   
    x = MsgBox("Ready? When you hit 'yes', everything will freeze. Please wait for the loop to finish ('no' to quit)", vbYesNo)
    If x = vbNo Then
        Exit Sub
    End If

    For x = 0 To MAX_MAPX
        For y = 0 To MAX_MAPY
            I = itemp
            With Map(GetPlayerMap(MyIndex)).Tile(x, y)
                If .Ground <> 0 Then
                    scripts = scripts & vbNewLine & "Call TileCreator(GetPlayerMap(index), " & x & ", " & y & ", " & .Ground & ", " & Chr(34) & "Ground" & Chr(34) & ", " & .GroundSet & ", " & I & ", " & Perm & ", " & Mapy & ", Index)"
                    I = 0
                End If
                If .Mask <> 0 Then
                    scripts = scripts & vbNewLine & "Call TileCreator(GetPlayerMap(index), " & x & ", " & y & ", " & .Mask & ", " & Chr(34) & "Mask" & Chr(34) & ", " & .MaskSet & ", " & I & ", " & Perm & ", " & Mapy & ", Index)"
                    I = 0
                End If
                If .Anim <> 0 Then
                    scripts = scripts & vbNewLine & "Call TileCreator(GetPlayerMap(index), " & x & ", " & y & ", " & .Anim & ", " & Chr(34) & "Anim" & Chr(34) & ", " & .AnimSet & ", " & I & ", " & Perm & ", " & Mapy & ", Index)"
                    I = 0
                End If
                If .Mask2 <> 0 Then
                    scripts = scripts & vbNewLine & "Call TileCreator(GetPlayerMap(index), " & x & ", " & y & ", " & .Mask2 & ", " & Chr(34) & "Mask2" & Chr(34) & ", " & .Mask2Set & ", " & I & ", " & Perm & ", " & Mapy & ", Index)"
                    I = 0
                End If
                If .M2Anim <> 0 Then
                    scripts = scripts & vbNewLine & "Call TileCreator(GetPlayerMap(index), " & x & ", " & y & ", " & .M2Anim & ", " & Chr(34) & "M2Anim" & Chr(34) & ", " & .M2AnimSet & ", " & I & ", " & Perm & ", " & Mapy & ", Index)"
                    I = 0
                End If
                If .Fringe <> 0 Then
                    scripts = scripts & vbNewLine & "Call TileCreator(GetPlayerMap(index), " & x & ", " & y & ", " & .Fringe & ", " & Chr(34) & "Fringe" & Chr(34) & ", " & .FringeSet & ", " & I & ", " & Perm & ", " & Mapy & ", Index)"
                    I = 0
                End If
                If .FAnim <> 0 Then
                    scripts = scripts & vbNewLine & "Call TileCreator(GetPlayerMap(index), " & x & ", " & y & ", " & .FAnim & ", " & Chr(34) & "FAnim" & Chr(34) & ", " & .FAnimSet & ", " & I & ", " & Perm & ", " & Mapy & ", Index)"
                    I = 0
                End If
                If .Fringe2 <> 0 Then
                    scripts = scripts & vbNewLine & "Call TileCreator(GetPlayerMap(index), " & x & ", " & y & ", " & .Fringe2 & ", " & Chr(34) & "Fringe2" & Chr(34) & ", " & .Fringe2Set & ", " & I & ", " & Perm & ", " & Mapy & ", Index)"
                    I = 0
                End If
                If .F2Anim <> 0 Then
                    scripts = scripts & vbNewLine & "Call TileCreator(GetPlayerMap(index), " & x & ", " & y & ", " & .F2Anim & ", " & Chr(34) & "F2Anim" & Chr(34) & ", " & .F2AnimSet & ", " & I & ", " & Perm & ", " & Mapy & ", Index)"
                    I = 0
                End If
           
                If .Type <> 0 Then
                    scripts = scripts & vbNewLine & "Call AttributeCreator(GetPlayerMap(index),  " & x & ",  " & y & "," & .Type & ", " & .Data1 & ", " & .Data2 & ", " & .Data3 & ", " & Chr(34) & Chr(34) & ", " & Chr(34) & Chr(34) & ", " & Chr(34) & Chr(34) & ", " & Perm & ", " & Mapy & ", Index)"
                End If
            End With
        Next
    Next
   
    Clipboard.Clear
    Clipboard.SetText scripts
   
    Call MsgBox("Done. The scripting material has been copied to the clipboard.")
End Sub

Private Sub cmdFill_Click()
Dim y As Long
Dim x As Long

x = MsgBox("Are you sure you want to fill the map?", vbYesNo)
If x = vbNo Then
    Exit Sub
End If

If frmMapEditor.optTiles.Value = True Then
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            With Map(GetPlayerMap(MyIndex)).Tile(x, y)
                If optGround.Value = True Then
                    .Ground = EditorTileY * TilesInSheets + EditorTileX
                    .GroundSet = EditorSet
                End If
                If optMask.Value = True Then
                    .Mask = EditorTileY * TilesInSheets + EditorTileX
                    .MaskSet = EditorSet
                End If
                If optAnim.Value = True Then
                    .Anim = EditorTileY * TilesInSheets + EditorTileX
                    .AnimSet = EditorSet
                End If
                If optMask2.Value = True Then
                    .Mask2 = EditorTileY * TilesInSheets + EditorTileX
                    .Mask2Set = EditorSet
                End If
                If optM2Anim.Value = True Then
                    .M2Anim = EditorTileY * TilesInSheets + EditorTileX
                    .M2AnimSet = EditorSet
                End If
                If optFringe.Value = True Then
                    .Fringe = EditorTileY * TilesInSheets + EditorTileX
                    .FringeSet = EditorSet
                End If
                If optFAnim.Value = True Then
                    .FAnim = EditorTileY * TilesInSheets + EditorTileX
                    .FAnimSet = EditorSet
                End If
                If optFringe2.Value = True Then
                    .Fringe2 = EditorTileY * TilesInSheets + EditorTileX
                    .Fringe2Set = EditorSet
                End If
                If optF2Anim.Value = True Then
                    .F2Anim = EditorTileY * TilesInSheets + EditorTileX
                    .F2AnimSet = EditorSet
                End If
            End With
        Next x
    Next y
ElseIf frmMapEditor.optAttributes = True Then
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            With Map(GetPlayerMap(MyIndex)).Tile(x, y)
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
                    .String1 = RoofId
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
                If optBank.Value = True Then
                    .Type = TILE_TYPE_BANK
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                    .String1 = ""
                    .String2 = ""
                    .String3 = ""
                End If
                If optSpawnGate.Value = True Then
                    .Type = TILE_TYPE_SPAWNGATE
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                    .String1 = ""
                    .String2 = ""
                    .String3 = ""
                End If
                If optFish.Value = True Then
                    .Type = TILE_TYPE_FISH
                    .Data1 = FishNumber
                    .Data2 = ToolNumber
                    .Data3 = 0
                    .String1 = FishName
                    .String2 = ToolName
                    .String3 = ""
               End If
               If optMine.Value = True Then
                    .Type = TILE_TYPE_MINE
                    .Data1 = OreNumber
                    .Data2 = ToolNumber
                    .Data3 = 0
                    .String1 = OreName
                    .String2 = ToolName
                    .String3 = ""
               End If
               If optLJacking.Value = True Then
                    .Type = TILE_TYPE_LJACKING
                    .Data1 = LogNumber
                    .Data2 = ToolNumber
                    .Data3 = 0
                    .String1 = LogName
                    .String2 = ToolName
                    .String3 = ""
               End If
            End With
        Next x
    Next y
ElseIf frmMapEditor.optlight.Value = True Then
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            Map(GetPlayerMap(MyIndex)).Tile(x, y).Light = EditorTileY * TilesInSheets + EditorTileX
        Next x
    Next y
End If
End Sub

Private Sub cmdMinnim_Click()
    frmMapEditor.WindowState = 1
End Sub

Private Sub cmdprops_Click()

    Unload frmMapProperties
    frmMapProperties.Show vbModeless, frmMirage
    'frmMapProperties.Show vbModal
    frmMapEditor.Visible = False
End Sub

Private Sub cmdExit_Click()
Dim x As Long

    x = MsgBox("Are you sure you want to discard your changes?", vbYesNo)
    If x = vbNo Then
        Exit Sub
    End If
    
    ScreenMode = 0
    Call EditorCancel
End Sub

Private Sub cmdSave_Click()
Dim x As Long

    x = MsgBox("Are you sure you want to make these changes?", vbYesNo)
    If x = vbNo Then
        Exit Sub
    End If
    
    ScreenMode = 0
    Call EditorSend
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyShift Then
        KeyShift = True
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyShift = False
End Sub

Private Sub Form_Resize()
    If frmMapEditor.WindowState = 0 Then
       ' If frmMapEditor.Width > picBack.Width + scrlPicture.Width + fraAttribs.Width Then frmMapEditor.Width = (picBack.Width + scrlPicture.Width + 8 + fraAttribs.Width) * Screen.TwipsPerPixelX
      '  picBack.Height = (frmMapEditor.Height - 800) / Screen.TwipsPerPixelX
        'scrlPicture.Height = (frmMapEditor.Height - 800) / Screen.TwipsPerPixelX
        'frmMapEditor.scrlPicture.Max = ((frmMapEditor.picBackSelect.Height - frmMapEditor.picBack.Height) / PIC_Y)
       ' If frmMapEditor.Height > (picBackSelect.Height * Screen.TwipsPerPixelX) + 800 Then frmMapEditor.Height = (picBackSelect.Height * Screen.TwipsPerPixelX) + 800
        
        WindowState = 0
    End If
End Sub

Private Sub optAttributes_Click()
    fraLayers.Visible = False
    fraAttribs.Visible = True
    shpSelected.Width = 32
    shpSelected.Height = 32
    SelectorWidth = 1
    SelectorHeight = 1
End Sub

Private Sub optlight_Click()
Dim sDc As Long

fraLayers.Visible = False
fraAttribs.Visible = False
sclTileset.Value = 6
            
sDc = DD_TileSurf(sclTileset.Value).GetDC
  With frmMapEditor.picBackSelect
      .Width = DDSD_Tile(sclTileset.Value).lWidth
      .Height = DDSD_Tile(sclTileset.Value).lHeight
      .Cls
      Call BitBlt(.hDC, 0, 0, .Width, .Height, sDc, 0, 0, SRCCOPY)
  End With
  Call DD_TileSurf(sclTileset.Value).ReleaseDC(sDc)
EditorSet = 6
            
scrlPicture.Max = ((picBackSelect.Height - picBack.Height) / PIC_Y)
'picBack.Width = picBackSelect.Width
'If frmMapEditor.Width > picBack.Width + scrlPicture.Width Then frmMapEditor.Width = (picBack.Width + scrlPicture.Width + 8) * Screen.TwipsPerPixelX
'If frmMapEditor.Height > (picBackSelect.Height * Screen.TwipsPerPixelX) + 800 Then frmMapEditor.Height = (picBackSelect.Height * Screen.TwipsPerPixelX) + 800
sclTileset.Enabled = False
End Sub

Private Sub optFish_Click()
frmFish.scrlFishNo.Max = MAX_ITEMS
frmFish.scrlToolNo.Max = MAX_ITEMS
frmFish.Show vbModal
End Sub

Private Sub optMine_Click()
frmMine.scrlOreNo.Max = MAX_ITEMS
frmMine.scrlToolNo.Max = MAX_ITEMS
frmMine.Show vbModal
End Sub

Private Sub optLJacking_Click()
frmLumberJacking.scrlLogNo.Max = MAX_ITEMS
frmLumberJacking.scrlToolNo.Max = MAX_ITEMS
frmLumberJacking.Show vbModal
End Sub

Private Sub optRoof_Click()
frmRoofTile.Show vbModal
End Sub

Private Sub optTiles_Click()
    fraLayers.Visible = True
    fraAttribs.Visible = False
    sclTileset.Enabled = True
    sclTileset.Value = 0
    TileSet = 0
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
            SelectorWidth = 1
            SelectorHeight = 1
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
        SelectorWidth = shpSelected.Width / PIC_X
        SelectorHeight = shpSelected.Height / PIC_Y
    End If
    
    If optAttributes.Value = True Then
        shpSelected.Width = 32
        shpSelected.Height = 32
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
    
    If optAttributes.Value = True Then
        shpSelected.Width = 32
        shpSelected.Height = 32
    End If
    
    EditorTileX = Int(shpSelected.Left / PIC_X)
    EditorTileY = Int(shpSelected.Top / PIC_Y)
End Sub

Private Sub sclTileset_Change()
Dim sDc As Long

    sDc = DD_TileSurf(sclTileset.Value).GetDC
    With picBackSelect
        .Width = DDSD_Tile(sclTileset.Value).lWidth
        .Height = DDSD_Tile(sclTileset.Value).lHeight
        .Cls
        Call BitBlt(.hDC, 0, 0, .Width, .Height, sDc, 0, 0, SRCCOPY)
    End With
    Call DD_TileSurf(sclTileset.Value).ReleaseDC(sDc)
    
    EditorSet = sclTileset.Value
    scrlPicture.Max = ((picBackSelect.Height - picBack.Height) / PIC_Y)
  '  frmMapEditor.picBack.Width = frmMapEditor.picBackSelect.Width
  '  If frmMapEditor.Width > picBack.Width + scrlPicture.Width Then frmMapEditor.Width = (picBack.Width + scrlPicture.Width + 8) * Screen.TwipsPerPixelX
   ' If frmMapEditor.Height > (picBackSelect.Height * Screen.TwipsPerPixelX) + 800 Then frmMapEditor.Height = (picBackSelect.Height * Screen.TwipsPerPixelX) + 800
    lblTileset = sclTileset.Value
End Sub

Private Sub scrlPicture_Change()
    Call EditorTileScroll
End Sub


Private Sub optArena_Click()
    frmArena.Show vbModal
    frmArena.scrlNum1.Max = MAX_MAPS
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

Private Sub optFurniture_Click()
    frmFurniture.scrlItem.Max = MAX_ITEMS
    frmFurniture.Show vbModal
End Sub

Private Sub optHouseBuy_Click()
    frmHouse.scrlItem.Max = MAX_ITEMS
    frmHouse.Show vbModal
End Sub


Private Sub optNPC_Click()
    frmNPCSpawn.Show vbModal
    frmNPCSpawn.scrlNum.Max = MAX_NPCS
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

Private Sub optSound_Click()
    frmSound.Show vbModal
End Sub

Private Sub optSprite_Click()
    frmSpriteChange.scrlItem.Max = MAX_ITEMS
    frmSpriteChange.Show vbModal
End Sub

Private Sub cmdClear_Click()
    Call EditorClearLayer
End Sub

Private Sub cmdClear2_Click()
    Call EditorClearAttribs
End Sub

