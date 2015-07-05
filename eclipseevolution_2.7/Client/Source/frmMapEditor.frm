VERSION 5.00
Begin VB.Form frmMapEditor 
   Caption         =   "Map Editor"
   ClientHeight    =   7440
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   496
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   712
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdED 
      Caption         =   "Eye Dropper"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   9
      Top             =   120
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
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdProp 
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
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   12
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdtype 
      Caption         =   "Light"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   9960
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdtype 
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
      Height          =   375
      Index           =   2
      Left            =   8880
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdtype 
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
      Height          =   375
      Index           =   1
      Left            =   8160
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdScreeny 
      Caption         =   "Screenshot Mode"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   8
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmddaynight 
      Caption         =   "Day/Night"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdGrid 
      Caption         =   "Map Grid "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
   Begin VB.VScrollBar scrlPicture 
      Height          =   6465
      LargeChange     =   10
      Left            =   120
      Max             =   512
      TabIndex        =   2
      Top             =   840
      Width           =   255
   End
   Begin VB.PictureBox picBack 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   6480
      Left            =   480
      ScaleHeight     =   432
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   448
      TabIndex        =   0
      Top             =   840
      Width           =   6720
      Begin VB.PictureBox picBackSelect 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   6480
         Left            =   0
         ScaleHeight     =   432
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   448
         TabIndex        =   1
         Top             =   0
         Width           =   6720
         Begin VB.Shape shpSelected 
            BorderColor     =   &H000000FF&
            Height          =   480
            Left            =   0
            Top             =   0
            Width           =   480
         End
      End
   End
   Begin VB.Frame fraAttribs 
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
      Height          =   6765
      Left            =   7440
      TabIndex        =   37
      Top             =   600
      Visible         =   0   'False
      Width           =   3105
      Begin VB.OptionButton optMinusStat 
         Caption         =   "Minus Stat"
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
         Left            =   1560
         TabIndex        =   66
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton optClick 
         Caption         =   "Click"
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
         Left            =   1560
         TabIndex        =   65
         Top             =   240
         Width           =   975
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
         TabIndex        =   64
         Top             =   6000
         Width           =   810
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
         TabIndex        =   63
         Top             =   5040
         Width           =   915
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
         TabIndex        =   62
         Top             =   1920
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
         TabIndex        =   61
         Top             =   1680
         Width           =   975
      End
      Begin VB.OptionButton optWalkThru 
         Caption         =   "Walk Through"
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
         TabIndex        =   60
         Top             =   5760
         Width           =   1335
      End
      Begin VB.OptionButton OptGHook 
         Caption         =   "Grapple Stone"
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
         TabIndex        =   59
         Top             =   5520
         Width           =   1215
      End
      Begin VB.OptionButton optGuildBlock 
         Caption         =   "Guild Block"
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
         TabIndex        =   58
         Top             =   5280
         Width           =   1215
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
         TabIndex        =   57
         Top             =   4800
         Width           =   1170
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
         Left            =   120
         TabIndex        =   56
         Top             =   4560
         Width           =   1170
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
         TabIndex        =   55
         Top             =   2400
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
         TabIndex        =   54
         Top             =   2640
         Width           =   1200
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
         Left            =   1560
         TabIndex        =   53
         Top             =   720
         Visible         =   0   'False
         Width           =   720
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
         TabIndex        =   52
         Top             =   2880
         Width           =   1155
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
         Left            =   120
         TabIndex        =   51
         Top             =   3120
         Width           =   960
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
         TabIndex        =   50
         Top             =   3360
         Width           =   1080
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
         TabIndex        =   49
         Top             =   3600
         Width           =   1200
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
         TabIndex        =   48
         Top             =   2160
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
         TabIndex        =   47
         Top             =   4320
         Width           =   1170
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
         TabIndex        =   46
         Top             =   4080
         Width           =   1170
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
         TabIndex        =   45
         Top             =   3840
         Width           =   810
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
         TabIndex        =   44
         Top             =   1440
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
         TabIndex        =   43
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
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
         TabIndex        =   42
         Top             =   480
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
         Height          =   285
         Left            =   120
         TabIndex        =   41
         Top             =   6360
         Width           =   975
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
         TabIndex        =   40
         Top             =   720
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
         TabIndex        =   39
         Top             =   960
         Width           =   1215
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
         TabIndex        =   38
         Top             =   1200
         Width           =   1215
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
      Height          =   6690
      Left            =   7440
      TabIndex        =   14
      Top             =   600
      Width           =   1680
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
         TabIndex        =   24
         Top             =   2640
         Width           =   1080
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
         TabIndex        =   23
         Top             =   2400
         Width           =   1080
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
         TabIndex        =   22
         Top             =   2040
         Width           =   1095
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
         TabIndex        =   21
         Top             =   1440
         Width           =   1245
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
         TabIndex        =   20
         Top             =   1200
         Width           =   1005
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
         TabIndex        =   19
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
         TabIndex        =   18
         Top             =   600
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
         TabIndex        =   17
         Top             =   840
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
         TabIndex        =   16
         Top             =   1800
         Width           =   855
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
         Left            =   360
         TabIndex        =   15
         Top             =   6240
         Width           =   975
      End
   End
   Begin VB.Frame frmtile 
      Caption         =   "Tile Sheet"
      Height          =   3135
      Left            =   9240
      TabIndex        =   25
      Top             =   600
      Width           =   1215
      Begin VB.OptionButton Option1 
         Caption         =   "0"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Value           =   -1  'True
         Width           =   375
      End
      Begin VB.OptionButton Option1 
         Caption         =   "1"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   35
         Top             =   480
         Width           =   375
      End
      Begin VB.OptionButton Option1 
         Caption         =   "2"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   34
         Top             =   720
         Width           =   375
      End
      Begin VB.OptionButton Option1 
         Caption         =   "3"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   33
         Top             =   960
         Width           =   375
      End
      Begin VB.OptionButton Option1 
         Caption         =   "4"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   32
         Top             =   1200
         Width           =   375
      End
      Begin VB.OptionButton Option1 
         Caption         =   "5"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   31
         Top             =   1440
         Width           =   375
      End
      Begin VB.OptionButton Option1 
         Caption         =   "6"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   30
         Top             =   1680
         Width           =   375
      End
      Begin VB.OptionButton Option1 
         Caption         =   "7"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   29
         Top             =   1920
         Width           =   375
      End
      Begin VB.OptionButton Option1 
         Caption         =   "8"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   28
         Top             =   2160
         Width           =   375
      End
      Begin VB.OptionButton Option1 
         Caption         =   "9"
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   27
         Top             =   2400
         Width           =   375
      End
      Begin VB.OptionButton Option1 
         Caption         =   "10"
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   26
         Top             =   2640
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmMapEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim KeyShift As Boolean

Private Sub cmdED_Click()
    If Me.MousePointer = 2 Or frmMirage.MousePointer = 2 Then
        Me.MousePointer = 1
        frmMirage.MousePointer = 1
    Else
        Me.MousePointer = 2
        frmMirage.MousePointer = 2
    End If
End Sub

Private Sub cmdExit_Click()
    Dim X As Long

    X = MsgBox("Are you sure you want to discard your changes?", vbYesNo)
    If X = vbNo Then
        Exit Sub
    End If

    Call EditorCancel
End Sub

Private Sub cmdFill_Click()
    Dim y As Long
    Dim X As Long

    X = MsgBox("Are you sure you want to fill the map?", vbYesNo)
    If X = vbNo Then
        Exit Sub
    End If

    If MapEditorSelectedType = 1 Then
        For y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                With Map(GetPlayerMap(MyIndex)).Tile(X, y)
                    If Me.optGround.Value Then
                        .Ground = EditorTileY * TilesInSheets + EditorTileX
                        .GroundSet = EditorSet
                    End If
                    If Me.optMask.Value Then
                        .Mask = EditorTileY * TilesInSheets + EditorTileX
                        .MaskSet = EditorSet
                    End If
                    If Me.optAnim.Value Then
                        .Anim = EditorTileY * TilesInSheets + EditorTileX
                        .AnimSet = EditorSet
                    End If
                    If Me.optMask2.Value Then
                        .Mask2 = EditorTileY * TilesInSheets + EditorTileX
                        .Mask2Set = EditorSet
                    End If
                    If Me.optM2Anim.Value Then
                        .M2Anim = EditorTileY * TilesInSheets + EditorTileX
                        .M2AnimSet = EditorSet
                    End If
                    If Me.optFringe.Value Then
                        .Fringe = EditorTileY * TilesInSheets + EditorTileX
                        .FringeSet = EditorSet
                    End If
                    If Me.optFAnim.Value Then
                        .FAnim = EditorTileY * TilesInSheets + EditorTileX
                        .FAnimSet = EditorSet
                    End If
                    If Me.optFringe2.Value Then
                        .Fringe2 = EditorTileY * TilesInSheets + EditorTileX
                        .Fringe2Set = EditorSet
                    End If
                    If Me.optF2Anim.Value Then
                        .F2Anim = EditorTileY * TilesInSheets + EditorTileX
                        .F2AnimSet = EditorSet
                    End If
                End With
            Next X
        Next y
    ElseIf MapEditorSelectedType = 2 Then
        For y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                With Map(GetPlayerMap(MyIndex)).Tile(X, y)
                    If Me.optBlocked.Value Then
                        .Type = TILE_TYPE_BLOCKED
                    End If
                    If Me.optWarp.Value Then
                        .Type = TILE_TYPE_WARP
                        .Data1 = EditorWarpMap
                        .Data2 = EditorWarpX
                        .Data3 = EditorWarpY
                        .String1 = vbNullString
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End If
                    If Me.optHeal.Value Then
                        .Type = TILE_TYPE_HEAL
                        .Data1 = 0
                        .Data2 = 0
                        .Data3 = 0
                        .String1 = vbNullString
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End If
                    If Me.optKill.Value Then
                        .Type = TILE_TYPE_KILL
                        .Data1 = 0
                        .Data2 = 0
                        .Data3 = 0
                        .String1 = vbNullString
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End If
                    If Me.optItem.Value Then
                        .Type = TILE_TYPE_ITEM
                        .Data1 = ItemEditorNum
                        .Data2 = ItemEditorValue
                        .Data3 = 0
                        .String1 = vbNullString
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End If
                    If Me.optNpcAvoid.Value Then
                        .Type = TILE_TYPE_NPCAVOID
                        .Data1 = 0
                        .Data2 = 0
                        .Data3 = 0
                        .String1 = vbNullString
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End If
                    If Me.optKey.Value Then
                        .Type = TILE_TYPE_KEY
                        .Data1 = KeyEditorNum
                        .Data2 = KeyEditorTake
                        .Data3 = 0
                        .String1 = vbNullString
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End If
                    If Me.optKeyOpen.Value Then
                        .Type = TILE_TYPE_KEYOPEN
                        .Data1 = KeyOpenEditorX
                        .Data2 = KeyOpenEditorY
                        .Data3 = 0
                        .String1 = KeyOpenEditorMsg
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End If
                    If Me.optShop.Value Then
                        .Type = TILE_TYPE_SHOP
                        .Data1 = EditorShopNum
                        .Data2 = 0
                        .Data3 = 0
                        .String1 = vbNullString
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End If
                    If Me.optCBlock.Value Then
                        .Type = TILE_TYPE_CBLOCK
                        .Data1 = EditorItemNum1
                        .Data2 = EditorItemNum2
                        .Data3 = EditorItemNum3
                        .String1 = vbNullString
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End If
                    If Me.optArena.Value Then
                        .Type = TILE_TYPE_ARENA
                        .Data1 = Arena1
                        .Data2 = Arena2
                        .Data3 = Arena3
                        .String1 = vbNullString
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End If
                    If Me.optSound.Value Then
                        .Type = TILE_TYPE_SOUND
                        .Data1 = 0
                        .Data2 = 0
                        .Data3 = 0
                        .String1 = SoundFileName
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End If
                    If Me.optSprite.Value Then
                        .Type = TILE_TYPE_SPRITE_CHANGE
                        .Data1 = SpritePic
                        .Data2 = SpriteItem
                        .Data3 = SpritePrice
                        .String1 = vbNullString
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End If
                    If Me.optSign.Value Then
                        .Type = TILE_TYPE_SIGN
                        .Data1 = 0
                        .Data2 = 0
                        .Data3 = 0
                        .String1 = SignLine1
                        .String2 = SignLine2
                        .String3 = SignLine3
                    End If
                    If Me.optDoor.Value Then
                        .Type = TILE_TYPE_DOOR
                        .Data1 = 0
                        .Data2 = 0
                        .Data3 = 0
                        .String1 = vbNullString
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End If
                    If Me.optNotice.Value Then
                        .Type = TILE_TYPE_NOTICE
                        .Data1 = 0
                        .Data2 = 0
                        .Data3 = 0
                        .String1 = NoticeTitle
                        .String2 = NoticeText
                        .String3 = NoticeSound
                    End If
                    If Me.optChest.Value Then
                        .Type = TILE_TYPE_CHEST
                        .Data1 = 0
                        .Data2 = 0
                        .Data3 = 0
                        .String1 = vbNullString
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End If
                    If Me.optClassChange.Value Then
                        .Type = TILE_TYPE_CLASS_CHANGE
                        .Data1 = ClassChange
                        .Data2 = ClassChangeReq
                        .Data3 = 0
                        .String1 = vbNullString
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End If
                    If Me.optScripted.Value Then
                        .Type = TILE_TYPE_SCRIPTED
                        .Data1 = ScriptNum
                        .Data2 = 0
                        .Data3 = 0
                        .String1 = vbNullString
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End If
                    If Me.optGuildBlock.Value Then
                        .Type = TILE_TYPE_GUILDBLOCK
                        .Data1 = 0
                        .Data2 = 0
                        .Data3 = 0
                        .String1 = GuildBlock
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End If
                    If Me.optBank.Value Then
                        .Type = TILE_TYPE_BANK
                        .Data1 = 0
                        .Data2 = 0
                        .Data3 = 0
                        .String1 = vbNullString
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End If
                    If Me.OptGHook.Value Then
                        .Type = TILE_TYPE_HOOKSHOT
                        .Data1 = 0
                        .Data2 = 0
                        .Data3 = 0
                        .String1 = vbNullString
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End If
                End With
            Next X
        Next y
    ElseIf MapEditorSelectedType = 3 Then
        For y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                Map(GetPlayerMap(MyIndex)).Tile(X, y).light = EditorTileY * TilesInSheets + EditorTileX
            Next X
        Next y
    End If
End Sub

Private Sub cmdGrid_Click()
    If GridMode = 0 Then
        GridMode = 1
    Else
        GridMode = 0
    End If
End Sub

Private Sub cmdScreeny_Click()
    If ScreenMode = 0 Then
        ScreenMode = 1
    Else
        ScreenMode = 0
    End If
End Sub

Private Sub cmddaynight_Click()
    If NightMode = 0 Then
        NightMode = 1
    Else
        NightMode = 0
    End If
End Sub

Private Sub cmdProp_Click()
    frmMapProperties.Show vbModal
End Sub

Private Sub cmdSave_Click()
    Dim X As Long

    X = MsgBox("Are you sure you want to make these changes?", vbYesNo)
    If X = vbNo Then
        Exit Sub
    End If

    Call EditorSend
End Sub

Private Sub cmdtype_Click(Index As Integer)
    If Index = 1 Then
        MapEditorSelectedType = 1

        Me.fraAttribs.Visible = False
        Me.fraLayers.Visible = True
        Me.frmtile.Visible = True
    ElseIf Index = 2 Then
        MapEditorSelectedType = 2

        Me.shpSelected.Width = 32
        Me.shpSelected.Height = 32

        Me.fraLayers.Visible = False
        Me.frmtile.Visible = False
        Me.fraAttribs.Visible = True
    Else
        MapEditorSelectedType = 3

        Me.fraAttribs.Visible = False
        Me.fraLayers.Visible = False
        Me.frmtile.Visible = False
        Me.Option1(10).Value = True

        Me.picBackSelect.Picture = LoadPicture(App.Path & "\GFX\Tiles10.bmp")

        EditorSet = 10

        scrlPicture.Max = Int((picBackSelect.Height - picBack.Height) / PIC_Y)
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyShift Then
        KeyShift = True
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyShift Then
        KeyShift = False
    End If
End Sub

Private Sub optClick_Click()
    frmClick.Show vbModal
End Sub

Private Sub optGuildBlock_Click()
    frmGuildBlock.Show vbModal
    frmGuildBlock.txtGuild.Text = vbNullString
End Sub

Private Sub Option1_Click(Index As Integer)
    Option1(Index).Value = True

    Me.picBackSelect.Picture = LoadPicture(App.Path & "\GFX\Tiles" & Index & ".bmp")

    EditorSet = Index

    scrlPicture.Max = Int((picBackSelect.Height - picBack.Height) / PIC_Y)
End Sub

Private Sub optMinusStat_Click()
    frmMinusStat.Show
    frmMinusStat.scrlNum1.Value = MinusHp
    frmMinusStat.lblNum1.Caption = MinusHp
    frmMinusStat.scrlNum2.Value = MinusMp
    frmMinusStat.lblNum2.Caption = MinusMp
    frmMinusStat.scrlNum3.Value = MinusSp
    frmMinusStat.lblNum3.Caption = MinusSp
    frmMinusStat.Text1.Text = MessageMinus
End Sub

Private Sub optRoof_Click()
    frmRoofTile.Show vbModal
End Sub

Private Sub picBackSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyShift Then
        KeyShift = True
    End If
End Sub

Private Sub picBackSelect_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyShift Then
        KeyShift = False
    End If
End Sub

Private Sub picBackSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = 1 Then
        If Not KeyShift Then
            Call EditorChooseTile(Button, Shift, X, y)

            shpSelected.Width = 32
            shpSelected.Height = 32
        Else
            EditorTileX = Int(X / PIC_X)
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

    If MapEditorSelectedType = 2 Then
        shpSelected.Width = 32
        shpSelected.Height = 32
    End If

    EditorTileX = Int((shpSelected.Left + PIC_X) / PIC_X)
    EditorTileY = Int((shpSelected.Top + PIC_Y) / PIC_Y)
End Sub

Private Sub picBackSelect_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = 1 Then
        If Not KeyShift Then
            Call EditorChooseTile(Button, Shift, X, y)

            shpSelected.Width = 32
            shpSelected.Height = 32
        Else
            EditorTileX = Int(X / PIC_X)
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

    If MapEditorSelectedType = 2 Then
        shpSelected.Width = 32
        shpSelected.Height = 32
    End If

    EditorTileX = Int(shpSelected.Left / PIC_X)
    EditorTileY = Int(shpSelected.Top / PIC_Y)
End Sub

Private Sub scrlPicture_Change()
    Call EditorTileScroll
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
    If SpriteSize = 1 Then
        frmSpriteChange.picSprite.Height = 960
    End If

    frmSpriteChange.scrlItem.Max = MAX_ITEMS
    frmSpriteChange.Show vbModal
End Sub

Private Sub cmdClear_Click()
    Call EditorClearLayer
End Sub

Private Sub cmdClear2_Click()
    Call EditorClearAttribs
End Sub

Private Sub optHouse_Click()
    frmHouse.scrlItem.Max = MAX_ITEMS
    frmHouse.Show vbModal
End Sub
