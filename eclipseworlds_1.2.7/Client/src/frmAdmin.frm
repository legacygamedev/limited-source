VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAdmin 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Admin Panel"
   ClientHeight    =   8640
   ClientLeft      =   19095
   ClientTop       =   1830
   ClientWidth     =   3045
   Icon            =   "frmAdmin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   576
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   203
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picPanel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
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
      ForeColor       =   &H80000008&
      Height          =   8655
      Left            =   0
      ScaleHeight     =   577
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   205
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   3075
      Begin VB.CommandButton cmdLoadTextures 
         Caption         =   "Load Textures"
         Height          =   255
         Left            =   60
         TabIndex        =   71
         Top             =   8220
         Width           =   1365
      End
      Begin VB.CheckBox chkEditor 
         Alignment       =   1  'Right Justify
         Caption         =   "Quests"
         Enabled         =   0   'False
         ForeColor       =   &H000000C0&
         Height          =   270
         Index           =   13
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   6720
         Width           =   1050
      End
      Begin VB.PictureBox picEye 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   13
         Left            =   2625
         ScaleHeight     =   210
         ScaleWidth      =   210
         TabIndex        =   69
         Top             =   6720
         Width           =   240
      End
      Begin VB.PictureBox picSizer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2565
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   68
         Top             =   3105
         Width           =   300
      End
      Begin VB.PictureBox picSpawner 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   1155
         ScaleHeight     =   210
         ScaleWidth      =   210
         TabIndex        =   67
         Top             =   4425
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.CheckBox chkEditor 
         Alignment       =   1  'Right Justify
         Caption         =   "Events"
         Enabled         =   0   'False
         ForeColor       =   &H000000C0&
         Height          =   270
         Index           =   12
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   6960
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.PictureBox picEye 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   12
         Left            =   2625
         ScaleHeight     =   210
         ScaleWidth      =   210
         TabIndex        =   65
         Top             =   6960
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox picEye 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   11
         Left            =   2625
         ScaleHeight     =   210
         ScaleWidth      =   210
         TabIndex        =   64
         Top             =   4590
         Width           =   240
      End
      Begin VB.PictureBox picEye 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   10
         Left            =   2625
         ScaleHeight     =   210
         ScaleWidth      =   210
         TabIndex        =   63
         Top             =   6465
         Width           =   240
      End
      Begin VB.PictureBox picEye 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   9
         Left            =   2625
         ScaleHeight     =   210
         ScaleWidth      =   210
         TabIndex        =   62
         Top             =   6195
         Width           =   240
      End
      Begin VB.PictureBox picEye 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   8
         Left            =   2625
         ScaleHeight     =   210
         ScaleWidth      =   210
         TabIndex        =   61
         Top             =   5925
         Width           =   240
      End
      Begin VB.PictureBox picEye 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   7
         Left            =   2625
         ScaleHeight     =   210
         ScaleWidth      =   210
         TabIndex        =   60
         Top             =   5670
         Width           =   240
      End
      Begin VB.PictureBox picEye 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   6
         Left            =   2625
         ScaleHeight     =   210
         ScaleWidth      =   210
         TabIndex        =   59
         Top             =   5400
         Width           =   240
      End
      Begin VB.PictureBox picEye 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   5
         Left            =   2625
         ScaleHeight     =   210
         ScaleWidth      =   210
         TabIndex        =   58
         Top             =   5130
         Width           =   240
      End
      Begin VB.PictureBox picEye 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   2625
         ScaleHeight     =   210
         ScaleWidth      =   210
         TabIndex        =   57
         Top             =   4860
         Width           =   240
      End
      Begin VB.PictureBox picEye 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   3
         Left            =   2625
         ScaleHeight     =   210
         ScaleWidth      =   210
         TabIndex        =   56
         Top             =   4335
         Width           =   240
      End
      Begin VB.PictureBox picEye 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   2625
         ScaleHeight     =   210
         ScaleWidth      =   210
         TabIndex        =   55
         Top             =   4065
         Width           =   240
      End
      Begin VB.PictureBox picEye 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   2625
         ScaleHeight     =   210
         ScaleWidth      =   210
         TabIndex        =   54
         Top             =   3795
         Width           =   240
      End
      Begin VB.PictureBox picEye 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   2625
         ScaleHeight     =   210
         ScaleWidth      =   210
         TabIndex        =   53
         Top             =   3525
         Width           =   240
      End
      Begin VB.CheckBox chkEditor 
         Alignment       =   1  'Right Justify
         Caption         =   "Title"
         ForeColor       =   &H000000C0&
         Height          =   270
         Index           =   11
         Left            =   1545
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   4590
         Width           =   1065
      End
      Begin VB.CheckBox chkEditor 
         Alignment       =   1  'Right Justify
         Caption         =   "Spell"
         ForeColor       =   &H000000C0&
         Height          =   270
         Index           =   10
         Left            =   1545
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   6465
         Width           =   1065
      End
      Begin VB.CheckBox chkEditor 
         Alignment       =   1  'Right Justify
         Caption         =   "Shop"
         ForeColor       =   &H000000C0&
         Height          =   270
         Index           =   9
         Left            =   1545
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   6195
         Width           =   1065
      End
      Begin VB.CheckBox chkEditor 
         Alignment       =   1  'Right Justify
         Caption         =   "Resource"
         ForeColor       =   &H000000C0&
         Height          =   270
         Index           =   8
         Left            =   1545
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   5925
         Width           =   1065
      End
      Begin VB.CheckBox chkEditor 
         Alignment       =   1  'Right Justify
         Caption         =   "NPC"
         ForeColor       =   &H000000C0&
         Height          =   270
         Index           =   7
         Left            =   1545
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   5655
         Width           =   1065
      End
      Begin VB.CheckBox chkEditor 
         Alignment       =   1  'Right Justify
         Caption         =   "Moral"
         ForeColor       =   &H000000C0&
         Height          =   270
         Index           =   6
         Left            =   1545
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   5385
         Width           =   1065
      End
      Begin VB.CheckBox chkEditor 
         Alignment       =   1  'Right Justify
         Caption         =   "Map"
         ForeColor       =   &H000000C0&
         Height          =   270
         Index           =   5
         Left            =   1545
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   5115
         Width           =   1065
      End
      Begin VB.CheckBox chkEditor 
         Alignment       =   1  'Right Justify
         Caption         =   "Item"
         ForeColor       =   &H000000C0&
         Height          =   270
         Index           =   4
         Left            =   1545
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   4845
         Width           =   1065
      End
      Begin VB.CheckBox chkEditor 
         Alignment       =   1  'Right Justify
         Caption         =   "Emoticon"
         ForeColor       =   &H000000C0&
         Height          =   270
         Index           =   3
         Left            =   1545
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   4320
         Width           =   1065
      End
      Begin VB.CheckBox chkEditor 
         Alignment       =   1  'Right Justify
         Caption         =   "Class"
         ForeColor       =   &H000000C0&
         Height          =   270
         Index           =   2
         Left            =   1545
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   4050
         Width           =   1065
      End
      Begin VB.CheckBox chkEditor 
         Alignment       =   1  'Right Justify
         Caption         =   "Ban"
         ForeColor       =   &H000000C0&
         Height          =   270
         Index           =   1
         Left            =   1545
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   3780
         Width           =   1065
      End
      Begin VB.CheckBox chkEditor 
         Alignment       =   1  'Right Justify
         Caption         =   "Animation"
         ForeColor       =   &H000000C0&
         Height          =   270
         Index           =   0
         Left            =   1545
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   3510
         Width           =   1065
      End
      Begin VB.OptionButton optCat 
         Height          =   420
         Index           =   9
         Left            =   90
         MaskColor       =   &H80000001&
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Change sprites via dbl click."
         Top             =   6090
         Width           =   420
      End
      Begin VB.PictureBox picRecentItem 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   195
         ScaleHeight     =   450
         ScaleWidth      =   450
         TabIndex        =   39
         Top             =   6870
         Width           =   480
      End
      Begin VB.OptionButton optCat 
         Height          =   420
         Index           =   8
         Left            =   990
         MaskColor       =   &H80000001&
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Revive Stones"
         Top             =   5640
         Width           =   420
      End
      Begin VB.OptionButton optCat 
         Height          =   420
         Index           =   7
         Left            =   540
         MaskColor       =   &H80000001&
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Reset Scrolls"
         Top             =   5640
         Width           =   420
      End
      Begin VB.OptionButton optCat 
         Height          =   420
         Index           =   6
         Left            =   90
         MaskColor       =   &H80000001&
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Teleport scrolls"
         Top             =   5640
         Width           =   420
      End
      Begin VB.OptionButton optCat 
         Height          =   420
         Index           =   5
         Left            =   990
         MaskColor       =   &H80000001&
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Spells, scrolls, magic."
         Top             =   5205
         Width           =   420
      End
      Begin VB.OptionButton optCat 
         Height          =   420
         Index           =   4
         Left            =   540
         MaskColor       =   &H80000001&
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Player's titles"
         Top             =   5205
         Width           =   420
      End
      Begin VB.OptionButton optCat 
         Height          =   420
         Index           =   3
         Left            =   90
         MaskColor       =   &H80000001&
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Potions, elixirs, food."
         Top             =   5205
         Width           =   420
      End
      Begin VB.OptionButton optCat 
         Height          =   420
         Index           =   2
         Left            =   990
         MaskColor       =   &H80000001&
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Things you can wear"
         Top             =   4755
         Width           =   420
      End
      Begin VB.OptionButton optCat 
         Height          =   420
         Index           =   1
         Left            =   540
         MaskColor       =   &H80000001&
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Things without a type"
         Top             =   4755
         Width           =   420
      End
      Begin VB.OptionButton optCat 
         Height          =   420
         Index           =   0
         Left            =   90
         MaskColor       =   &H80000001&
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Recently spawned items"
         Top             =   4755
         Width           =   420
      End
      Begin MSComCtl2.UpDown rcSwitcher 
         Height          =   255
         Left            =   765
         TabIndex        =   27
         Top             =   6990
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   450
         _Version        =   393216
         Max             =   20
         Orientation     =   1
         Enabled         =   0   'False
      End
      Begin VB.CommandButton cmdSpawnRecent 
         Caption         =   "Spawn Recent"
         Enabled         =   0   'False
         Height          =   255
         Left            =   60
         TabIndex        =   26
         Top             =   7935
         Width           =   1365
      End
      Begin VB.TextBox txtRecentAmount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00800080&
         Height          =   285
         Left            =   60
         TabIndex        =   25
         Text            =   "Recent Amount"
         Top             =   7620
         Width           =   1365
      End
      Begin VB.PictureBox picSprite 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   720
         Left            =   1830
         ScaleHeight     =   46
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   22
         Top             =   1995
         Width           =   510
      End
      Begin VB.TextBox txtSprite 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1770
         TabIndex        =   21
         Text            =   "0"
         Top             =   2760
         Width           =   600
      End
      Begin VB.ComboBox cmbAccess 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   315
         ItemData        =   "frmAdmin.frx":038A
         Left            =   585
         List            =   "frmAdmin.frx":039D
         TabIndex        =   20
         Text            =   "Player's Access"
         Top             =   780
         Width           =   1695
      End
      Begin VB.ComboBox cmbPlayersOnline 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         ForeColor       =   &H80000002&
         Height          =   315
         ItemData        =   "frmAdmin.frx":03CE
         Left            =   240
         List            =   "frmAdmin.frx":03D0
         TabIndex        =   19
         Text            =   "Choose Player"
         Top             =   390
         Width           =   2055
      End
      Begin VB.PictureBox picRefresh 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2370
         ScaleHeight     =   285
         ScaleWidth      =   345
         TabIndex        =   18
         Top             =   390
         Width           =   375
      End
      Begin VB.CommandButton cmdCharEditor 
         Caption         =   "Character Editor"
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
         Left            =   1170
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1710
         Width           =   1725
      End
      Begin VB.CommandButton cmdAMute 
         Caption         =   "Mute"
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
         Left            =   120
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1440
         Width           =   915
      End
      Begin VB.CommandButton cmdLevelUp 
         Caption         =   "Level Up"
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
         Left            =   120
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1980
         Width           =   915
      End
      Begin VB.CommandButton cmdARespawn 
         Caption         =   "Respawn"
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
         Left            =   165
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   3750
         Width           =   1155
      End
      Begin VB.CommandButton cmdAMapReport 
         Caption         =   "Map Report"
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
         Left            =   180
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   3240
         Width           =   1140
      End
      Begin VB.CommandButton cmdALoc 
         Caption         =   "Location"
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
         Left            =   165
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   3480
         Width           =   1155
      End
      Begin VB.CommandButton cmdAWarp 
         Caption         =   "Warp To"
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
         Left            =   195
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2985
         Width           =   1125
      End
      Begin VB.TextBox txtAMap 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   825
         TabIndex        =   0
         Top             =   2640
         Width           =   465
      End
      Begin VB.CommandButton cmdAWarpMeTo 
         Caption         =   "Admin To Player"
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
         Left            =   1170
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1725
      End
      Begin VB.CommandButton cmdAWarpToMe 
         Caption         =   "Summon Player"
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
         Left            =   1170
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1170
         Width           =   1725
      End
      Begin VB.CommandButton cmdABan 
         Caption         =   "Ban"
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
         Left            =   120
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1170
         Width           =   915
      End
      Begin VB.CommandButton cmdAKick 
         Caption         =   "Kick"
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
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1710
         Width           =   915
      End
      Begin MSComCtl2.UpDown upSprite 
         Height          =   555
         Left            =   2430
         TabIndex        =   23
         Top             =   2070
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   979
         _Version        =   393216
         BuddyControl    =   "txtSprite"
         BuddyDispid     =   196620
         OrigLeft        =   3990
         OrigTop         =   1770
         OrigRight       =   4245
         OrigBottom      =   2265
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00800080&
         X1              =   3
         X2              =   97
         Y1              =   438
         Y2              =   438
      End
      Begin VB.Label lblRecent 
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Recent"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   435
         TabIndex        =   38
         Top             =   6645
         Width           =   555
      End
      Begin VB.Label lblItemName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Item Name: None"
         Enabled         =   0   'False
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   -45
         TabIndex        =   28
         Top             =   7380
         Width           =   1470
      End
      Begin VB.Label lblCat 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Categories"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   4440
         Width           =   1035
      End
      Begin VB.Label lblMap 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Map"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   270
         TabIndex        =   17
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label lblEditors 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Editors"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   1680
         TabIndex        =   16
         Top             =   3180
         Width           =   765
      End
      Begin VB.Label lblSpawning 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Spawning"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   285
         Left            =   150
         TabIndex        =   15
         Top             =   4080
         Width           =   1140
      End
      Begin VB.Label lblPlayers 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Players"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   285
         TabIndex        =   14
         Top             =   0
         Width           =   2505
      End
      Begin VB.Line Line3 
         BorderWidth     =   3
         X1              =   16
         X2              =   188
         Y1              =   18
         Y2              =   18
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00800080&
         BorderWidth     =   3
         X1              =   8
         X2              =   89
         Y1              =   290
         Y2              =   290
      End
      Begin VB.Line lineEditors 
         BorderColor     =   &H000000C0&
         BorderWidth     =   3
         X1              =   103
         X2              =   187
         Y1              =   229
         Y2              =   229
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Map #:"
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
         Left            =   195
         TabIndex        =   12
         Top             =   2685
         Width           =   735
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   10
         X2              =   90
         Y1              =   171
         Y2              =   171
      End
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim refreshDown As Boolean
Dim autoAccess As Boolean, autoSprite As Boolean
Dim currentSprite As Long
Private currentRecentAm As Long
Private catSub As Boolean
Public lastIndex As Integer
Public currentCategory As String
Public ignoreChange As Boolean
Public reverse As Boolean

Public Sub ShowEyeFor(Editor As Byte)
    picEye(Editor).Visible = True
End Sub

Public Sub chkEditor_Click(Index As Integer)

    ' If debug mode, handle error then exit out
    
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    If ignoreChange Then
        ignoreChange = False
        Exit Sub
    End If
    
    Select Case Index
    
        Case 0 ' Animation
            If chkEditor(Index).Value = 1 Then
                If GetPlayerAccess(MyIndex) < STAFF_DEVELOPER Then
                    AddText "You have insufficent access to do this!", BrightRed
                    ignoreChange = True
                    chkEditor(Index).Value = 0
                    Exit Sub
                End If
            
                SendRequestEditAnimation
                chkEditor(Index).FontBold = True
            Else
                chkEditor(Index).FontBold = False
                frmEditor_Animation.Visible = False
                BringWindowToTop (frmAdmin.hWnd)
            End If
        Case 1 ' Ban
            If chkEditor(Index).Value = 1 Then
                If GetPlayerAccess(MyIndex) < STAFF_ADMIN Then
                    AddText "You have insufficent access to do this!", BrightRed
                    ignoreChange = True
                    chkEditor(Index).Value = 0
                    Exit Sub
                End If
                
                SendRequestEditBan
                chkEditor(Index).FontBold = True
            Else
                chkEditor(Index).FontBold = False
                frmEditor_Ban.Visible = False
                    BringWindowToTop (frmAdmin.hWnd)
            End If
        Case 2 'Class
            If chkEditor(Index).Value = 1 Then
                If GetPlayerAccess(MyIndex) < STAFF_DEVELOPER Then
                    AddText "You have insufficent access to do this!", BrightRed
                    ignoreChange = True
                    chkEditor(Index).Value = 0
                    Exit Sub
                End If
                
                SendRequestEditClass
                chkEditor(Index).FontBold = True
            Else
                chkEditor(Index).FontBold = False
                frmEditor_Class.Visible = False
                    BringWindowToTop (frmAdmin.hWnd)
            End If
        Case 3 ' Emoticons
            If chkEditor(Index).Value = 1 Then
                If GetPlayerAccess(MyIndex) < STAFF_DEVELOPER Then
                    AddText "You have insufficent access to do this!", BrightRed
                    ignoreChange = True
                    chkEditor(Index).Value = 0
                    Exit Sub
                End If
    
                SendRequestEditEmoticon
                chkEditor(Index).FontBold = True
            Else
                chkEditor(Index).FontBold = False
                frmEditor_Emoticon.Visible = False
                BringWindowToTop (frmAdmin.hWnd)
            End If
        Case 4 ' Items
            If chkEditor(Index).Value = 1 Then
                If GetPlayerAccess(MyIndex) < STAFF_DEVELOPER Then
                    AddText "You have insufficent access to do this!", BrightRed
                    ignoreChange = True
                    chkEditor(Index).Value = 0
                    Exit Sub
                End If
            
                SendRequestEditItem
                chkEditor(Index).FontBold = True
            Else
                chkEditor(Index).FontBold = False
                frmEditor_Item.Visible = False
                BringWindowToTop (frmAdmin.hWnd)
            End If
        Case 5 ' Map
                If chkEditor(Index).Value = 1 Then
                    If GetPlayerAccess(MyIndex) < STAFF_MAPPER Then
                        AddText "You have insufficent access to do this!", BrightRed
                        ignoreChange = True
                        chkEditor(Index).Value = 0
                        Exit Sub
                    End If
                    
                    SendRequestEditMap
                    chkEditor(Index).FontBold = True
                Else
                    If FormVisible("frmEditor_Map") Then
                        ignoreChange = True
                        chkEditor(Index).Value = 1
                        LeaveMapEditorMode True
                    End If
                End If
        Case 6 ' Moral
            If chkEditor(Index).Value = 1 Then
                If GetPlayerAccess(MyIndex) < STAFF_DEVELOPER Then
                    AddText "You have insufficent access to do this!", BrightRed
                    ignoreChange = True
                    chkEditor(Index).Value = 0
                    Exit Sub
                End If
                
                SendRequestEditMoral
                chkEditor(Index).FontBold = True
            Else
                chkEditor(Index).FontBold = False
                frmEditor_Moral.Visible = False
                    BringWindowToTop (frmAdmin.hWnd)
            End If
        Case 7 'NPC
            If chkEditor(Index).Value = 1 Then
                If GetPlayerAccess(MyIndex) < STAFF_DEVELOPER Then
                    AddText "You have insufficent access to do this!", BrightRed
                    ignoreChange = True
                    chkEditor(Index).Value = 0
                    Exit Sub
                End If
            
                SendRequestEditNPC
                chkEditor(Index).FontBold = True
            Else
                chkEditor(Index).FontBold = False
                frmEditor_NPC.Visible = False
                    BringWindowToTop (frmAdmin.hWnd)
            End If
         Case 8 ' Resource

            If chkEditor(Index).Value = 1 Then
                If GetPlayerAccess(MyIndex) < STAFF_DEVELOPER Then
                    AddText "You have insufficent access to do this!", BrightRed
                    ignoreChange = True
                    chkEditor(Index).Value = 0
                    Exit Sub
                End If
            
                SendRequestEditResource
                chkEditor(Index).FontBold = True
            Else
                chkEditor(Index).FontBold = False
                frmEditor_Resource.Visible = False
                    BringWindowToTop (frmAdmin.hWnd)
            End If
        Case 9 ' Shop
            If chkEditor(Index).Value = 1 Then
                If GetPlayerAccess(MyIndex) < STAFF_DEVELOPER Then
                    AddText "You have insufficent access to do this!", BrightRed
                    ignoreChange = True
                    chkEditor(Index).Value = 0
                    Exit Sub
                End If
            
                SendRequestEditShop
                chkEditor(Index).FontBold = True
            Else
                chkEditor(Index).FontBold = False
                frmEditor_Shop.Visible = False
                    BringWindowToTop (frmAdmin.hWnd)
            End If
        Case 10 ' Spell
            If chkEditor(Index).Value = 1 Then
                If GetPlayerAccess(MyIndex) < STAFF_DEVELOPER Then
                    AddText "You have insufficent access to do this!", BrightRed
                    ignoreChange = True
                    chkEditor(Index).Value = 0
                    Exit Sub
                End If
            
                SendRequestEditSpell
                chkEditor(Index).FontBold = True
            Else
                chkEditor(Index).FontBold = False
                frmEditor_Spell.Visible = False
                    BringWindowToTop (frmAdmin.hWnd)
            End If
        Case 11 ' Title
            If chkEditor(Index).Value = 1 Then
                If GetPlayerAccess(MyIndex) < STAFF_DEVELOPER Then
                    AddText "You have insufficent access to do this!", BrightRed
                    ignoreChange = True
                    chkEditor(Index).Value = 0
                    Exit Sub
                End If
                
                SendRequestEditTitle
                chkEditor(Index).FontBold = True
            Else
                chkEditor(Index).FontBold = False
                frmEditor_Title.Visible = False
                BringWindowToTop (frmAdmin.hWnd)
            End If
        Case 12 ' Events
            If chkEditor(Index).Value = 1 Then
                If GetPlayerAccess(MyIndex) < STAFF_MAPPER Then
                    AddText "You have insufficent access to do this!", BrightRed
                    ignoreChange = True
                    chkEditor(Index).Value = 0
                End If
                    
                SendRequestEditEvent
                chkEditor(Index).FontBold = True
            Else
                chkEditor(Index).FontBold = False
                frmEditor_Events.Visible = False
                BringWindowToTop (frmAdmin.hWnd)
            End If
        Case 13 ' Quests
            If chkEditor(Index).Value = 1 Then
                If GetPlayerAccess(MyIndex) < STAFF_MAPPER Then
                    AddText "You have insufficent access to do this!", BrightRed
                    ignoreChange = True
                    chkEditor(Index).Value = 0
                End If
                    
                SendRequestEditQuests
                chkEditor(Index).FontBold = True
            Else
                chkEditor(Index).FontBold = False
                frmEditor_Quest.Visible = False
                BringWindowToTop (frmAdmin.hWnd)
            End If

    End Select
    
    picEye(Index).Visible = False
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "chkEditor_Click", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmbAccess_Click()
    If autoAccess Then
        autoAccess = False
    Else
        cmbAccess.Enabled = False
        cmbPlayersOnline.Enabled = False
        SendSetAccess cmbPlayersOnline.text, cmbAccess.ListIndex
    End If
End Sub

Public Sub VerifyAccess(PlayerName As String, Success As Byte, Message As String, CurrentAccess As Byte)
    Dim I As Long
    If PlayerName = cmbPlayersOnline.text Then
        If Success = 0 Then
            For I = 0 To UBound(g_playersOnline)
                If InStr(1, g_playersOnline(I), PlayerName) Then
                    Mid$(g_playersOnline(I), InStr(1, g_playersOnline(I), ":"), 2) = ":" & CurrentAccess
                    setAdminAccessLevel
                End If
            Next I
        ElseIf Success = 1 Then
            Mid$(g_playersOnline(I), InStr(1, g_playersOnline(I), ":"), 2) = ":" & CurrentAccess
            setAdminAccessLevel
        End If
    End If
    cmbPlayersOnline.Enabled = True
End Sub

Private Sub cmbPlayersOnline_Click()
    Dim I As Long, Length As Long
    
    Length = UBound(ignoreIndexes)
    For I = 0 To Length
        If cmbPlayersOnline.ListIndex = ignoreIndexes(I) Then
            cmbPlayersOnline.ListIndex = ignoreIndexes(I) + 1
            cmbPlayersOnline.text = cmbPlayersOnline.List(cmbPlayersOnline.ListIndex)
            Exit Sub
        End If
    Next
    autoAccess = True
    autoSprite = True
    For I = 0 To UBound(g_playersOnline)
            If InStr(1, g_playersOnline(I), cmbPlayersOnline.text) Then
                txtSprite.text = Split(g_playersOnline(I), ":")(2)
            End If
    Next I
    If Player(MyIndex).Access < 4 Then
        txtSprite.Enabled = False
        upSprite.Enabled = False
    Else
        txtSprite.Enabled = True
        upSprite.Enabled = True
    End If
    setAdminAccessLevel

    
End Sub

Private Sub setAdminAccessLevel()
    Dim accessLvl As String, tempTxt As String, I As Long
    
    ' Set Access Level
    For I = 0 To UBound(g_playersOnline)
        If InStr(1, g_playersOnline(I), cmbPlayersOnline.List(cmbPlayersOnline.ListIndex)) Then
            accessLvl = Split(g_playersOnline(I), ":")(1)
            txtSprite.text = Split(g_playersOnline(I), ":")(2)
            
            If accessLvl = "5" Then
                accessLvl = "4"
                tempTxt = "Owner"

            Else
                tempTxt = cmbAccess.List(CLng(accessLvl))

            End If
            
            If Player(MyIndex).Access > CLng(accessLvl) And Player(MyIndex).Access >= 4 And Trim$(Player(MyIndex).Name) <> cmbPlayersOnline.text Then
                cmbAccess.Enabled = True
            Else
                cmbAccess.Enabled = False
            End If
            If Player(MyIndex).Access < 4 Then
                txtSprite.Enabled = False
                upSprite.Enabled = False
            Else
                txtSprite.Enabled = True
                upSprite.Enabled = True
            End If
            cmbAccess.ListIndex = accessLvl
            cmbAccess.text = tempTxt
        End If
    Next I
End Sub




'Character Editor
Private Sub cmdCharEditor_Click()
    ' Send request for character names
    Tex_CharSprite.Texture = 0
    SendRequestAllCharacters
End Sub

Private Sub cmdLevelUp_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < STAFF_DEVELOPER Then
        AddText "You have insufficent access to do this!", BrightRed
        Exit Sub
    End If

    SendRequestLevelUp
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdLevelUp_Click", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdALoc_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < STAFF_MAPPER Then
        AddText "You have insufficent access to do this!", BrightRed
        Exit Sub
    End If
    
    BLoc = Not BLoc
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdALoc_Click", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub



Private Sub cmdAWarpToMe_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < STAFF_MAPPER Then
        AddText "You have insufficent access to do this!", BrightRed
        Exit Sub
    End If

    If Len(Trim$(cmbPlayersOnline.text)) < 1 Then Exit Sub
    If IsNumeric(Trim$(cmbPlayersOnline.text)) Then Exit Sub

    WarpToMe Trim$(cmbPlayersOnline.text)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdAWarpToMe_Click", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdAWarpMeTo_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < STAFF_MAPPER Then
        AddText "You have insufficent access to do this!", BrightRed
        Exit Sub
    End If

    ' Subscript out of range
    If Len(Trim$(cmbPlayersOnline.text)) < 1 Then Exit Sub
    If IsNumeric(Trim$(cmbPlayersOnline.text)) Then Exit Sub

    WarpMeTo Trim$(cmbPlayersOnline.text)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdAWarpMeTo_Click", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdAWarp_Click()
    Dim n As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < STAFF_MAPPER Then
        AddText "You have insufficent access to do this!", BrightRed
        Exit Sub
    End If

    If Len(Trim$(txtAMap.text)) < 1 Then Exit Sub
    If Not IsNumeric(Trim$(txtAMap.text)) Then Exit Sub
    
    n = CLng(Trim$(txtAMap.text))

    ' Check to make sure its a valid map #
    If n > 0 And n <= MAX_MAPS Then
        MapEditorLeaveMap
        Call WarpTo(n)
    Else
        Call AddText("Invalid map number.", BrightRed)
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdAWarp_Click", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdAMapReport_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < STAFF_MAPPER Then
        AddText "You have insufficent access to do this!", BrightRed
        Exit Sub
    End If

    SendMapReport
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdAMapReport_Click", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdARespawn_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < STAFF_MAPPER Then
        AddText "You have insufficent access to do this!", BrightRed
        Exit Sub
    End If
    
    SendMapRespawn
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdARespawn_Click", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdAKick_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < STAFF_MODERATOR Then
        AddText "You have insufficent access to do this!", BrightRed
        Exit Sub
    End If

    If Len(Trim$(cmbPlayersOnline.text)) < 1 Then Exit Sub

    SendKick Trim$(cmbPlayersOnline.text)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdAKick_Click", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdABan_Click()
    Dim StrInput As String

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < STAFF_ADMIN Then
        AddText "You have insufficent access to do this!", BrightRed
        Exit Sub
    End If

    If Len(Trim$(cmbPlayersOnline.text)) < 1 Then Exit Sub

    StrInput = InputBox("Reason: ", "Ban")

    SendBan Trim$(cmbPlayersOnline.text), Trim$(StrInput)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdABan_Click", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdAMute_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < STAFF_MODERATOR Then
        AddText "You have insufficent access to do this!", BrightRed
        Exit Sub
    End If

    If Len(Trim$(cmbPlayersOnline.text)) < 1 Then Exit Sub

    SendMute Trim$(cmbPlayersOnline.text)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdAMute_Click", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdLoadTextures_Click()
    Dim I As Long
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    For I = 1 To NumTextures
        If gTexture(I).Timer < timeGetTime And gTexture(I).Timer <> 0 Then
            UnsetTexture I
            DoEvents
        End If
    Next

    LoadTextures
    Exit Sub
    
' Error Handler
ErrorHandler:
    HandleError "cmdLoadTextures_Click", "frmAdmin", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdSpawnRecent_Click()
    Dim Item As Byte
    Dim I    As Byte

    If GetPlayerAccess(MyIndex) < STAFF_DEVELOPER Then
        AddText "You have insufficent access to do this!", BrightRed
        Exit Sub
    End If

    Item = lastSpawnedItems(rcSwitcher.Value)

    SendSpawnItem Item, CLng(txtRecentAmount), True

    Dim found As Integer, limit As Integer

    For I = 0 To UBound(lastSpawnedItems) - 1
        If lastSpawnedItems(I) = Item Then
            found = I
            Exit For
        End If
    Next
    
    If found = -1 Then
        If UBound(lastSpawnedItems) = 20 Then
            DeleteByPtr lastSpawnedItems, 20
        End If

        InsertByPtr lastSpawnedItems, 0
    Else
        DeleteByPtr lastSpawnedItems, found
        InsertByPtr lastSpawnedItems, 0
    End If

    lastSpawnedItems(0) = Item
    frmAdmin.UpdateRecentSpawner
    If FormVisible("frmItemSpawner") Then
        If frmItemSpawner.tabItems.SelectedItem.Index = 1 Then
            frmItemSpawner.updatingItem = True
            frmItemSpawner.tabItems_Click
        End If
    End If
    Exit Sub
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyInsert
            If Player(MyIndex).Access >= STAFF_MODERATOR Then
                If frmAdmin.Visible And GetForegroundWindow = frmAdmin.hWnd Then
                    Unload frmAdmin
                End If
            End If
    End Select
End Sub

Private Sub optCat_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 10 Then Exit Sub
    
    Select Case Index
        Case 0
            lblCat.Caption = "Recent"
        Case 1
            lblCat.Caption = "None"
        Case 2
            lblCat.Caption = "Equipment"
        Case 3
            lblCat.Caption = "Consumable"
        Case 4
            lblCat.Caption = "Title"
        Case 5
            lblCat.Caption = "Spell"
        Case 6
            lblCat.Caption = "Teleport"
        Case 7
            lblCat.Caption = "Reset Stats"
        Case 8
            lblCat.Caption = "Auto Life"
        Case 9
            lblCat.Caption = "Change Sprite"
        'Case 10
        '    lblCat.Caption = "Recipe"
        
    End Select
End Sub

Public Sub optCat_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim test As Boolean
    
    If Index = 10 Then Exit Sub
    
    If ignoreChange Then
        ignoreChange = False
        Exit Sub
    End If
    
    If optCat(Index).Value = False Then
        optCat(Index).Picture = LoadResPicture(100 + Index, vbResBitmap)
    Else
        optCat(Index).Picture = LoadResPicture(110 + Index, vbResBitmap)
    If lastIndex = Index And optCat(Index).Value = True Then
        frmAdmin.currentCategory = "Categories"
    Else
        Select Case Index
            Case 0
                currentCategory = "Recent"
            Case 1
                currentCategory = "None"
            Case 2
                currentCategory = "Equipment"
            Case 3
                currentCategory = "Consumable"
            Case 4
                currentCategory = "Title"
            Case 5
                currentCategory = "Spell"
            Case 6
                currentCategory = "Teleport"
            Case 7
                currentCategory = "Reset Stats"
            Case 8
                currentCategory = "Auto Life"
            Case 9
                currentCategory = "Change Sprite"
            'Case 10
            '    currencycategory = "Recipe"
        End Select
        lblCat.Caption = currentCategory
    End If
        If lastIndex <> -1 And lastIndex <> 10 Then
            If optCat(lastIndex).Value = False Then
                optCat(lastIndex).Picture = LoadResPicture(100 + lastIndex, vbResBitmap)
            End If
        End If
        
        If Button <> 0 Then
            If lastIndex = Index Then
                Unload frmItemSpawner
                frmItemSpawner.lastTab = -1
                optCat(Index).Value = False
                optCat(Index).Picture = LoadResPicture(100 + Index, vbResBitmap)
                lastIndex = -1
                Exit Sub
            Else
                frmItemSpawner.Visible = True
                picSpawner.Visible = True
                ignoreChange = True
                frmItemSpawner.tabItems.SelectedItem = frmItemSpawner.tabItems.Tabs(Index + 1)
                frmItemSpawner.updateFreeSlots
                frmItemSpawner.tabItems.Tabs(Index + 1).Selected = True
                BringWindowToTop (frmItemSpawner.hWnd)
            End If
        End If

        lastIndex = Index
    End If
End Sub

Private Sub picEye_Click(Index As Integer)
    Select Case Index
        Case 0 ' Animation
            If chkEditor(Index).Value = 1 Then
                BringWindowToTop (frmEditor_Animation.hWnd)
            End If
            Exit Sub
        Case 1 ' Ban
            If chkEditor(Index).Value = 1 Then
                BringWindowToTop (frmEditor_Ban.hWnd)
            End If
            Exit Sub
        Case 2 'Class
            If chkEditor(Index).Value = 1 Then
                BringWindowToTop (frmEditor_Class.hWnd)
            End If
            Exit Sub
        Case 3 ' Emoticons
            If chkEditor(Index).Value = 1 Then
                BringWindowToTop (frmEditor_Emoticon.hWnd)
            End If
            Exit Sub
        Case 4 ' Items
            If chkEditor(Index).Value = 1 Then
                BringWindowToTop (frmEditor_Item.hWnd)
            End If
            Exit Sub
        Case 5 ' Map
            If chkEditor(Index).Value = 1 Then
                BringWindowToTop (frmEditor_Map.hWnd)
            End If
            Exit Sub
        Case 6 ' Moral
            If chkEditor(Index).Value = 1 Then
                BringWindowToTop (frmEditor_Moral.hWnd)
            End If
            Exit Sub
        Case 7 'NPC
            If chkEditor(Index).Value = 1 Then
                BringWindowToTop (frmEditor_NPC.hWnd)
            End If
            Exit Sub
         Case 8 ' Resource
            If chkEditor(Index).Value = 1 Then
                BringWindowToTop (frmEditor_Resource.hWnd)
            End If
            Exit Sub
        Case 9 'Shop
            If chkEditor(Index).Value = 0 Then
                BringWindowToTop (frmEditor_Shop.hWnd)
            End If
            Exit Sub
        Case 10 ' Spell
            If chkEditor(Index).Value = 0 Then
                BringWindowToTop (frmEditor_Spell.hWnd)
            End If
            Exit Sub
        Case 11 ' Title
            If chkEditor(Index).Value = 0 Then
                BringWindowToTop (frmEditor_Title.hWnd)
            End If
            Exit Sub
        Case 12 ' Events
            If chkEditor(Index).Value = 1 Then
                BringWindowToTop (frmEditor_Events.hWnd)
            End If
            Exit Sub
        Case 13 ' Quest
            If chkEditor(Index).Value = 1 Then
                BringWindowToTop (frmEditor_Quest.hWnd)
            End If
            Exit Sub

    End Select
End Sub

Private Sub picPanel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picRefresh.Picture = LoadResPicture("REFRESH_UP", vbResBitmap)
    lblCat.Caption = currentCategory
End Sub

Private Sub picPanel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    refreshDown = False
End Sub

Private Sub picRefresh_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    refreshDown = True
    picRefresh.Picture = LoadResPicture("REFRESH_DOWN", vbResBitmap)
End Sub

Private Sub picRefresh_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not refreshDown Then
        picRefresh.Picture = LoadResPicture("REFRESH_OVER", vbResBitmap)
    End If
End Sub

Private Sub picRefresh_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    refreshDown = False
    refreshingAdminList = True
    SendRequestPlayersOnline
End Sub

Public Sub selectMyself()
Dim I As Integer
    For I = 0 To cmbPlayersOnline.ListCount
        If Trim$(cmbPlayersOnline.List(I)) = Trim$(Player(MyIndex).Name) Then
            cmbPlayersOnline.ListIndex = I
            cmbPlayersOnline_Click
            Exit Sub
        End If
    Next
End Sub

Public Sub UpdatePlayersOnline()
    Dim players() As String, Staff() As String, tempTxt As String, temp() As String, Length As Long, I As Long, currentIgnore As Long
    Dim stuffCounter As Long, playersCounter As Long, overallCounter As Long, foundStuff As Boolean, foundPlayer As Boolean
    
    tempTxt = cmbPlayersOnline.text
    cmbPlayersOnline.Clear
    cmbPlayersOnline.text = tempTxt
    
    ' Get Stuff
    For I = 0 To UBound(g_playersOnline)
        If CByte(Split(g_playersOnline(I), ":")(1)) > 0 Then
            foundStuff = True
            ReDim Preserve Staff(stuffCounter)
            Staff(stuffCounter) = Split(g_playersOnline(I), ":")(0)
            stuffCounter = stuffCounter + 1
        End If
    Next
    
    'Get Players
    For I = 0 To UBound(g_playersOnline)
        If CByte(Split(g_playersOnline(I), ":")(1)) = 0 Then
            foundPlayer = True
            ReDim Preserve players(playersCounter)
            players(playersCounter) = Split(g_playersOnline(I), ":")(0)
            playersCounter = playersCounter + 1
        End If
    Next
    
    If foundStuff Then
        cmbPlayersOnline.AddItem ("Staff: " & stuffCounter)
        
            ReDim Preserve ignoreIndexes(0)
            ignoreIndexes(0) = currentIgnore
            currentIgnore = currentIgnore + 1
            
        For I = 0 To UBound(Staff)
            cmbPlayersOnline.AddItem (Trim$(Staff(I)))
            currentIgnore = currentIgnore + 1
        Next
        overallCounter = overallCounter + stuffCounter
    End If

    If foundPlayer Then
        cmbPlayersOnline.AddItem ("Players: " & playersCounter)
        
            ReDim Preserve ignoreIndexes(1)
            ignoreIndexes(1) = currentIgnore
            currentIgnore = currentIgnore + 1
        For I = 0 To UBound(players)
            cmbPlayersOnline.AddItem (Trim$(players(I)))
            currentIgnore = currentIgnore + 1
        Next
        overallCounter = overallCounter + playersCounter
    End If
    
    lblPlayers.Caption = "Players: " & overallCounter
End Sub

Public Sub styleButtons()
    Dim I As Long, temp1 As Long, temp2 As Long
    
    For I = 0 To optCat.UBound
        optCat(I).Value = False
        optCat(I).Picture = LoadResPicture(100 + I, vbResBitmap)
    Next
    
    For I = 0 To picEye.UBound
        picEye(I).Visible = False
        picEye(I).Picture = LoadResPicture("BRING_FRONT", vbResBitmap)
    Next
    
    temp1 = getWndProcAddr
    
    If GetWindowLong(optCat(0).hWnd, -4) <> temp1 Then
        For I = 0 To optCat.UBound
            SubClassHwnd optCat(I).hWnd
        Next
        For I = 0 To chkEditor.UBound
            picEye(I).BorderStyle = 0
            SubClassHwnd chkEditor(I).hWnd
        Next
        catSub = True
        picSpawner.Picture = LoadResPicture("BRING_FRONT", vbResBitmap)
        picSpawner.BorderStyle = 0
    End If
End Sub

Public Sub findVisibleEditors()
    Dim I As Long, tempCtl As Control, frm As Form
    
    For Each frm In Forms
        If frm.Visible = True Then
            Select Case frm.Name
                Case "frmEditor_Animation"
                    ignoreChange = True
                    chkEditor(EDITOR_ANIMATION).Value = 1
                    chkEditor(EDITOR_ANIMATION).FontBold = True
                    picEye(EDITOR_ANIMATION).Visible = True
                Case "frmEditor_Ban"
                    ignoreChange = True
                    chkEditor(EDITOR_BAN).Value = 1
                    chkEditor(EDITOR_BAN).FontBold = True
                    picEye(EDITOR_BAN).Visible = True
                Case "frmEditor_Class"
                    ignoreChange = True
                    chkEditor(EDITOR_CLASS).Value = 1
                    chkEditor(EDITOR_CLASS).FontBold = True
                    picEye(EDITOR_CLASS).Visible = True
                Case "frmEditor_Emoticon"
                    ignoreChange = True
                    chkEditor(EDITOR_EMOTICON).Value = 1
                    chkEditor(EDITOR_EMOTICON).FontBold = True
                    picEye(EDITOR_EMOTICON).Visible = True
                Case "frmEditor_Events"
                    ignoreChange = True
                    chkEditor(EDITOR_EVENTS).Value = 1
                    chkEditor(EDITOR_EVENTS).FontBold = True
                    picEye(EDITOR_EVENTS).Visible = True
                Case "frmEditor_Item"
                    ignoreChange = True
                    chkEditor(EDITOR_ITEM).Value = 1
                    chkEditor(EDITOR_ITEM).FontBold = True
                    picEye(EDITOR_ITEM).Visible = True
                Case "frmEditor_Map"
                    ignoreChange = True
                    chkEditor(EDITOR_MAP).Value = 1
                    chkEditor(EDITOR_MAP).FontBold = True
                    picEye(EDITOR_MAP).Visible = True
                Case "frmEditor_Moral"
                    ignoreChange = True
                    chkEditor(EDITOR_MORAL).Value = 1
                    chkEditor(EDITOR_MORAL).FontBold = True
                    picEye(EDITOR_MORAL).Visible = True
                Case "frmEditor_NPC"
                    ignoreChange = True
                    chkEditor(EDITOR_NPC).Value = 1
                    chkEditor(EDITOR_NPC).FontBold = True
                    picEye(EDITOR_NPC).Visible = True
                Case "frmEditor_Resource"
                    ignoreChange = True
                    chkEditor(EDITOR_RESOURCE).Value = 1
                    chkEditor(EDITOR_RESOURCE).FontBold = True
                    picEye(EDITOR_RESOURCE).Visible = True
                Case "frmEditor_Shop"
                    ignoreChange = True
                    chkEditor(EDITOR_SHOP).Value = 1
                    chkEditor(EDITOR_SHOP).FontBold = True
                    picEye(EDITOR_SHOP).Visible = True
                Case "frmEditor_Spell"
                    ignoreChange = True
                    chkEditor(EDITOR_SPELL).Value = 1
                    chkEditor(EDITOR_SPELL).FontBold = True
                    picEye(EDITOR_SPELL).Visible = True
                Case "frmEditor_Title"
                    ignoreChange = True
                    chkEditor(EDITOR_TITLE).Value = 1
                    chkEditor(EDITOR_TITLE).FontBold = True
                    picEye(EDITOR_TITLE).Visible = True
                Case "frmItemSpawner"
                    picSpawner.Visible = True
            End Select
        End If
    Next
End Sub

Public Sub UpdateRecentSpawner()
    If ArrayIsInitialized(lastSpawnedItems) Then
        If UBound(lastSpawnedItems) > 0 Then
            lblItemName.Enabled = True
            txtRecentAmount.Enabled = True
            cmdSpawnRecent.Enabled = True
            rcSwitcher.max = UBound(lastSpawnedItems) - 1
            rcSwitcher.Enabled = True
            picRecentItem.Enabled = True
            lblItemName.Caption = Item(lastSpawnedItems(rcSwitcher.Value)).Name
            txtRecentAmount.text = 1
        End If
    End If
End Sub

Public Sub Form_Load()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ignoreChange = False
    frmAdmin.picRefresh.BorderStyle = 0
    
    Me.Move frmMain.Left + frmMain.Width, frmMain.Top
    If Trim$(cmbPlayersOnline.text) = "Choose Player" Then
        txtSprite.Enabled = False
        upSprite.Enabled = False
    End If
    
    lastIndex = -1
    styleButtons
    
    findVisibleEditors
    
    upSprite.max = NumCharacters
    upSprite.min = 0
    
    currentCategory = lblCat.Caption
    
    LastAdminSpriteTimer = timeGetTime
    
    UpdateRecentSpawner
    picRefresh.Picture = LoadResPicture("REFRESH_UP", vbResBitmap)
    refreshingAdminList = True
    SendRequestPlayersOnline
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Form_Load", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub picSizer_Click()
    Dim ctrl As Control
    
    If adminMin And reverse Then
        For Each ctrl In Controls
            Select Case ctrl.Name
                Case "lblEditors", "picSizer", "chkEditor", "picEye"
                    ctrl.Left = ctrl.Left + 100
                    ctrl.Top = ctrl.Top + 206
                Case "lineEditors"
                    ctrl.X1 = ctrl.X1 + 100
                    ctrl.X2 = ctrl.X2 + 100
                    ctrl.Y1 = ctrl.Y1 + 206
                    ctrl.Y2 = ctrl.Y2 + 206
                Case "picPanel"
                Case Else
                    ctrl.Visible = True
            End Select
        Next
        Width = 3135
        adminMin = False
        picSizer.Picture = LoadResPicture("MIN", vbResBitmap)
        picPanel.Top = picPanel.Top + 2
        BorderStyle = 1
        Caption = "Admin Panel"
        Height = 9060
        frmAdmin.Left = frmMain.Left + frmMain.Width + 145
        frmAdmin.Top = frmMain.Top
        frmAdmin.picSpawner.Visible = False
    Else
        For Each ctrl In Controls
            Select Case ctrl.Name
            
                Case "lblEditors", "picSizer", "chkEditor", "picEye"
                    ctrl.Left = ctrl.Left - 100
                    ctrl.Top = ctrl.Top - 206
                Case "lineEditors"
                    ctrl.X1 = ctrl.X1 - 100
                    ctrl.X2 = ctrl.X2 - 100
                    ctrl.Y1 = ctrl.Y1 - 206
                    ctrl.Y2 = ctrl.Y2 - 206
                Case "picPanel"
                Case Else
                    ctrl.Visible = False
            End Select
        Next

        Width = 1485
        Height = 4680
        picPanel.Top = picPanel.Top - 2
        BorderStyle = 4
        Caption = "Mini Panel"
        adminMin = True
        picSizer.Picture = LoadResPicture("MAX", vbResBitmap)
        frmAdmin.centerMiniVert frmMain.Width, frmMain.Height, frmMain.Left, frmMain.Top + 145
        frmAdmin.picSpawner.Visible = False
    End If
    reverse = True
End Sub

Public Sub centerMiniVert(pWidth As Long, pHeight As Long, pLeft As Long, pTop As Long)
    Left = pLeft + pWidth + 165
    Top = pTop + ((pHeight - Height) / 2)
    BeforeTopMost (frmAdmin.hWnd)
    BringWindowToTop (hwndLastActiveWnd)
End Sub

Private Sub picSpawner_Click()
    If FormVisible("frmItemSpawner") Then
        BringWindowToTop (frmItemSpawner.hWnd)
    End If
End Sub

Private Sub rcSwitcher_Change()
    frmAdmin.UpdateRecentSpawner
End Sub

Private Sub txtAMap_GotFocus()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    txtAMap.SelStart = Len(txtAMap)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtAMap_GotFocus", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Function correctValue(ByRef textBox As textBox, ByRef valueToChange, min As Long, max As Long, Optional defaultVal As Long = 0) As Boolean
    Dim test As textBox, TempValue As String
    
    If textBox.text = "" Then
        textBox.text = CStr(defaultVal)
        valueToChange = defaultVal
        correctValue = True
    End If

    If Len(textBox.text) = 1 And InStr(1, textBox.text, "-") = 1 Then
        correctValue = True
        Exit Function
    ElseIf Len(textBox.text) = 1 And IsNumeric(textBox.text) Then
        If verifyValue(textBox, min, max) Then
            TempValue = textBox.text
            valueToChange = TempValue
            correctValue = True
        Else
            textBox.text = CStr(valueToChange)
            textBox.SelStart = Len(textBox.text)
            correctValue = False
        End If
    ElseIf Len(textBox.text) > 1 And InStr(1, textBox.text, "-") = 0 And InStrRev(textBox.text, "-") = 0 And IsNumeric(textBox.text) Then

        If verifyValue(textBox, min, max) Then
            TempValue = textBox.text
            valueToChange = TempValue
            correctValue = True
        Else
            textBox.text = CStr(valueToChange)
            textBox.SelStart = Len(textBox.text)
            correctValue = False
        End If

    ElseIf Len(textBox.text) > 1 And InStr(1, textBox.text, "-") = 1 And InStrRev(textBox.text, "-") = 1 And IsNumeric(textBox.text) Then

        If verifyValue(textBox, min, max) Then
            TempValue = textBox.text
            valueToChange = TempValue
            correctValue = True
        Else
            textBox.text = CStr(valueToChange)
            textBox.SelStart = Len(textBox.text)
        correctValue = False
        End If
        
    Else
        textBox.text = CStr(valueToChange)
        textBox.SelStart = Len(textBox.text)
        correctValue = False
    End If
End Function

Private Sub reviseValue(ByRef textBox As textBox, ByRef valueToChange)
    If Not IsNumeric(textBox.text) Then
        textBox.text = CStr(valueToChange)
    Else
        textBox.text = CStr(valueToChange)
    End If
End Sub

Private Function verifyValue(txtBox As textBox, min As Long, max As Long)
    Dim Msg As String
    
    If (CLng(txtBox.text) >= min And CLng(txtBox.text) <= max) Then
        verifyValue = True
    Else
        Msg = " field accepts only values: " & CStr(min) & " < value < " & CStr(max) & "." & vbCrLf & "Reverting value..."
        verifyValue = False
    End If
End Function

Private Sub selectValue(ByRef textBox As textBox)
    textBox.SelStart = 0
    textBox.SelLength = Len(textBox.text)
End Sub

Private Sub txtRecentAmount_Change()
    correctValue txtRecentAmount, currentRecentAm, 0, 999999
End Sub

Private Sub txtRecentAmount_Click()
    selectValue txtRecentAmount
End Sub

Private Sub txtRecentAmount_GotFocus()
    selectValue txtRecentAmount
End Sub

Private Sub txtRecentAmount_LostFocus()
    reviseValue txtRecentAmount, currentRecentAm
End Sub

Private Sub txtSprite_Change()
    Dim I As Long
    If autoSprite Then
        autoSprite = False
        Exit Sub
    End If
    
     If correctValue(txtSprite, currentSprite, 0, NumCharacters) Then
        If txtSprite.text = 0 Then picSprite.Picture = Nothing
        If GetPlayerAccess(MyIndex) < STAFF_ADMIN Then
            AddText "You have insufficent access to do this!", BrightRed
            Exit Sub
        ElseIf txtSprite.text > 0 Then
            For I = 0 To UBound(g_playersOnline)
                If InStr(1, g_playersOnline(I), cmbPlayersOnline.text) Then
                    Mid$(g_playersOnline(I), InStr(InStr(1, g_playersOnline(I), ":") + 1, g_playersOnline(I), ":"), Len(txtSprite.text) + 1) = ":" & txtSprite.text
                End If
            Next I

            SendSetPlayerSprite Trim$(cmbPlayersOnline.text), currentSprite
        End If
     End If
End Sub

Private Sub txtSprite_Click()
     selectValue txtSprite
End Sub

Private Sub txtSprite_GotFocus()
    selectValue txtSprite
End Sub

Private Sub txtSprite_LostFocus()
    reviseValue txtSprite, currentSprite
End Sub

