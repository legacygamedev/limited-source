VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loading..."
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9360
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   9360
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   5953
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Console"
      TabPicture(0)   =   "frmServer.frx":1708A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtChat"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtText"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Players"
      TabPicture(1)   =   "frmServer.frx":170A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdEditPlayer"
      Tab(1).Control(1)=   "cmdSavePlayers"
      Tab(1).Control(2)=   "lvwInfo"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Control "
      TabPicture(2)   =   "frmServer.frx":170C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraDatabase"
      Tab(2).Control(1)=   "fraServer"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "News"
      TabPicture(3)   =   "frmServer.frx":170DE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdLoadNews"
      Tab(3).Control(1)=   "cmdSaveNews"
      Tab(3).Control(2)=   "txtNews"
      Tab(3).ControlCount=   3
      Begin VB.TextBox txtText 
         Appearance      =   0  'Flat
         Height          =   2415
         Left            =   3360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   48
         Top             =   420
         Width           =   5595
      End
      Begin VB.TextBox txtChat 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3360
         TabIndex        =   47
         Top             =   3000
         Width           =   5655
      End
      Begin VB.Frame Frame4 
         Caption         =   "Connectivity"
         Height          =   2895
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   3015
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Packets out/Sec:"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CPS:"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   2400
            Width           =   495
         End
         Begin VB.Label lblCpsLock 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "[Unlock]"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   480
            TabIndex        =   44
            Top             =   2400
            Width           =   720
         End
         Begin VB.Label lblCPS 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "xxxx"
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   1380
            TabIndex        =   43
            Top             =   2400
            Width           =   1455
         End
         Begin VB.Label lblTime 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "xx:xx:xx"
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   1140
            TabIndex        =   42
            Top             =   2100
            Width           =   1695
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Online For:"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   2100
            Width           =   1095
         End
         Begin VB.Line Line 
            Index           =   2
            X1              =   120
            X2              =   2880
            Y1              =   1920
            Y2              =   1920
         End
         Begin VB.Label lblPackOut 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "xxxx"
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   1620
            TabIndex        =   40
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label lblPackIn 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "xxxx"
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   1620
            TabIndex        =   39
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Packets in/Sec:"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Line Line 
            Index           =   1
            X1              =   120
            X2              =   2880
            Y1              =   1320
            Y2              =   1320
         End
         Begin VB.Label lblPlayers 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "xx:xx"
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   1500
            TabIndex        =   37
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Players Online:"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label lblPort 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "xxxx"
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   1260
            TabIndex        =   35
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label lblIP 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "xxx.xxx.xxx.xxx"
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   1260
            TabIndex        =   34
            Top             =   240
            Width           =   1575
         End
         Begin VB.Line Line 
            Index           =   0
            X1              =   120
            X2              =   2880
            Y1              =   840
            Y2              =   840
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Port:"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IP Address:"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmdEditPlayer 
         Caption         =   "Edit Players"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -74880
         TabIndex        =   30
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CommandButton cmdSavePlayers 
         Caption         =   "Save Players"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -73560
         TabIndex        =   29
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CommandButton cmdLoadNews 
         Caption         =   "Load News"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -67680
         TabIndex        =   18
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton cmdSaveNews 
         Caption         =   "Save News"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   17
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtNews 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   960
         Width           =   8835
      End
      Begin VB.Frame fraServer 
         Caption         =   "Server"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   -69360
         TabIndex        =   1
         Top             =   480
         Width           =   1575
         Begin VB.CommandButton cmdSet 
            Caption         =   "Set"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   26
            Top             =   1920
            Width           =   1275
         End
         Begin VB.TextBox txtExpRate 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   25
            Text            =   "1"
            Top             =   1560
            Width           =   1275
         End
         Begin VB.CheckBox chkServerLog 
            Caption         =   "Server Log"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   2350
            Value           =   1  'Checked
            Width           =   1260
         End
         Begin VB.CommandButton cmdExit 
            Caption         =   "Exit"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   720
            Width           =   1275
         End
         Begin VB.CommandButton cmdShutDown 
            Caption         =   "Shut Down"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label lblExpRate 
            Alignment       =   2  'Center
            Caption         =   "Exp Rate:"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   1320
            Width           =   795
         End
      End
      Begin VB.Frame fraDatabase 
         Caption         =   "Reload"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   -74880
         TabIndex        =   4
         Top             =   480
         Width           =   5415
         Begin VB.CommandButton cmdReloadQuests 
            Caption         =   "Quests"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2760
            TabIndex        =   28
            Top             =   2160
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadEmoticons 
            Caption         =   "Emoticons"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2760
            TabIndex        =   27
            Top             =   1680
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadTitles 
            Caption         =   "Titles"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2760
            TabIndex        =   24
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadMorals 
            Caption         =   "Morals"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2760
            TabIndex        =   23
            Top             =   720
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadBans 
            Caption         =   "Bans"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2760
            TabIndex        =   20
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadAll 
            Caption         =   "All"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadGuilds 
            Caption         =   "Guilds"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1440
            TabIndex        =   14
            Top             =   2160
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadOptions 
            Caption         =   "Options"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4080
            TabIndex        =   13
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadAnimations 
            Caption         =   "Animations"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1440
            TabIndex        =   12
            Top             =   1680
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadResources 
            Caption         =   "Resources"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1440
            TabIndex        =   11
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadItems 
            Caption         =   "Items"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1440
            TabIndex        =   10
            Top             =   720
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadNPCs 
            Caption         =   "NPCs"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1440
            TabIndex        =   9
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadShops 
            Caption         =   "Shops"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   8
            Top             =   2160
            Width           =   1215
         End
         Begin VB.CommandButton CmdReloadSpells 
            Caption         =   "Spells"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   7
            Top             =   1680
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadMaps 
            Caption         =   "Maps"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   6
            Top             =   720
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadClasses 
            Caption         =   "Classes"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   1200
            Width           =   1215
         End
      End
      Begin MSComctlLib.ListView lvwInfo 
         Height          =   2355
         Left            =   -74880
         TabIndex        =   19
         Top             =   480
         Width           =   8835
         _ExtentX        =   15584
         _ExtentY        =   4154
         View            =   3
         Arrange         =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Index"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "IP Address"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Account"
            Object.Width           =   3527
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Character"
            Object.Width           =   3529
         EndProperty
      End
   End
   Begin VB.Menu mnuKick 
      Caption         =   "&Kick"
      Visible         =   0   'False
      Begin VB.Menu mnuKickPlayer 
         Caption         =   "Kick"
      End
      Begin VB.Menu mnuDisconnectPlayer 
         Caption         =   "Disconnect"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBanPlayer 
         Caption         =   "Ban"
      End
      Begin VB.Menu mnuAdminPlayer 
         Caption         =   "Make Owner"
      End
      Begin VB.Menu mnuRemoveAdmin 
         Caption         =   "Remove Access"
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSavePlayers_Click()
    Dim i As Long
    
    For i = 1 To Player_HighIndex
        Call SaveAccount(i)
    Next
End Sub

' ********************
' ** Winsock object **
' ********************
Private Sub Socket_ConnectionRequest(index As Integer, ByVal requestID As Long)
    Call AcceptConnection(index, requestID)
End Sub

Private Sub Socket_Accept(index As Integer, SocketId As Integer)
    Call AcceptConnection(index, SocketId)
End Sub

Private Sub Socket_DataArrival(index As Integer, ByVal bytesTotal As Long)
    If IsConnected(index) Then
        Call IncomingData(index, bytesTotal)
    End If
End Sub

Private Sub Socket_Close(index As Integer)
    Call CloseSocket(index)
End Sub

Private Sub chkServerLog_Click()
    ' If it's not 0, then it's true
    If Not chkServerLog.Value Then
        ServerLog = True
    Else
        ServerLog = False
    End If
End Sub

Private Sub cmdExit_Click()
    Call DestroyServer
End Sub

Private Sub cmdLoadNews_Click()
    txtNews.Text = Options.News
End Sub

Private Sub cmdSaveNews_Click()
    Dim i As Long
    
    Options.News = txtNews.Text
    SaveOptions
    
    ' Send the news modified to any players not actually playing
    For i = 1 To Player_HighIndex
        If Not IsPlaying(i) And IsConnected(i) Then
            Call SendNews(i)
        End If
    Next
End Sub

Private Sub lblCPSLock_Click()
    If CPSUnlock Then
        CPSUnlock = False
        lblCpsLock.Caption = "[Unlock]"
    Else
        CPSUnlock = True
        lblCpsLock.Caption = "[Lock]"
    End If
End Sub

Public Sub cmdReloadAll_Click()
    Call cmdReloadClasses_Click
    Call cmdReloadMaps_Click
    Call cmdReloadSpells_Click
    Call cmdReloadShops_Click
    Call cmdReloadNPCs_Click
    Call cmdReloadItems_Click
    Call cmdReloadResources_Click
    Call cmdReloadAnimations_Click
    Call cmdReloadBans_Click
    Call cmdReloadMorals_Click
    Call cmdReloadTitles_Click
    Call cmdReloadOptions_Click
    Call cmdReloadEmoticons_Click
    Call cmdReLoadGuilds_Click
    Call cmdReloadQuests_Click
End Sub

Public Sub cmdReloadTitles_Click()
    Dim i As Long
    
    LoadTitles
    Call TextAdd("All titles reloaded.")
    
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendTitles i
        End If
    Next
End Sub

Public Sub cmdReloadMorals_Click()
    Dim i As Long
    
    LoadMorals
    Call TextAdd("All morals reloaded.")
    
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendMorals i
        End If
    Next
End Sub

Public Sub cmdReloadBans_Click()
    LoadBans
    Call TextAdd("All bans reloaded.")
End Sub

Public Sub cmdReloadOptions_Click()
    LoadOptions
    LoadDataSizes
    Call TextAdd("All options reloaded.")
End Sub

Public Sub cmdReloadClasses_Click()
    Dim i As Long
    
    Call LoadClasses
    Call TextAdd("All Classes reloaded.")
    
    For i = 1 To Player_HighIndex
        If IsConnected(i) Then
            SendClasses i
        End If
    Next
End Sub

Public Sub cmdReloadItems_Click()
    Dim i As Long
    
    Call LoadItems
    Call TextAdd("All items reloaded.")
    
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendItems i
        End If
    Next
End Sub

Public Sub cmdReloadMaps_Click()
    Dim i As Long
    
    Call LoadMaps
    Call CreateFullMapCache
    Call TextAdd("All maps reloaded.")
    
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            PlayerWarp i, GetPlayerMap(i), GetPlayerX(i), GetPlayerY(i), True
        End If
    Next
End Sub

Public Sub cmdReloadNPCs_Click()
    Dim i As Long
    
    Call LoadNPCs
    Call TextAdd("All npcs reloaded.")
    
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendNPCs i
        End If
    Next
End Sub

Public Sub cmdReloadShops_Click()
    Dim i As Long
    
    Call LoadShops
    Call TextAdd("All shops reloaded.")
    
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendShops i
        End If
    Next
End Sub

Public Sub cmdReloadSpells_Click()
    Dim i As Long
    
    Call LoadSpells
    Call TextAdd("All spells reloaded.")
    
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendSpells i
        End If
    Next
End Sub

Public Sub cmdReloadResources_Click()
    Dim i As Long
    
    Call LoadResources
    Call TextAdd("All resources reloaded.")
    
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendResources i
        End If
    Next
End Sub

Public Sub cmdReloadAnimations_Click()
    Dim i As Long
    
    Call LoadAnimations
    Call TextAdd("All animations reloaded.")
    
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendAnimations i
        End If
    Next
End Sub

Public Sub cmdReloadEmoticons_Click()
    Dim i As Long
    
    LoadEmoticons
    Call TextAdd("All emoticons reloaded.")
    
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendEmoticons i
        End If
    Next
End Sub

Public Sub cmdReLoadGuilds_Click()
    Dim i As Long
    
    LoadGuilds
    Call TextAdd("All guilds reloaded.")
    
    ' Update guilds
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendPlayerGuild i
        End If
    Next
    
    ' Update guild list
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If GetPlayerGuild(i) > 0 Then
                SendPlayerGuildMembers i
            End If
        End If
    Next
End Sub

Public Sub cmdReloadQuests_Click()
    Dim i As Long
    
    Call LoadQuests
    Call TextAdd("All quests reloaded.")
    
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendQuests i
        End If
    Next
End Sub

Private Sub cmdSet_Click()
    EXP_RATE = txtExpRate.Text
   
    If EXP_RATE > 1 Then
        Call GlobalMsg("The experience rate has been changed to " & EXP_RATE & "x!", Yellow)
    ElseIf EXP_RATE = 0 Then
        Call GlobalMsg("The experience rate has been frozen!", BrightBlue)
    Else
        Call GlobalMsg("The experience rate has been changed back to normal.", BrightGreen)
    End If
End Sub

Private Sub cmdShutDown_Click()
    If IsShuttingDown Then
        IsShuttingDown = False
        cmdShutDown.Caption = "Shutdown"
        GlobalMsg "Shutdown canceled.", BrightBlue
    Else
        IsShuttingDown = True
        cmdShutDown.Caption = "Cancel"
    End If
End Sub

Private Sub Form_Load()
    Call UsersOnline_Start
End Sub

Private Sub Form_Resize()
On Error Resume Next
    If frmServer.WindowState = vbMinimized Then
        frmServer.Hide
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
    Call DestroyServer
End Sub

Private Sub lvwInfo_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    ' When a ColumnHeader object is clicked, the ListView control is sorted by the subitems of that column.
    ' Set the SortKey to the Index of the ColumnHeader - 1
    ' Set Sorted to True to sort the list.
    
    If lvwInfo.SortOrder = lvwAscending Then
        lvwInfo.SortOrder = lvwDescending
    Else
        lvwInfo.SortOrder = lvwAscending
    End If

    lvwInfo.SortKey = ColumnHeader.index - 1
    lvwInfo.Sorted = True
End Sub

Private Sub txtExpRate_Change()
    If Not IsNumeric(txtExpRate.Text) Then txtExpRate.Text = 1
    If txtExpRate.Text < 0 Then txtExpRate = 0
    If txtExpRate.Text > 1000 Then txtExpRate.Text = 1000
End Sub

Private Sub txtText_GotFocus()
    txtChat.SetFocus
End Sub

Private Sub txtChat_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Len(Trim$(txtChat.Text)) > 0 Then
            Call GlobalMsg(txtChat.Text, White)
            Call TextAdd("Server: " & txtChat.Text)
            txtChat.Text = vbNullString
        End If

        KeyAscii = 0
    End If
End Sub

Sub UsersOnline_Start()
    Dim i As Integer

    For i = 1 To MAX_PLAYERS
        frmServer.lvwInfo.ListItems.Add (i)

        If i < 10 Then
            frmServer.lvwInfo.ListItems(i).Text = "00" & i
        ElseIf i < 100 Then
            frmServer.lvwInfo.ListItems(i).Text = "0" & i
        Else
            frmServer.lvwInfo.ListItems(i).Text = i
        End If

        frmServer.lvwInfo.ListItems(i).SubItems(1) = vbNullString
        frmServer.lvwInfo.ListItems(i).SubItems(2) = vbNullString
        frmServer.lvwInfo.ListItems(i).SubItems(3) = vbNullString
    Next

End Sub

Private Sub lvwInfo_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuKick
    End If
End Sub

Private Sub mnuKickPlayer_Click()
    Dim Name As String
    
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If IsPlaying(FindPlayer(Name)) Then
        tempplayer(FindPlayer(Name)).HasLogged = True
        Call AlertMsg(FindPlayer(Name), "You have been kicked by the server!")
        Call LeftGame(FindPlayer(Name))
    End If

End Sub

Sub mnuBanPlayer_click()
    Dim Name As String
    Dim index As Long
    
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)
    index = FindPlayer(Name)

    If index > 0 And index <= MAX_PLAYERS Then
        If IsConnected(index) Then
            Call BanIndex(index, "server", vbNullString)
        End If
    End If
End Sub

Sub mnuAdminPlayer_click()
    Dim Name As String
    
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = vbNullString Then
        Call SetPlayerAccess(FindPlayer(Name), 5)
        Call SendPlayerData(FindPlayer(Name))
        Call PlayerMsg(FindPlayer(Name), "You have been granted administrator access.", BrightCyan)
    End If

End Sub

Sub mnuRemoveAdmin_click()
    Dim Name As String
    
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = vbNullString Then
        Call SetPlayerAccess(FindPlayer(Name), 0)
        Call SendPlayerData(FindPlayer(Name))
        Call PlayerMsg(FindPlayer(Name), "You have had your administrator access revoked.", BrightRed)
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim lmsg As Long
    
    lmsg = x / Screen.TwipsPerPixelX

    Select Case lmsg
        Case WM_LBUTTONDBLCLK
            frmServer.WindowState = vbNormal
            frmServer.Show
            txtText.SelStart = Len(txtText.Text)
    End Select

End Sub
