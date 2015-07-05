VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmadmin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administration Panel"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4785
   LinkTopic       =   "frmadmin"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   4785
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   9763
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   353
      TabMaxWidth     =   1764
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Menu"
      TabPicture(0)   =   "frmadmin.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Line1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "btnclose"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Commands"
      TabPicture(1)   =   "frmadmin.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "GamePlay"
      TabPicture(2)   =   "frmadmin.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame9"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame10"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame11"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Credits"
      TabPicture(3)   =   "frmadmin.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label10"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label9"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label8"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label7"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Label6"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Frame12"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Frame13"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).ControlCount=   7
      Begin VB.CommandButton btnclose 
         Caption         =   "Close Admin Panel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   46
         Top             =   5160
         Width           =   1695
      End
      Begin VB.Frame Frame4 
         Caption         =   "Developer Commands"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   2520
         TabIndex        =   36
         Top             =   2040
         Width           =   1935
         Begin VB.CommandButton Command8 
            Caption         =   "Edit Options"
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
            TabIndex        =   54
            Top             =   2640
            Width           =   1695
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Edit Story"
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
            TabIndex        =   53
            Top             =   2400
            Width           =   1695
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Edit Emoticon"
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
            TabIndex        =   45
            Top             =   2160
            Width           =   1695
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Edit MOTD"
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
            TabIndex        =   44
            Top             =   1920
            Width           =   1695
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Edit Arrow"
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
            Top             =   1680
            Width           =   1695
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Edit Weather"
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
            Top             =   1440
            Width           =   1695
         End
         Begin VB.CommandButton tnEditNPC 
            Caption         =   "Edit NPC's"
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
            Top             =   1200
            Width           =   1695
         End
         Begin VB.CommandButton btnEditShops 
            Caption         =   "Edit Shops"
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
            Top             =   960
            Width           =   1695
         End
         Begin VB.CommandButton btnedititem 
            Caption         =   "Edit Items"
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
            TabIndex        =   39
            Top             =   720
            Width           =   1695
         End
         Begin VB.CommandButton btneditspell 
            Caption         =   "Edit Spells"
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
            TabIndex        =   38
            Top             =   480
            Width           =   1695
         End
         Begin VB.CommandButton btnMapeditor 
            Caption         =   "Map Editor"
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
            TabIndex        =   37
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Sprite Commands"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   2520
         TabIndex        =   31
         Top             =   720
         Width           =   1935
         Begin VB.CommandButton btnPlayerSprite 
            Caption         =   "Set Player Sprite"
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
            TabIndex        =   34
            Top             =   480
            Width           =   1695
         End
         Begin VB.TextBox txtSprite 
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
            TabIndex        =   33
            Top             =   960
            Width           =   1695
         End
         Begin VB.CommandButton btnSprite 
            Caption         =   "Set Sprite"
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
            TabIndex        =   32
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label3 
            Caption         =   "Sprite Number:"
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
            Top             =   745
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Map Commands"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   120
         TabIndex        =   25
         Top             =   3480
         Width           =   1935
         Begin VB.CommandButton btnWarpto 
            Caption         =   "Warp To"
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
            TabIndex        =   29
            Top             =   720
            Width           =   1695
         End
         Begin VB.CommandButton btnRespawn 
            Caption         =   "Respawn"
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
            TabIndex        =   28
            Top             =   480
            Width           =   1695
         End
         Begin VB.CommandButton btnLOC 
            Caption         =   "Location"
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
            TabIndex        =   27
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtMap 
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
            TabIndex        =   26
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "Map Number:"
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
            TabIndex        =   30
            Top             =   1080
            Width           =   1695
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Player Commands"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   1935
         Begin VB.CommandButton btnSetAccess 
            Caption         =   "Set Access"
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
            TabIndex        =   22
            Top             =   960
            Width           =   1695
         End
         Begin VB.CommandButton btnKick 
            Caption         =   "Kick"
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
            Top             =   720
            Width           =   1695
         End
         Begin VB.CommandButton btnWarpMeTo 
            Caption         =   "Warp Me To"
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
            Top             =   480
            Width           =   1695
         End
         Begin VB.TextBox txtPlayer 
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
            TabIndex        =   19
            Top             =   2280
            Width           =   1695
         End
         Begin VB.CommandButton btnBan 
            Caption         =   "Banish"
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
            TabIndex        =   18
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtAccess 
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
            TabIndex        =   17
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000010&
            X1              =   120
            X2              =   1800
            Y1              =   1920
            Y2              =   1920
         End
         Begin VB.Label Label4 
            Caption         =   "Access Level:"
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
            TabIndex        =   24
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Player Name:"
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
            TabIndex        =   23
            Top             =   2040
            Width           =   1575
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Game Commands"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4935
         Left            =   -74880
         TabIndex        =   14
         Top             =   360
         Width           =   4215
         Begin VB.TextBox Text9 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4575
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   15
            Text            =   "frmadmin.frx":0070
            Top             =   240
            Width           =   3975
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Ask For Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   -72720
         TabIndex        =   12
         Top             =   480
         Width           =   2175
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   13
            Text            =   "frmadmin.frx":06DC
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "More Buttons"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   -74880
         TabIndex        =   10
         Top             =   2160
         Width           =   4335
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            Text            =   "frmadmin.frx":074B
            Top             =   240
            Width           =   4095
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Controls"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   -74880
         TabIndex        =   8
         Top             =   480
         Width           =   2055
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   9
            Text            =   "frmadmin.frx":08DA
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Engine History"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -72480
         TabIndex        =   6
         Top             =   360
         Width           =   1935
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   7
            Text            =   "frmadmin.frx":0936
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Engine Creator"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   -74880
         TabIndex        =   4
         Top             =   360
         Width           =   2295
         Begin VB.TextBox Text7 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   5
            Text            =   "frmadmin.frx":0970
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Options"
         Height          =   855
         Left            =   -74880
         TabIndex        =   1
         Top             =   4200
         Width           =   1935
         Begin VB.CommandButton Command6 
            Caption         =   "Refresh"
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
            TabIndex        =   3
            Top             =   480
            Width           =   1695
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Check FPS"
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
            TabIndex        =   2
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Label Label5 
         Caption         =   "Administrative panel version 0.5"
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
         TabIndex        =   52
         Top             =   360
         Width           =   2175
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   2280
         X2              =   2280
         Y1              =   720
         Y2              =   5160
      End
      Begin VB.Label Label6 
         Caption         =   "http://www.Ngage-Online.de/forums/_hosted_/..."
         Height          =   255
         Left            =   -74880
         TabIndex        =   51
         Top             =   3600
         Width           =   4335
      End
      Begin VB.Label Label7 
         Caption         =   "Open the offical forum for Bevrias Engine:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   50
         Top             =   3360
         Width           =   3735
      End
      Begin VB.Label Label8 
         Caption         =   "http://www.Bevrias-Engine.tk"
         Height          =   255
         Left            =   -74880
         TabIndex        =   49
         Top             =   3120
         Width           =   3015
      End
      Begin VB.Label Label9 
         Caption         =   "http://www.Bevrias.com"
         Height          =   255
         Left            =   -74880
         TabIndex        =   48
         Top             =   2880
         Width           =   3375
      End
      Begin VB.Label Label10 
         Caption         =   "Open the offical site for Bevrias Engine:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   47
         Top             =   2640
         Width           =   3735
      End
   End
End
Attribute VB_Name = "frmadmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnPlayerSprite_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
        If Trim(txtPlayer.Text) <> "" Then
            If Trim(txtSprite.Text) <> "" Then
                Call SendSetPlayerSprite(Trim(txtPlayer.Text), Trim(txtSprite.Text))
            End If
        End If
    Else: Call AddText("You are not authorized to carry out that action", BrightRed)
    End If
End Sub
Private Sub btnBan_Click()
If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
Call SendBan(Trim(txtPlayer.Text))
Else: Call AddText("You are not authorized to carry out that action", BrightRed)
End If
End Sub




Private Sub btnedititem_Click()
If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
Call SendRequestEditItem
Else: Call AddText("You are not authorized to carry out that action", BrightRed)
End If
End Sub

Private Sub btnEditNPC_Click()
If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
Call SendRequestEditNpc
Else: Call AddText("You are not authorized to carry out that action", BrightRed)
End If
End Sub

Private Sub btnEditShops_Click()
If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
Call SendRequestEditShop
Else: Call AddText("You are not authorized to carry out that action", BrightRed)
End If
End Sub

Private Sub btneditspell_Click()
If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
Call SendRequestEditSpell
Else: Call AddText("You are not authorized to carry out that action", BrightRed)
End If
End Sub

Private Sub btnkick_Click()
If GetPlayerAccess(MyIndex) >= ADMIN_MONITER Then
Call SendKick(Trim(txtPlayer.Text))
Else: Call AddText("You are not authorized to carry out that action", BrightRed)
End If
End Sub

Private Sub btnLOC_Click()
If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
Call SendRequestLocation
Else: Call AddText("You are not authorized to carry out that action", BrightRed)
End If
End Sub

Private Sub btnMapeditor_Click()
If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
Call SendRequestEditMap
Else: Call AddText("You are not authorized to carry out that action", BrightRed)
End If
End Sub
Private Sub btnRespawn_Click()
If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
Call SendMapRespawn
Else: Call AddText("You are not authorized to carry out that action", BrightRed)
End If
End Sub
Private Sub btnWarpmeTo_Click()
If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
Call WarpMeTo(Trim(txtPlayer.Text))
Else: Call AddText("You are not authorized to carry out that action", BrightRed)
End If
End Sub

Private Sub btnWarpto_Click()
If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
Call WarpTo(Val(txtMap.Text))
Else: Call AddText("You are not authorized to carry out that action", BrightRed)
End If
End Sub

Private Sub btnWarptome_Click()
If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
Call WarpToMe(Trim(txtPlayer.Text))
Else: Call AddText("You are not authorized to carry out that action", BrightRed)
End If
End Sub

Private Sub btnclose_Click()
frmadmin.Visible = False
End Sub

Private Sub btnSprite_Click()
If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
frmSprite.Visible = True
Else: Call AddText("You are not authorized to carry out that action", BrightRed)
End If
End Sub

Private Sub btnSetAccess_Click()
   If GetPlayerAccess(MyIndex) >= ADMIN_CREATOR Then
      Call SendSetAccess(Trim(txtPlayer.Text), Trim(txtAccess.Text))
   Else: Call AddText("You are not authorized to carry out that action", BrightRed)
   End If
End Sub

Private Sub Command1_Click()
frmadmin.Visible = False
frmWeather.Visible = True
End Sub

Private Sub Command2_Click()
frmadmin.Visible = False
Call SendRequestEditArrow
End Sub

Private Sub Command3_Click()
frmadmin.Visible = False
frmEditMOTD.Visible = True
End Sub

Private Sub Command4_Click()
Call AddText("FPS: " & GameFPS, Yellow)
End Sub

Private Sub Command5_Click()
frmadmin.Visible = False
frmEmoticonEditor.Visible = True
End Sub

Private Sub Command6_Click()
Call SendData("refresh" & SEP_CHAR & END_CHAR)
Call AddText("Screen is Refreshed!", Yellow)
End Sub

Private Sub Command7_Click()
frmEditStory.Visible = True
End Sub

Private Sub Command8_Click()
frmOptions.Visible = True
End Sub

Private Sub Label6_Click()
Shell ("explorer http://www.ngage-online.de/forums/_hosted_/bevrias-engine/index.php"), vbNormalNoFocus
End Sub

Private Sub Label8_Click()
Shell ("explorer http://www.Bevrias-engine.tk"), vbNormalNoFocus
End Sub

Private Sub Label9_Click()
Shell ("explorer http://www.Bevrias.com"), vbNormalNoFocus
End Sub

Private Sub tnEditNPC_Click()
If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
Call SendRequestEditNpc
Else: Call AddText("You are not authorized to carry out that action", BrightRed)
End If
End Sub

