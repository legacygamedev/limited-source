VERSION 5.00
Begin VB.Form frmNpcEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NPC Editor (Index: 0)"
   ClientHeight    =   7695
   ClientLeft      =   150
   ClientTop       =   420
   ClientWidth     =   9735
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNpcEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   513
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   649
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Settings"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   4920
      TabIndex        =   49
      Top             =   120
      Width           =   4695
      Begin VB.CheckBox chkNight 
         Caption         =   "Night"
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
         Left            =   3720
         TabIndex        =   55
         Top             =   600
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkDay 
         Caption         =   "Day"
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
         Left            =   2640
         TabIndex        =   54
         Top             =   600
         Value           =   1  'Checked
         Width           =   615
      End
      Begin VB.ComboBox cmbBehavior 
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
         ItemData        =   "frmNpcEditor.frx":0FC2
         Left            =   240
         List            =   "frmNpcEditor.frx":0FD8
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   1200
         Width           =   4215
      End
      Begin VB.TextBox txtSpawnSecs 
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
         Left            =   240
         MaxLength       =   10
         TabIndex        =   52
         Text            =   "0"
         Top             =   600
         Width           =   1695
      End
      Begin VB.HScrollBar scrlScript 
         Height          =   255
         Left            =   240
         Max             =   10000
         TabIndex        =   51
         Top             =   2520
         Width           =   4215
      End
      Begin VB.HScrollBar scrlElement 
         Height          =   255
         Left            =   240
         Max             =   1000
         TabIndex        =   50
         Top             =   1920
         Value           =   1
         Width           =   4215
      End
      Begin VB.Label Label2 
         Caption         =   "Behavior:"
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
         Left            =   240
         TabIndex        =   62
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label16 
         Caption         =   "Spawn Rate (In Seconds):"
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
         Left            =   240
         TabIndex        =   61
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Spawn Time:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   2640
         TabIndex        =   60
         Top             =   360
         Width           =   840
      End
      Begin VB.Label lblScript 
         Caption         =   "Script:"
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
         Left            =   240
         TabIndex        =   59
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label lblScriptNum 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
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
         TabIndex        =   58
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Element:"
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
         Left            =   240
         TabIndex        =   57
         Top             =   1680
         Width           =   555
      End
      Begin VB.Label lblElement 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "None"
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
         Left            =   3000
         TabIndex        =   56
         Top             =   1680
         Width           =   1410
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Drop Table"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   4920
      TabIndex        =   30
      Top             =   3360
      Width           =   4695
      Begin VB.HScrollBar scrlChance 
         Height          =   255
         Left            =   240
         Max             =   10000
         TabIndex        =   63
         Top             =   2400
         Value           =   1
         Width           =   4215
      End
      Begin VB.HScrollBar scrlValue 
         Height          =   255
         Left            =   240
         Max             =   10000
         TabIndex        =   33
         Top             =   1800
         Value           =   1
         Width           =   4215
      End
      Begin VB.HScrollBar scrlNum 
         Height          =   255
         Left            =   240
         Max             =   500
         TabIndex        =   32
         Top             =   1200
         Value           =   1
         Width           =   4215
      End
      Begin VB.HScrollBar scrlDropItem 
         Height          =   255
         Left            =   240
         Max             =   5
         Min             =   1
         TabIndex        =   31
         Top             =   480
         Value           =   1
         Width           =   4215
      End
      Begin VB.Label lblChance 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "1 In X"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   3240
         TabIndex        =   64
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Chance:"
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
         Left            =   240
         TabIndex        =   41
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label lblValue 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
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
         Left            =   4005
         TabIndex        =   40
         Top             =   1560
         Width           =   435
      End
      Begin VB.Label Label7 
         Caption         =   "Value:"
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
         Left            =   240
         TabIndex        =   39
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lblNum 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
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
         Left            =   4000
         TabIndex        =   38
         Top             =   960
         Width           =   435
      End
      Begin VB.Label Label9 
         Caption         =   "Number:"
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
         Left            =   240
         TabIndex        =   37
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblItemName 
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
         Left            =   960
         TabIndex        =   36
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label lblDropItem 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "1"
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
         Left            =   4005
         TabIndex        =   35
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label13 
         Caption         =   "Dropping :"
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
         Left            =   240
         TabIndex        =   34
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "General Information"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7455
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4695
      Begin VB.TextBox txtName 
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
         Left            =   840
         TabIndex        =   46
         Top             =   360
         Width           =   3615
      End
      Begin VB.TextBox txtAttackSay 
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
         Left            =   840
         TabIndex        =   45
         Top             =   720
         Width           =   3615
      End
      Begin VB.OptionButton Opt64 
         Caption         =   "64x32"
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
         Left            =   2280
         TabIndex        =   44
         Top             =   2520
         Width           =   735
      End
      Begin VB.OptionButton Opt32 
         Caption         =   "32x32"
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
         Left            =   2280
         TabIndex        =   43
         Top             =   2280
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.HScrollBar ExpGive 
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   6960
         Width           =   4215
      End
      Begin VB.HScrollBar StartHP 
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   6360
         Width           =   4215
      End
      Begin VB.CheckBox BigNpc 
         Caption         =   "Big NPC"
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
         Left            =   2280
         TabIndex        =   23
         Top             =   1800
         Width           =   855
      End
      Begin VB.PictureBox picSprite 
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
         Height          =   480
         Left            =   3600
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   9
         Top             =   2160
         Width           =   480
         Begin VB.PictureBox picSprites 
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
            Height          =   12000
            Left            =   120
            ScaleHeight     =   800
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   456
            TabIndex        =   42
            Top             =   120
            Width           =   6840
         End
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Left            =   240
         Max             =   500
         TabIndex        =   8
         Top             =   1320
         Width           =   4215
      End
      Begin VB.HScrollBar scrlRange 
         Height          =   255
         Left            =   240
         Max             =   30
         TabIndex        =   7
         Top             =   3360
         Width           =   4215
      End
      Begin VB.HScrollBar scrlSTR 
         Height          =   255
         Left            =   240
         Max             =   1000
         TabIndex        =   6
         Top             =   3960
         Width           =   4215
      End
      Begin VB.HScrollBar scrlDEF 
         Height          =   255
         Left            =   240
         Max             =   1000
         TabIndex        =   5
         Top             =   4560
         Width           =   4215
      End
      Begin VB.HScrollBar scrlSPEED 
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         Max             =   1000
         TabIndex        =   4
         Top             =   5160
         Width           =   4215
      End
      Begin VB.HScrollBar scrlMAGI 
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         Max             =   1000
         TabIndex        =   3
         Top             =   5760
         Width           =   4215
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1155
         Left            =   3240
         ScaleHeight     =   75
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   78
         TabIndex        =   22
         Top             =   1800
         Width           =   1200
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
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
         Left            =   240
         TabIndex        =   48
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label14 
         Caption         =   "Speak:"
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
         Left            =   240
         TabIndex        =   47
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblExpGiven 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
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
         TabIndex        =   28
         Top             =   6720
         Width           =   495
      End
      Begin VB.Label lblNpcExp 
         Caption         =   "Experience:"
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
         Left            =   240
         TabIndex        =   27
         Top             =   6720
         Width           =   735
      End
      Begin VB.Label lblStartHP 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
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
         TabIndex        =   26
         Top             =   6120
         Width           =   495
      End
      Begin VB.Label lblNpcHP 
         Caption         =   "Hit Points:"
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
         Left            =   240
         TabIndex        =   24
         Top             =   6120
         Width           =   735
      End
      Begin VB.Label lblSprite 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
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
         TabIndex        =   21
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblNpcSprite 
         Caption         =   "Sprite:"
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
         Left            =   240
         TabIndex        =   20
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblRange 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
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
         TabIndex        =   19
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label lblNpcSight 
         BackStyle       =   0  'Transparent
         Caption         =   "Sight (By Tile):"
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
         Left            =   240
         TabIndex        =   18
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label lblSTR 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
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
         TabIndex        =   17
         Top             =   3720
         Width           =   495
      End
      Begin VB.Label lblNpcStr 
         Caption         =   "Strength:"
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
         Left            =   240
         TabIndex        =   16
         Top             =   3720
         Width           =   615
      End
      Begin VB.Label lblDEF 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
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
         TabIndex        =   15
         Top             =   4320
         Width           =   495
      End
      Begin VB.Label lblNpcDef 
         Caption         =   "Defense:"
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
         Left            =   240
         TabIndex        =   14
         Top             =   4320
         Width           =   615
      End
      Begin VB.Label lblSPEED 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
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
         TabIndex        =   13
         Top             =   4920
         Width           =   495
      End
      Begin VB.Label lblNpcSpd 
         Caption         =   "Speed:"
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
         Left            =   240
         TabIndex        =   12
         Top             =   4920
         Width           =   615
      End
      Begin VB.Label lblMAGI 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
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
         TabIndex        =   11
         Top             =   5520
         Width           =   495
      End
      Begin VB.Label lblNpcMagic 
         Caption         =   "Magic:"
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
         Left            =   240
         TabIndex        =   10
         Top             =   5520
         Width           =   495
      End
   End
   Begin VB.Timer tmrSprite 
      Interval        =   50
      Left            =   4920
      Top             =   6360
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
      Height          =   375
      Left            =   7680
      TabIndex        =   1
      Top             =   7200
      Width           =   1935
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Apply"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   0
      Top             =   7200
      Width           =   1935
   End
End
Attribute VB_Name = "frmNpcEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BigNpc_Click()
    frmNpcEditor.ScaleMode = 3

    If BigNpc.Value = Checked Then
        picSprite.Width = 960
        picSprite.Height = 960
        picSprite.Top = 1900
        picSprite.Left = 3360

        picSprites.Picture = LoadPicture(App.Path & "\GFX\BigSprites.bmp")
    Else
        picSprite.Width = 480
        picSprite.Left = 3600

        If Opt64.Value Then
            picSprite.Height = 960
            picSprite.Top = 1900
        Else
            picSprite.Height = 480
            picSprite.Top = 2160
        End If

        picSprites.Picture = LoadPicture(App.Path & "\GFX\Sprites.bmp")
    End If
End Sub

Private Sub chkDay_Click()
    If chkNight.Value = Unchecked Then
        If chkDay.Value = Unchecked Then
            chkDay.Value = Checked
        End If
    End If
End Sub

Private Sub chkNight_Click()
    If chkDay.Value = Unchecked Then
        If chkNight.Value = Unchecked Then
            chkNight.Value = Checked
        End If
    End If
End Sub

Private Sub ExpGive_Change()
    lblExpGiven.Caption = CStr(ExpGive.Value)
End Sub

Private Sub Form_Load()
    frmNpcEditor.Caption = "NPC Editor (Index: " & EditorIndex & ")"

    scrlElement.Max = MAX_ELEMENTS
    scrlDropItem.Max = MAX_NPC_DROPS

    If SpriteSize = 1 Then
        picSprite.Height = 960
    End If

    picSprites.Picture = LoadPicture(App.Path & "\GFX\Sprites.bmp")

    If Opt64.Value = Checked Then
        picSprite.Height = 960
        picSprite.Top = 1900
    Else
        picSprite.Height = 480
        picSprite.Top = 2160
    End If
End Sub

Private Sub Opt32_Click()
    If Not BigNpc.Value = Checked Then
        picSprites.Picture = LoadPicture(App.Path & "\GFX\Sprites.bmp")

        picSprite.Height = 480
        picSprite.Top = 2160
    End If
End Sub

Private Sub Opt64_Click()
    If Not BigNpc.Value = Checked Then
        picSprites.Picture = LoadPicture(App.Path & "\GFX\Sprites.bmp")

        picSprite.Height = 960
        picSprite.Top = 1900
    End If
End Sub

Private Sub scrlDropItem_Change()
    scrlNum.Value = Npc(EditorIndex).ItemNPC(scrlDropItem.Value).ItemNum
    scrlValue.Value = Npc(EditorIndex).ItemNPC(scrlDropItem.Value).ItemValue
    scrlChance.Value = Npc(EditorIndex).ItemNPC(scrlDropItem.Value).chance

    lblDropItem.Caption = CStr(scrlDropItem.Value)
End Sub

Private Sub scrlElement_Change()
    lblElement.Caption = CStr(scrlElement.Value)
End Sub

Private Sub scrlSprite_Change()
    lblSprite.Caption = CStr(scrlSprite.Value)
End Sub

Private Sub scrlRange_Change()
    lblRange.Caption = CStr(scrlRange.Value)
End Sub

Private Sub scrlSTR_Change()
    lblSTR.Caption = CStr(scrlSTR.Value)
End Sub

Private Sub scrlDEF_Change()
    lblDEF.Caption = CStr(scrlDEF.Value)
End Sub

Private Sub scrlSPEED_Change()
    lblSPEED.Caption = CStr(scrlSPEED.Value)
End Sub

Private Sub scrlMAGI_Change()
    lblMAGI.Caption = CStr(scrlMAGI.Value)
End Sub

Private Sub scrlNum_Change()
    lblNum.Caption = CStr(scrlNum.Value)
    lblItemName.Caption = vbNullString

    If scrlNum.Value > 0 Then
        lblItemName.Caption = Trim$(Item(scrlNum.Value).Name)
    End If

    Npc(EditorIndex).ItemNPC(scrlDropItem.Value).ItemNum = scrlNum.Value
End Sub

Private Sub scrlValue_Change()
    Npc(EditorIndex).ItemNPC(scrlDropItem.Value).ItemValue = scrlValue.Value

    lblValue.Caption = CStr(scrlValue.Value)
End Sub

Private Sub cmdOk_Click()
    Call NpcEditorOk
End Sub

Private Sub cmdCancel_Click()
    Call NpcEditorCancel
End Sub

Private Sub StartHP_Change()
    lblStartHP.Caption = CStr(StartHP.Value)
End Sub

Private Sub tmrSprite_Timer()
    Call NpcEditorBltSprite
End Sub

Private Sub scrlChance_Change()
    lblChance.Caption = "1 In " & scrlChance.Value
    Npc(EditorIndex).ItemNPC(scrlDropItem.Value).chance = scrlChance.Value
End Sub

Private Sub cmbBehavior_Click()
    If cmbBehavior.ListIndex = NPC_BEHAVIOR_SCRIPTED Then
        scrlScript.Enabled = True
    Else
        scrlScript.Enabled = False
    End If
End Sub

Private Sub scrlScript_Change()
    lblScriptNum.Caption = CStr(scrlScript.Value)
End Sub
