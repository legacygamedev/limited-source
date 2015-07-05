VERSION 5.00
Begin VB.Form frmNpcEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Npc Editor"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9945
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   455
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   663
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar ScrlDarkDEF 
      Height          =   135
      Left            =   6000
      Max             =   1000
      TabIndex        =   95
      Top             =   4920
      Width           =   2895
   End
   Begin VB.HScrollBar ScrlDarkSTR 
      Height          =   135
      Left            =   960
      Max             =   1000
      TabIndex        =   94
      Top             =   6600
      Width           =   2895
   End
   Begin VB.HScrollBar ScrlLightDEF 
      Height          =   135
      Left            =   6000
      Max             =   1000
      TabIndex        =   93
      Top             =   4680
      Width           =   2895
   End
   Begin VB.HScrollBar ScrlLightSTR 
      Height          =   135
      Left            =   960
      Max             =   1000
      TabIndex        =   92
      Top             =   6360
      Width           =   2895
   End
   Begin VB.HScrollBar ScrlColdDEF 
      Height          =   135
      Left            =   6000
      Max             =   1000
      TabIndex        =   87
      Top             =   4440
      Width           =   2895
   End
   Begin VB.HScrollBar ScrlColdSTR 
      Height          =   135
      Left            =   960
      Max             =   1000
      TabIndex        =   86
      Top             =   6120
      Width           =   2895
   End
   Begin VB.HScrollBar ScrlHeatDEF 
      Height          =   135
      Left            =   6000
      Max             =   1000
      TabIndex        =   85
      Top             =   4200
      Width           =   2895
   End
   Begin VB.HScrollBar ScrlHeatSTR 
      Height          =   135
      Left            =   960
      Max             =   1000
      TabIndex        =   84
      Top             =   5880
      Width           =   2895
   End
   Begin VB.HScrollBar ScrlAirDEF 
      Height          =   135
      Left            =   6000
      Max             =   1000
      TabIndex        =   78
      Top             =   3960
      Width           =   2895
   End
   Begin VB.HScrollBar ScrlAirSTR 
      Height          =   135
      Left            =   960
      Max             =   1000
      TabIndex        =   74
      Top             =   5640
      Width           =   2895
   End
   Begin VB.HScrollBar ScrlEarthDEF 
      Height          =   135
      Left            =   6000
      Max             =   1000
      TabIndex        =   72
      Top             =   3720
      Width           =   2895
   End
   Begin VB.HScrollBar ScrlEarthSTR 
      Height          =   135
      Left            =   960
      Max             =   1000
      TabIndex        =   69
      Top             =   5400
      Width           =   2895
   End
   Begin VB.HScrollBar ScrlWaterDEF 
      Height          =   135
      Left            =   6000
      Max             =   1000
      TabIndex        =   65
      Top             =   3480
      Width           =   2895
   End
   Begin VB.HScrollBar ScrlWaterSTR 
      Height          =   135
      Left            =   960
      Max             =   1000
      TabIndex        =   62
      Top             =   5160
      Width           =   2895
   End
   Begin VB.HScrollBar ScrlFireSTR 
      Height          =   135
      Left            =   960
      Max             =   1000
      TabIndex        =   59
      Top             =   4920
      Width           =   2895
   End
   Begin VB.HScrollBar ScrlFireDEF 
      Height          =   135
      Left            =   6000
      Max             =   1000
      TabIndex        =   56
      Top             =   3240
      Width           =   2895
   End
   Begin VB.Frame Frame3 
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   5160
      TabIndex        =   35
      Top             =   240
      Width           =   4695
      Begin VB.TextBox txtSpawnSecs 
         Alignment       =   1  'Right Justify
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
         Left            =   2760
         TabIndex        =   54
         Text            =   "0"
         Top             =   2520
         Width           =   1815
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
         ItemData        =   "frmNpcEditor.frx":0000
         Left            =   960
         List            =   "frmNpcEditor.frx":0013
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   2160
         Width           =   3615
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
         Left            =   3840
         TabIndex        =   49
         Top             =   1800
         Value           =   1  'Checked
         Width           =   615
      End
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
         Left            =   3120
         TabIndex        =   50
         Top             =   1800
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.TextBox txtChance 
         Alignment       =   1  'Right Justify
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
         Left            =   2640
         TabIndex        =   39
         Text            =   "0"
         Top             =   1440
         Width           =   1815
      End
      Begin VB.HScrollBar scrlValue 
         Height          =   135
         Left            =   960
         Max             =   10000
         TabIndex        =   38
         Top             =   1200
         Value           =   1
         Width           =   3255
      End
      Begin VB.HScrollBar scrlNum 
         Height          =   135
         Left            =   960
         Max             =   1000
         TabIndex        =   37
         Top             =   960
         Value           =   1
         Width           =   3255
      End
      Begin VB.HScrollBar scrlDropItem 
         Height          =   135
         Left            =   960
         Max             =   30
         Min             =   1
         TabIndex        =   36
         Top             =   360
         Value           =   1
         Width           =   3255
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Spawn Time :"
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
         Left            =   2160
         TabIndex        =   55
         Top             =   1800
         Width           =   885
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Spawn Rate (in seconds) :"
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
         TabIndex        =   53
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Behavior :"
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
         Left            =   120
         TabIndex        =   51
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Drop Item Chance 1 out of :"
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
         Left            =   720
         TabIndex        =   48
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label lblValue 
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
         Height          =   165
         Left            =   4300
         TabIndex        =   47
         Top             =   1200
         Width           =   75
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Value :"
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
         TabIndex        =   46
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lblNum 
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
         Height          =   165
         Left            =   4300
         TabIndex        =   45
         Top             =   960
         Width           =   75
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Number :"
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
         Top             =   960
         Width           =   735
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
         TabIndex        =   43
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Item :"
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
         Left            =   120
         TabIndex        =   42
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblDropItem 
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
         Height          =   165
         Left            =   4300
         TabIndex        =   41
         Top             =   360
         Width           =   75
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
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
         Left            =   120
         TabIndex        =   40
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "General Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   4695
      Begin VB.HScrollBar ExpGive 
         Height          =   135
         Left            =   1080
         TabIndex        =   34
         Top             =   3480
         Width           =   2895
      End
      Begin VB.HScrollBar StartHP 
         Height          =   135
         Left            =   1080
         TabIndex        =   30
         Top             =   3240
         Width           =   2895
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
         TabIndex        =   28
         Top             =   1560
         Width           =   855
      End
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
         Height          =   480
         Left            =   480
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   26
         Top             =   720
         Width           =   480
         Visible         =   0   'False
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
         Left            =   1400
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   13
         Top             =   1040
         Width           =   480
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   135
         Left            =   1080
         Max             =   500
         TabIndex        =   12
         Top             =   360
         Width           =   2895
      End
      Begin VB.HScrollBar scrlRange 
         Height          =   135
         Left            =   1080
         Max             =   30
         TabIndex        =   11
         Top             =   2040
         Value           =   1
         Width           =   2895
      End
      Begin VB.HScrollBar scrlSTR 
         Height          =   135
         Left            =   1080
         Max             =   1000
         TabIndex        =   10
         Top             =   2280
         Width           =   2895
      End
      Begin VB.HScrollBar scrlDEF 
         Height          =   135
         Left            =   1080
         Max             =   1000
         TabIndex        =   9
         Top             =   2520
         Width           =   2895
      End
      Begin VB.HScrollBar scrlSPEED 
         Enabled         =   0   'False
         Height          =   135
         Left            =   1080
         Max             =   1000
         TabIndex        =   8
         Top             =   2760
         Width           =   2895
      End
      Begin VB.HScrollBar scrlMAGI 
         Enabled         =   0   'False
         Height          =   135
         Left            =   1080
         Max             =   1000
         TabIndex        =   7
         Top             =   3000
         Width           =   2895
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1125
         Left            =   1080
         ScaleHeight     =   73
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   73
         TabIndex        =   27
         Top             =   720
         Width           =   1125
      End
      Begin VB.Label lblExpGiven 
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
         Height          =   135
         Left            =   4080
         TabIndex        =   33
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Exp Given :"
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
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label lblStartHP 
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
         Height          =   135
         Left            =   4080
         TabIndex        =   31
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Starting Hp :"
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
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label lblSprite 
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
         Height          =   375
         Left            =   4080
         TabIndex        =   25
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Sprite :"
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
         Left            =   360
         TabIndex        =   24
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblRange 
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
         Left            =   4080
         TabIndex        =   23
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
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
         Height          =   375
         Left            =   0
         TabIndex        =   22
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label lblSTR 
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
         Left            =   4080
         TabIndex        =   21
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Strength :"
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
         Left            =   360
         TabIndex        =   20
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label lblDEF 
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
         Left            =   4080
         TabIndex        =   19
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Defence :"
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
         Left            =   360
         TabIndex        =   18
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label lblSPEED 
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
         Height          =   135
         Left            =   4080
         TabIndex        =   17
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Speed :"
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
         Left            =   360
         TabIndex        =   16
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label lblMAGI 
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
         Height          =   135
         Left            =   4080
         TabIndex        =   15
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Magic :"
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
         Left            =   360
         TabIndex        =   14
         Top             =   3000
         Width           =   615
      End
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
      Left            =   960
      TabIndex        =   5
      Top             =   600
      Width           =   3975
   End
   Begin VB.Timer tmrSprite 
      Interval        =   50
      Left            =   9360
      Top             =   5880
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
      Height          =   255
      Left            =   6360
      TabIndex        =   3
      Top             =   6480
      Width           =   1695
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
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
      Left            =   8160
      TabIndex        =   2
      Top             =   6480
      Width           =   1695
   End
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
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label LblDarkDEF 
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
      Left            =   9000
      TabIndex        =   103
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label LblDarkSTR 
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
      TabIndex        =   102
      Top             =   6600
      Width           =   615
   End
   Begin VB.Label LblLightDEF 
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
      Left            =   9000
      TabIndex        =   101
      Top             =   4680
      Width           =   615
   End
   Begin VB.Label LblLightSTR 
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
      TabIndex        =   100
      Top             =   6360
      Width           =   615
   End
   Begin VB.Label LblColdDEF 
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
      Left            =   9000
      TabIndex        =   99
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label LblColdSTR 
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
      TabIndex        =   98
      Top             =   6120
      Width           =   735
   End
   Begin VB.Label LblHeatDEF 
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
      Left            =   9000
      TabIndex        =   97
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label LblHeatSTR 
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
      TabIndex        =   96
      Top             =   5880
      Width           =   615
   End
   Begin VB.Label Label35 
      Alignment       =   1  'Right Justify
      Caption         =   "Dark DEF:"
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
      Left            =   5040
      TabIndex        =   91
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label Label34 
      Alignment       =   1  'Right Justify
      Caption         =   "Dark STR:"
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
      Left            =   -120
      TabIndex        =   90
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label Label33 
      Alignment       =   1  'Right Justify
      Caption         =   "Light DEF:"
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
      Left            =   5040
      TabIndex        =   89
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Label32 
      Alignment       =   1  'Right Justify
      Caption         =   "Light STR:"
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
      Left            =   0
      TabIndex        =   88
      Top             =   6360
      Width           =   855
   End
   Begin VB.Label Label31 
      Alignment       =   1  'Right Justify
      Caption         =   "Cold DEF:"
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
      Left            =   5040
      TabIndex        =   83
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label29 
      Alignment       =   1  'Right Justify
      Caption         =   "Cold STR:"
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
      Left            =   0
      TabIndex        =   82
      Top             =   6120
      Width           =   855
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      Caption         =   "Heat DEF:"
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
      Left            =   5040
      TabIndex        =   81
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label27 
      Alignment       =   1  'Right Justify
      Caption         =   "Heat STR:"
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
      Left            =   0
      TabIndex        =   80
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label LblAirDEF 
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
      Left            =   9000
      TabIndex        =   79
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Label30 
      Alignment       =   1  'Right Justify
      Caption         =   "Air DEF:"
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
      Left            =   5040
      TabIndex        =   77
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label LblAirSTR 
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
      TabIndex        =   76
      Top             =   5640
      Width           =   615
   End
   Begin VB.Label Label25 
      Alignment       =   1  'Right Justify
      Caption         =   "Air STR:"
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
      Left            =   0
      TabIndex        =   75
      Top             =   5640
      Width           =   855
   End
   Begin VB.Label LblEarthDEF 
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
      Left            =   9000
      TabIndex        =   73
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      Caption         =   "Earth DEF:"
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
      Left            =   5040
      TabIndex        =   71
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label LblEarthSTR 
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
      TabIndex        =   70
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      Caption         =   "Earth STR:"
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
      Left            =   0
      TabIndex        =   68
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label LblWaterDEF 
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
      Left            =   9000
      TabIndex        =   67
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      Caption         =   "Water DEF:"
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
      Left            =   5040
      TabIndex        =   66
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label LblWaterSTR 
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
      TabIndex        =   64
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      Caption         =   "Water STR:"
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
      Left            =   0
      TabIndex        =   63
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label LblFireSTR 
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
      TabIndex        =   61
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      Caption         =   "Fire STR:"
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
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label LblFireDEF 
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
      Left            =   9000
      TabIndex        =   58
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      Caption         =   "Fire DEF:"
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
      Left            =   5160
      TabIndex        =   57
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "Speak :"
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
      TabIndex        =   4
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Name :"
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
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   495
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
        frmNpcEditor.picSprites.Picture = LoadPicture(App.Path & "\GFX\bigsprites.bmp")
        picSprite.Width = 960
        picSprite.Height = 960
        picSprite.Top = 800
        picSprite.Left = 1170
    Else
        frmNpcEditor.picSprites.Picture = LoadPicture(App.Path & "\GFX\sprites.bmp")
        picSprite.Width = 480
        picSprite.Height = 480
        picSprite.Top = 1040
        picSprite.Left = 1400
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
    lblExpGiven.Caption = ExpGive.Value
End Sub

Private Sub Form_Load()
    scrlDropItem.Max = MAX_NPC_DROPS
    picSprites.Picture = LoadPicture(App.Path & "\GFX\sprites.bmp")
End Sub

Private Sub scrlDropItem_Change()
    txtChance.Text = Npc(EditorIndex).ItemNPC(scrlDropItem.Value).Chance
    scrlNum.Value = Npc(EditorIndex).ItemNPC(scrlDropItem.Value).ItemNum
    scrlValue.Value = Npc(EditorIndex).ItemNPC(scrlDropItem.Value).ItemValue
    lblDropItem.Caption = scrlDropItem.Value
End Sub

Private Sub ScrlEarthDEF_Change()
    LblEarthDEF.Caption = STR(ScrlEarthDEF.Value)
End Sub

Private Sub ScrlEarthSTR_Change()
    LblEarthSTR.Caption = STR(ScrlEarthSTR.Value)
End Sub
Private Sub ScrlAirDEF_Change()
    LblAirDEF.Caption = STR(ScrlAirDEF.Value)
End Sub

Private Sub ScrlAirSTR_Change()
    LblAirSTR.Caption = STR(ScrlAirSTR.Value)
End Sub
Private Sub ScrlFireDEF_Change()
    LblFireDEF.Caption = STR(ScrlFireDEF.Value)
End Sub

Private Sub ScrlFireSTR_Change()
    LblFireSTR.Caption = STR(ScrlFireSTR.Value)
End Sub

Private Sub scrlSprite_Change()
    lblSprite.Caption = STR(scrlSprite.Value)
End Sub

Private Sub scrlRange_Change()
    lblRange.Caption = STR(scrlRange.Value)
End Sub

Private Sub scrlSTR_Change()
    lblSTR.Caption = STR(scrlSTR.Value)
End Sub

Private Sub scrlDEF_Change()
    lblDEF.Caption = STR(scrlDEF.Value)
End Sub

Private Sub scrlSPEED_Change()
    lblSPEED.Caption = STR(scrlSPEED.Value)
End Sub

Private Sub scrlMAGI_Change()
    lblMAGI.Caption = STR(scrlMAGI.Value)
End Sub

Private Sub scrlNum_Change()
    lblNum.Caption = STR(scrlNum.Value)
    lblItemName.Caption = ""
    If scrlNum.Value > 0 Then
        lblItemName.Caption = Trim(Item(scrlNum.Value).Name)
    End If
    Npc(EditorIndex).ItemNPC(scrlDropItem.Value).ItemNum = scrlNum.Value
End Sub

Private Sub scrlValue_Change()
    Npc(EditorIndex).ItemNPC(scrlDropItem.Value).ItemValue = scrlValue.Value
    lblValue.Caption = STR(scrlValue.Value)
End Sub

Private Sub cmdOk_Click()
    Call NpcEditorOk
End Sub

Private Sub cmdCancel_Click()
    Call NpcEditorCancel
End Sub

Private Sub ScrlWaterDEF_Change()
    LblWaterDEF.Caption = STR(ScrlWaterDEF.Value)
End Sub

Private Sub ScrlWaterSTR_Change()
    LblWaterSTR.Caption = STR(ScrlWaterSTR.Value)
End Sub
Private Sub ScrlHeatDEF_Change()
    LblHeatDEF.Caption = STR(ScrlHeatDEF.Value)
End Sub

Private Sub ScrlHeatSTR_Change()
    LblHeatSTR.Caption = STR(ScrlHeatSTR.Value)
End Sub
Private Sub ScrlColdDEF_Change()
    LblColdDEF.Caption = STR(ScrlColdDEF.Value)
End Sub

Private Sub ScrlColdSTR_Change()
    LblColdSTR.Caption = STR(ScrlColdSTR.Value)
End Sub
Private Sub ScrlLightDEF_Change()
    LblLightDEF.Caption = STR(ScrlLightDEF.Value)
End Sub

Private Sub ScrlLightSTR_Change()
    LblLightSTR.Caption = STR(ScrlLightSTR.Value)
End Sub
Private Sub ScrlDarkDEF_Change()
    LblDarkDEF.Caption = STR(ScrlDarkDEF.Value)
End Sub

Private Sub ScrlDarkSTR_Change()
    LblDarkSTR.Caption = STR(ScrlDarkSTR.Value)
End Sub
Private Sub StartHP_Change()
    lblStartHP.Caption = StartHP.Value
End Sub

Private Sub tmrSprite_Timer()
    Call NpcEditorBltSprite
End Sub

Private Sub txtChance_Change()
    Npc(EditorIndex).ItemNPC(scrlDropItem.Value).Chance = Val(txtChance.Text)
End Sub

