VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmTradeSkills 
   BorderStyle     =   0  'None
   Caption         =   "Skills Menu"
   ClientHeight    =   7995
   ClientLeft      =   0
   ClientTop       =   -150
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTradeSkills.frx":0000
   ScaleHeight     =   7995
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   195
      Left            =   1320
      TabIndex        =   49
      Top             =   7800
      Width           =   2055
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   13573
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   353
      TabMaxWidth     =   1587
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Trade Skills"
      TabPicture(0)   =   "frmTradeSkills.frx":0333
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Timer1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Warrior Skills"
      TabPicture(1)   =   "frmTradeSkills.frx":034F
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Mage Skills"
      TabPicture(2)   =   "frmTradeSkills.frx":036B
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Player Skills"
      TabPicture(3)   =   "frmTradeSkills.frx":0387
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      Begin VB.Frame Frame6 
         Caption         =   "Warrior Skills"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4815
         Left            =   120
         TabIndex        =   21
         Top             =   2760
         Width           =   4620
         Begin VB.CommandButton btnclose 
            Caption         =   "Close Skills Panel"
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
            Left            =   1320
            TabIndex        =   48
            Top             =   3000
            Width           =   1695
         End
         Begin VB.Shape shpBows 
            BackColor       =   &H000000FF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H000000FF&
            Height          =   165
            Left            =   1450
            Top             =   4570
            Width           =   2370
         End
         Begin VB.Shape shpXbows 
            BackColor       =   &H000000FF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H000000FF&
            Height          =   165
            Left            =   1450
            Top             =   4210
            Width           =   2370
         End
         Begin VB.Shape shpThrown 
            BackColor       =   &H000000FF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H000000FF&
            Height          =   165
            Left            =   1450
            Top             =   3850
            Width           =   2370
         End
         Begin VB.Shape shpAxes 
            BackColor       =   &H000000FF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H000000FF&
            Height          =   165
            Left            =   1450
            Top             =   2530
            Width           =   2370
         End
         Begin VB.Shape shpPoles 
            BackColor       =   &H000000FF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H000000FF&
            Height          =   165
            Left            =   1450
            Top             =   2170
            Width           =   2370
         End
         Begin VB.Shape shpBluntWeapons 
            BackColor       =   &H000000FF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H000000FF&
            Height          =   165
            Left            =   1450
            Top             =   1810
            Width           =   2370
         End
         Begin VB.Shape shpSmallBlades 
            BackColor       =   &H000000FF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H000000FF&
            Height          =   165
            Left            =   1450
            Top             =   1450
            Width           =   2370
         End
         Begin VB.Shape shpLargeBlades 
            BackColor       =   &H000000FF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H000000FF&
            Height          =   165
            Left            =   1450
            Top             =   1090
            Width           =   2370
         End
         Begin VB.Label lblBows 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
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
            Height          =   180
            Left            =   1440
            TabIndex        =   47
            Top             =   4560
            Width           =   2385
         End
         Begin VB.Label lblXbows 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
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
            Height          =   180
            Left            =   1440
            TabIndex        =   46
            Top             =   4200
            Width           =   2385
         End
         Begin VB.Label lblThrown 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
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
            Height          =   180
            Left            =   1440
            TabIndex        =   45
            Top             =   3840
            Width           =   2385
         End
         Begin VB.Label lblAxes 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
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
            Height          =   180
            Left            =   1440
            TabIndex        =   44
            Top             =   2520
            Width           =   2385
         End
         Begin VB.Label lblPoles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
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
            Height          =   180
            Left            =   1440
            TabIndex        =   43
            Top             =   2160
            Width           =   2385
         End
         Begin VB.Label lblSmallBlades 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
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
            Height          =   180
            Left            =   1440
            TabIndex        =   42
            Top             =   1440
            Width           =   2385
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Ranged Weapons"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   225
            Left            =   1320
            TabIndex        =   41
            Top             =   3480
            Width           =   1785
         End
         Begin VB.Label lblKills 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Melee"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   225
            Left            =   1560
            TabIndex        =   40
            Top             =   720
            Width           =   1185
         End
         Begin VB.Label lblBowsLevel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Height          =   240
            Left            =   3960
            TabIndex        =   39
            Top             =   4560
            Width           =   120
         End
         Begin VB.Label lblXbowsLevel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Height          =   240
            Left            =   3960
            TabIndex        =   38
            Top             =   4200
            Width           =   120
         End
         Begin VB.Label Label19 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Bows:"
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
            Height          =   240
            Left            =   0
            TabIndex        =   37
            Top             =   4560
            Width           =   585
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Crossbows:"
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
            Height          =   240
            Left            =   0
            TabIndex        =   36
            Top             =   4200
            Width           =   1125
         End
         Begin VB.Label lblAxesLevel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Height          =   240
            Left            =   3960
            TabIndex        =   35
            Top             =   2520
            Width           =   120
         End
         Begin VB.Label lblPolesLevel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Height          =   240
            Left            =   3960
            TabIndex        =   34
            Top             =   2160
            Width           =   120
         End
         Begin VB.Label lblBluntWeaponsLevel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Height          =   240
            Left            =   3960
            TabIndex        =   33
            Top             =   1800
            Width           =   120
         End
         Begin VB.Label lblSmallBladesLevel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Height          =   240
            Left            =   3960
            TabIndex        =   32
            Top             =   1440
            Width           =   120
         End
         Begin VB.Label lblLargeBladesLevel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Height          =   240
            Left            =   3960
            TabIndex        =   31
            Top             =   1080
            Width           =   120
         End
         Begin VB.Label Label40 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Axes:"
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
            Height          =   240
            Left            =   0
            TabIndex        =   30
            Top             =   2520
            Width           =   555
         End
         Begin VB.Label Label29 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Polearms:"
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
            Height          =   240
            Left            =   0
            TabIndex        =   29
            Top             =   2160
            Width           =   960
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "B. Weapons:"
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
            Height          =   240
            Left            =   0
            TabIndex        =   28
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Small Blades:"
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
            Height          =   240
            Left            =   0
            TabIndex        =   27
            Top             =   1440
            Width           =   1260
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Large Blades:"
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
            Height          =   240
            Left            =   0
            TabIndex        =   26
            Top             =   1080
            Width           =   1320
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "T. Weapons:"
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
            Height          =   240
            Left            =   0
            TabIndex        =   25
            Top             =   3840
            Width           =   1200
         End
         Begin VB.Label lblThrownLevel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Height          =   240
            Left            =   3960
            TabIndex        =   24
            Top             =   3840
            Width           =   120
         End
         Begin VB.Label lblLargeBlades 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
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
            Height          =   180
            Left            =   1440
            TabIndex        =   23
            Top             =   1080
            Width           =   2385
         End
         Begin VB.Label lblBluntWeapons 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
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
            Height          =   180
            Left            =   1440
            TabIndex        =   22
            Top             =   1800
            Width           =   2385
         End
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   4080
         Top             =   240
      End
      Begin VB.Frame Frame1 
         Caption         =   "Wood Working Skills"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   4695
         Begin VB.Label lblJacking 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
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
            Height          =   180
            Left            =   1320
            TabIndex        =   13
            Top             =   600
            Width           =   2385
         End
         Begin VB.Label lblJackingLevel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Height          =   240
            Left            =   3840
            TabIndex        =   12
            Top             =   555
            Width           =   120
         End
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "L. Jacking:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   570
            Width           =   870
         End
         Begin VB.Label AddStr 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   165
            Left            =   0
            TabIndex        =   10
            Top             =   0
            Width           =   105
         End
         Begin VB.Shape shpJacking 
            BackColor       =   &H000000FF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H000000FF&
            Height          =   165
            Left            =   1330
            Top             =   610
            Width           =   2370
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Iron Skills"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   4695
         Begin VB.Label lblMine 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
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
            Height          =   180
            Left            =   1320
            TabIndex        =   8
            Top             =   360
            Width           =   2385
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Mining:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Width           =   600
         End
         Begin VB.Label lblMineLevel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Height          =   240
            Left            =   3840
            TabIndex        =   6
            Top             =   315
            Width           =   120
         End
         Begin VB.Shape shpMine 
            BackColor       =   &H000000FF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H000000FF&
            Height          =   165
            Left            =   1330
            Top             =   370
            Width           =   2370
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Fishing Skills"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   2040
         Width           =   4695
         Begin VB.Label lblFish 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
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
            Height          =   180
            Left            =   1320
            TabIndex        =   4
            Top             =   360
            Width           =   2385
         End
         Begin VB.Label Label21 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Fishing:"
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
            Height          =   240
            Left            =   120
            TabIndex        =   3
            Top             =   320
            Width           =   720
         End
         Begin VB.Label lblFishLevel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Height          =   240
            Left            =   3840
            TabIndex        =   2
            Top             =   315
            Width           =   120
         End
         Begin VB.Shape shpFish 
            BackColor       =   &H000000FF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H000000FF&
            Height          =   165
            Left            =   1330
            Top             =   370
            Width           =   2370
         End
      End
      Begin VB.Label lblRegister 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   -74880
         TabIndex        =   20
         Top             =   5280
         Width           =   1860
      End
      Begin VB.Label Label34 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Graveyard Studios"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   -72600
         TabIndex        =   19
         Top             =   4800
         Width           =   2085
      End
      Begin VB.Label lblLevel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Player Level"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   -74760
         TabIndex        =   18
         Top             =   3480
         Width           =   2835
      End
      Begin VB.Label lblUserName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Player Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   -74760
         TabIndex        =   17
         Top             =   3240
         Width           =   2835
      End
      Begin VB.Label Label36 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Visions"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   -73800
         TabIndex        =   16
         Top             =   2880
         Width           =   810
      End
      Begin VB.Label Label37 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Server:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   -74760
         TabIndex        =   15
         Top             =   2880
         Width           =   840
      End
      Begin VB.Label Label38 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright 2006"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   270
         Left            =   -72240
         TabIndex        =   14
         Top             =   4560
         Width           =   1755
      End
   End
End
Attribute VB_Name = "frmTradeSkills"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AddAlchemy_Click()
Call SendData("useexppool" & SEP_CHAR & 12 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddDoubleAttackExp_Click()
Call SendData("useexppool" & SEP_CHAR & 41 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddQuesting_Click()
Call SendData("useexppool" & SEP_CHAR & 39 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddFirstAid_Click()
Call SendData("useexppool" & SEP_CHAR & 40 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddAxes_Click()
Call SendData("useexppool" & SEP_CHAR & 18 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddBlunts_Click()
Call SendData("useexppool" & SEP_CHAR & 16 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddBody_Click()
Call SendData("useexppool" & SEP_CHAR & 8 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddBows_Click()
Call SendData("useexppool" & SEP_CHAR & 22 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddCarp_Click()
Call SendData("useexppool" & SEP_CHAR & 23 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddCasting_Click()
Call SendData("useexppool" & SEP_CHAR & 7 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddCombat_Click()
Call SendData("useexppool" & SEP_CHAR & 13 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddFishing_Click()
Call SendData("useexppool" & SEP_CHAR & 32 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddForaging_Click()
Call SendData("useexppool" & SEP_CHAR & 35 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddForging_Click()
Call SendData("useexppool" & SEP_CHAR & 28 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddHarvesting_Click()
Call SendData("useexppool" & SEP_CHAR & 34 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddLB_Click()
Call SendData("useexppool" & SEP_CHAR & 14 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddLjack_Click()
Call SendData("useexppool" & SEP_CHAR & 24 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddLW_Click()
Call SendData("useexppool" & SEP_CHAR & 36 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddMilling_Click()
Call SendData("useexppool" & SEP_CHAR & 25 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddMind_Click()
Call SendData("useexppool" & SEP_CHAR & 9 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddMining_Click()
Call SendData("useexppool" & SEP_CHAR & 26 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddNature_Click()
Call SendData("useexppool" & SEP_CHAR & 11 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddPK_Click()
Call SendData("useexppool" & SEP_CHAR & 5 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddPlanting_Click()
Call SendData("useexppool" & SEP_CHAR & 33 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddPoles_Click()
Call SendData("useexppool" & SEP_CHAR & 17 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddRep_Click()
Call SendData("useexppool" & SEP_CHAR & 4 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddRuneCraftExp_Click()
Call SendData("useexppool" & SEP_CHAR & 42 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddSB_Click()
Call SendData("useexppool" & SEP_CHAR & 15 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddSewing_Click()
Call SendData("useexppool" & SEP_CHAR & 31 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddSkinning_Click()
Call SendData("useexppool" & SEP_CHAR & 37 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddSmelting_Click()
Call SendData("useexppool" & SEP_CHAR & 27 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddSoul_Click()
Call SendData("useexppool" & SEP_CHAR & 10 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddSpinning_Click()
Call SendData("useexppool" & SEP_CHAR & 29 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddTanning_Click()
Call SendData("useexppool" & SEP_CHAR & 38 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddThief_Click()
Call SendData("useexppool" & SEP_CHAR & 6 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddThrown_Click()
Call SendData("useexppool" & SEP_CHAR & 20 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddUnarmed_Click()
Call SendData("useexppool" & SEP_CHAR & 19 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddWeaving_Click()
Call SendData("useexppool" & SEP_CHAR & 30 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddXbows_Click()
Call SendData("useexppool" & SEP_CHAR & 21 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddCritical_Click()
Call SendData("useexppool" & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddDodge_Click()
Call SendData("useexppool" & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddGovernment_Click()
Call SendData("useexppool" & SEP_CHAR & 3 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddLeadership_Click()
Call SendData("useexppool" & SEP_CHAR & 2 & SEP_CHAR & END_CHAR)
End Sub

Private Sub btnclose_Click()
frmTradeSkills.Visible = False
Timer1.Enabled = False
End Sub

Private Sub Command1_Click()
frmTradeSkills.Visible = False
End Sub

Private Sub Timer1_Timer()
Dim Index As Long

'If GetPlayerEPool(MyIndex) > 9999 Then
'frmTradeSkills.AddCritical.Visible = True
'frmTradeSkills.AddDodge.Visible = True
'frmTradeSkills.AddLeadership.Visible = True
'frmTradeSkills.AddGovernment.Visible = True
'frmTradeSkills.AddRep.Visible = True
'frmTradeSkills.AddPK.Visible = True
'frmTradeSkills.AddThief.Visible = True
'frmTradeSkills.AddCasting.Visible = True
'frmTradeSkills.AddBody.Visible = True
'frmTradeSkills.AddMind.Visible = True
'frmTradeSkills.AddSoul.Visible = True
'frmTradeSkills.AddNature.Visible = True
'frmTradeSkills.AddAlchemy.Visible = True
'frmTradeSkills.AddQuesting.Visible = True
'frmTradeSkills.AddFirstAid.Visible = True
'frmTradeSkills.AddLB.Visible = True
'frmTradeSkills.AddSB.Visible = True
'frmTradeSkills.AddBlunts.Visible = True
'frmTradeSkills.AddPoles.Visible = True
'frmTradeSkills.AddAxes.Visible = True
'frmTradeSkills.AddUnarmed.Visible = True
'frmTradeSkills.AddThrown.Visible = True
'frmTradeSkills.AddXbows.Visible = True
'frmTradeSkills.AddBows.Visible = True
'frmTradeSkills.AddCarp.Visible = True
'frmTradeSkills.AddLjack.Visible = True
'frmTradeSkills.AddMilling.Visible = True
'frmTradeSkills.AddCombat.Visible = True
'frmTradeSkills.AddMining.Visible = True
'frmTradeSkills.AddSmelting.Visible = True
'frmTradeSkills.AddForging.Visible = True
'frmTradeSkills.AddSpinning.Visible = True
'frmTradeSkills.AddWeaving.Visible = True
'frmTradeSkills.AddSewing.Visible = True
'frmTradeSkills.AddFishing.Visible = True
'frmTradeSkills.AddPlanting.Visible = True
'frmTradeSkills.AddHarvesting.Visible = True
'frmTradeSkills.AddForaging.Visible = True
'frmTradeSkills.AddLW.Visible = True
'frmTradeSkills.AddSkinning.Visible = True
'frmTradeSkills.AddTanning.Visible = True
'frmTradeSkills.AddDoubleAttackExp.Visible = True
'frmTradeSkills.AddRuneCraftExp.Visible = True
'Else
'frmTradeSkills.AddCritical.Visible = False
'frmTradeSkills.AddDodge.Visible = False
'frmTradeSkills.AddLeadership.Visible = False
'frmTradeSkills.AddGovernment.Visible = False
'frmTradeSkills.AddRep.Visible = False
'frmTradeSkills.AddPK.Visible = False
'frmTradeSkills.AddThief.Visible = False
'frmTradeSkills.AddCasting.Visible = False
'frmTradeSkills.AddBody.Visible = False
'frmTradeSkills.AddMind.Visible = False
'frmTradeSkills.AddSoul.Visible = False
'frmTradeSkills.AddNature.Visible = False
'frmTradeSkills.AddAlchemy.Visible = False
'frmTradeSkills.AddQuesting.Visible = False
'frmTradeSkills.AddFirstAid.Visible = False
'frmTradeSkills.AddDoubleAttackExp.Visible = False
'frmTradeSkills.AddLB.Visible = False
'frmTradeSkills.AddSB.Visible = False
'frmTradeSkills.AddBlunts.Visible = False
'frmTradeSkills.AddPoles.Visible = False
'frmTradeSkills.AddAxes.Visible = False
'frmTradeSkills.AddUnarmed.Visible = False
'frmTradeSkills.AddThrown.Visible = False
'frmTradeSkills.AddXbows.Visible = False
'frmTradeSkills.AddBows.Visible = False
'frmTradeSkills.AddMining.Visible = False
'frmTradeSkills.AddSmelting.Visible = False
'frmTradeSkills.AddForging.Visible = False
'frmTradeSkills.AddSpinning.Visible = False
'frmTradeSkills.AddWeaving.Visible = False
'frmTradeSkills.AddSewing.Visible = False
'frmTradeSkills.AddFishing.Visible = False
'frmTradeSkills.AddPlanting.Visible = False
'frmTradeSkills.AddHarvesting.Visible = False
'frmTradeSkills.AddForaging.Visible = False
'frmTradeSkills.AddLW.Visible = False
'frmTradeSkills.AddSkinning.Visible = False
'frmTradeSkills.AddTanning.Visible = False
'frmTradeSkills.AddCarp.Visible = False
'frmTradeSkills.AddLjack.Visible = False
'frmTradeSkills.AddMilling.Visible = False
'frmTradeSkills.AddCombat.Visible = False
'frmTradeSkills.AddRuneCraftExp.Visible = False
'End If
End Sub
