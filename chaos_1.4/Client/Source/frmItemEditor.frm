VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmItemEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Editor"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10575
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
   ScaleHeight     =   495
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox picSelect 
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
      Left            =   390
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   27
      Top             =   4590
      Width           =   480
   End
   Begin VB.Frame fraEquipment 
      Caption         =   "Equipment Data"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   3960
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   3135
      Begin VB.HScrollBar scrlMagicReq 
         Height          =   255
         Left            =   120
         Max             =   10000
         TabIndex        =   83
         Top             =   3480
         Width           =   2895
      End
      Begin VB.HScrollBar scrlAccessReq 
         Height          =   255
         Left            =   120
         Max             =   4
         TabIndex        =   41
         Top             =   4680
         Width           =   2895
      End
      Begin VB.HScrollBar scrlClassReq 
         Height          =   255
         Left            =   120
         Max             =   3
         TabIndex        =   40
         Top             =   4080
         Width           =   2895
      End
      Begin VB.HScrollBar scrlSpeedReq 
         Height          =   255
         Left            =   120
         Max             =   10000
         TabIndex        =   30
         Top             =   2880
         Width           =   2895
      End
      Begin VB.HScrollBar scrlDefReq 
         Height          =   255
         Left            =   120
         Max             =   10000
         TabIndex        =   29
         Top             =   2280
         Width           =   2895
      End
      Begin VB.HScrollBar scrlStrReq 
         Height          =   255
         Left            =   120
         Max             =   10000
         TabIndex        =   28
         Top             =   1680
         Width           =   2895
      End
      Begin VB.HScrollBar scrlStrength 
         Height          =   255
         Left            =   120
         Max             =   10000
         TabIndex        =   7
         Top             =   1080
         Value           =   1
         Width           =   2895
      End
      Begin VB.HScrollBar scrlDurability 
         Height          =   255
         Left            =   120
         Max             =   10000
         TabIndex        =   5
         Top             =   480
         Width           =   2895
      End
      Begin VB.CheckBox chkRepair 
         Caption         =   "Repairable?"
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
         Left            =   1920
         TabIndex        =   82
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label30 
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
         Height          =   375
         Left            =   960
         TabIndex        =   85
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         Caption         =   "Magic Req :"
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
         TabIndex        =   84
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "0 - Anyone"
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
         Left            =   1080
         TabIndex        =   43
         Top             =   4440
         Width           =   1695
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   165
         Left            =   1080
         TabIndex        =   42
         Top             =   3840
         Width           =   330
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Access Req :"
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
         TabIndex        =   39
         Top             =   4440
         Width           =   975
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Class Req :"
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
         TabIndex        =   38
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Speed Req :"
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
         TabIndex        =   37
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label13 
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
         Height          =   375
         Left            =   960
         TabIndex        =   35
         Top             =   2640
         Width           =   495
      End
      Begin VB.Label Label12 
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
         Height          =   375
         Left            =   960
         TabIndex        =   34
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label11 
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
         Height          =   375
         Left            =   960
         TabIndex        =   33
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Defence Req :"
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
         TabIndex        =   32
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Strength Req :"
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
         TabIndex        =   31
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblStrength 
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
         Height          =   375
         Left            =   960
         TabIndex        =   8
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblDurability 
         Alignment       =   1  'Right Justify
         Caption         =   "Ind."
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
         Left            =   960
         TabIndex        =   6
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Damage :"
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
         TabIndex        =   4
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Durability :"
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
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.PictureBox picPic 
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
      Height          =   2520
      Left            =   360
      ScaleHeight     =   168
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   192
      TabIndex        =   15
      Top             =   1800
      Width           =   2880
      Begin VB.PictureBox picItems 
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
         Left            =   0
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   192
         TabIndex        =   25
         Top             =   0
         Width           =   2880
      End
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
      Left            =   1320
      TabIndex        =   14
      Top             =   480
      Width           =   975
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
      Left            =   2400
      TabIndex        =   13
      Top             =   480
      Width           =   855
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
      Height          =   270
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   3135
   End
   Begin VB.ComboBox cmbType 
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
      ItemData        =   "frmItemEditor.frx":0000
      Left            =   360
      List            =   "frmItemEditor.frx":0055
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Frame fraVitals 
      Caption         =   "Vitals Data"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   3960
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   3135
      Begin VB.HScrollBar scrlVitalMod 
         Height          =   255
         Left            =   240
         Max             =   255
         TabIndex        =   11
         Top             =   1080
         Value           =   1
         Width           =   2655
      End
      Begin VB.Label lblVitalMod 
         Alignment       =   1  'Right Justify
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
         Height          =   375
         Left            =   840
         TabIndex        =   12
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Vital Mod :"
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
         TabIndex        =   10
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.Frame fraSpell 
      Caption         =   "Spell Data"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   3960
      TabIndex        =   16
      Top             =   600
      Visible         =   0   'False
      Width           =   3135
      Begin VB.HScrollBar scrlSpell 
         Height          =   255
         Left            =   240
         Max             =   255
         Min             =   1
         TabIndex        =   17
         Top             =   1200
         Value           =   1
         Width           =   2775
      End
      Begin VB.Label lblSpellName 
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
         TabIndex        =   21
         Top             =   600
         Width           =   2760
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Spell Name :"
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
         TabIndex        =   20
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Spell Number :"
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
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblSpell 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   1200
         TabIndex        =   18
         Top             =   960
         Width           =   495
      End
   End
   Begin VB.Frame fraPet 
      Caption         =   "Pet Data"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   3960
      TabIndex        =   86
      Top             =   600
      Visible         =   0   'False
      Width           =   3135
      Begin VB.HScrollBar scrlPetLevel 
         Height          =   255
         Left            =   240
         Max             =   255
         Min             =   1
         TabIndex        =   91
         Top             =   1920
         Value           =   1
         Width           =   2655
      End
      Begin VB.HScrollBar scrlPet 
         Height          =   255
         Left            =   240
         Max             =   255
         Min             =   1
         TabIndex        =   87
         Top             =   840
         Value           =   1
         Width           =   2655
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   1200
         TabIndex        =   93
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         Caption         =   "Level :"
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
         TabIndex        =   92
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   1200
         TabIndex        =   90
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         Caption         =   "Sprite Number :"
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
         TabIndex        =   89
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label31 
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
         TabIndex        =   88
         Top             =   600
         Width           =   2760
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7305
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   12885
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   397
      TabMaxWidth     =   1984
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Edit Item"
      TabPicture(0)   =   "frmItemEditor.frx":016F
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label26"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label36"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label39"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblElement"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblStamRemove"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label42"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Picture1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "VScroll1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "fraAttributes"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtDesc"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "fraBow"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtPrice"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "fraScript"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "chkBound"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "chkStackable"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "scrlElement"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "scrlStamRemove"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "fraRarity"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Command1"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "chkLB"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "chkSB"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "chkBW"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "chkPoles"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "chkAxes"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "chkThrown"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "chkBA"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "chkXbow"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Command2"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "FrameSkillsReq"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "chkDisease"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "chkPoison"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).ControlCount=   33
      Begin VB.CheckBox chkPoison 
         Caption         =   "Poison"
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
         Left            =   8400
         TabIndex        =   159
         Top             =   6600
         Width           =   855
      End
      Begin VB.CheckBox chkDisease 
         Caption         =   "Disease"
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
         Left            =   9360
         TabIndex        =   158
         Top             =   6600
         Width           =   855
      End
      Begin VB.Frame FrameSkillsReq 
         Caption         =   "Skill Reqs"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5895
         Left            =   7080
         TabIndex        =   129
         Top             =   120
         Visible         =   0   'False
         Width           =   3255
         Begin VB.HScrollBar scrlLargeBladesReq 
            Height          =   270
            Left            =   120
            Max             =   110
            TabIndex        =   139
            Top             =   600
            Value           =   1
            Width           =   2055
         End
         Begin VB.HScrollBar scrlSmallBladesReq 
            Height          =   270
            Left            =   120
            Max             =   110
            TabIndex        =   138
            Top             =   1200
            Value           =   1
            Width           =   2055
         End
         Begin VB.HScrollBar scrlBluntWeaponsReq 
            Height          =   270
            Left            =   120
            Max             =   110
            TabIndex        =   137
            Top             =   1800
            Value           =   1
            Width           =   2055
         End
         Begin VB.HScrollBar scrlPoleArmsReq 
            Height          =   270
            Left            =   120
            Max             =   110
            TabIndex        =   136
            Top             =   2400
            Value           =   1
            Width           =   2055
         End
         Begin VB.HScrollBar scrlAxesReq 
            Height          =   270
            Left            =   120
            Max             =   110
            TabIndex        =   135
            Top             =   3000
            Value           =   1
            Width           =   2055
         End
         Begin VB.HScrollBar scrlThrownReq 
            Height          =   270
            Left            =   120
            Max             =   110
            TabIndex        =   134
            Top             =   3600
            Value           =   1
            Width           =   2055
         End
         Begin VB.HScrollBar scrlXbowsReq 
            Height          =   270
            Left            =   120
            Max             =   110
            TabIndex        =   133
            Top             =   4200
            Value           =   1
            Width           =   2055
         End
         Begin VB.HScrollBar scrlBowsReq 
            Height          =   270
            Left            =   120
            Max             =   110
            TabIndex        =   132
            Top             =   4800
            Value           =   1
            Width           =   2055
         End
         Begin VB.HScrollBar scrlLevelReq 
            Height          =   270
            Left            =   120
            Max             =   110
            TabIndex        =   131
            Top             =   5400
            Value           =   1
            Width           =   2055
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Save"
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
            Left            =   2400
            TabIndex        =   130
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label44 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LB Required:"
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
            Left            =   120
            TabIndex        =   157
            Top             =   360
            Width           =   780
         End
         Begin VB.Label lblLargeBladesReq 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "God's Only Item"
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
            Left            =   1320
            TabIndex        =   156
            Top             =   360
            Width           =   1050
         End
         Begin VB.Label Label45 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SB Required:"
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
            Left            =   120
            TabIndex        =   155
            Top             =   960
            Width           =   810
         End
         Begin VB.Label lblSmallBladesReq 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "God's Only Item"
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
            Left            =   1320
            TabIndex        =   154
            Top             =   960
            Width           =   1050
         End
         Begin VB.Label Label46 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BW Required:"
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
            Left            =   120
            TabIndex        =   153
            Top             =   1560
            Width           =   840
         End
         Begin VB.Label lblBluntWeaponsReq 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "God's Only Item"
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
            Left            =   1320
            TabIndex        =   152
            Top             =   1560
            Width           =   1050
         End
         Begin VB.Label Label47 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Poles Required:"
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
            Left            =   120
            TabIndex        =   151
            Top             =   2160
            Width           =   960
         End
         Begin VB.Label lblPoleArmsReq 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "God's Only Item"
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
            Left            =   1320
            TabIndex        =   150
            Top             =   2160
            Width           =   1050
         End
         Begin VB.Label Label48 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Axes Required:"
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
            Left            =   120
            TabIndex        =   149
            Top             =   2760
            Width           =   930
         End
         Begin VB.Label lblAxesReq 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "God's Only Item"
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
            Left            =   1320
            TabIndex        =   148
            Top             =   2760
            Width           =   1050
         End
         Begin VB.Label Label49 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Thrown Required:"
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
            Left            =   120
            TabIndex        =   147
            Top             =   3360
            Width           =   1110
         End
         Begin VB.Label lblThrownReq 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "God's Only Item"
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
            Left            =   1320
            TabIndex        =   146
            Top             =   3360
            Width           =   1050
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Xbows Required:"
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
            Left            =   120
            TabIndex        =   145
            Top             =   3960
            Width           =   1035
         End
         Begin VB.Label lblXbowsReq 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "God's Only Item"
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
            Left            =   1320
            TabIndex        =   144
            Top             =   3960
            Width           =   1050
         End
         Begin VB.Label Label43 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "B. Level Required:"
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
            Left            =   120
            TabIndex        =   143
            Top             =   4560
            Width           =   1140
         End
         Begin VB.Label lblBowsReq 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "God's Only Item"
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
            Left            =   1320
            TabIndex        =   142
            Top             =   4560
            Width           =   1050
         End
         Begin VB.Label Label37 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Level Required:"
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
            Left            =   240
            TabIndex        =   141
            Top             =   5160
            Width           =   960
         End
         Begin VB.Label lblLevelReq 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "God's Only Item"
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
            Left            =   1320
            TabIndex        =   140
            Top             =   5160
            Width           =   1050
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Skill Reqs"
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
         Left            =   3600
         TabIndex        =   128
         Top             =   6480
         Width           =   855
      End
      Begin VB.CheckBox chkXbow 
         Caption         =   "Xbows"
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
         Left            =   6480
         TabIndex        =   127
         Top             =   6960
         Width           =   855
      End
      Begin VB.CheckBox chkBA 
         Caption         =   "Bows -"
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
         Left            =   7320
         TabIndex        =   126
         Top             =   6960
         Width           =   855
      End
      Begin VB.CheckBox chkThrown 
         Caption         =   "Thrown"
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
         Left            =   5520
         TabIndex        =   125
         Top             =   6960
         Width           =   855
      End
      Begin VB.CheckBox chkAxes 
         Caption         =   "Axes"
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
         Left            =   4800
         TabIndex        =   124
         Top             =   6960
         Width           =   735
      End
      Begin VB.CheckBox chkPoles 
         Caption         =   "Poles"
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
         Left            =   4080
         TabIndex        =   123
         Top             =   6960
         Width           =   735
      End
      Begin VB.CheckBox chkBW 
         Caption         =   "Blunt Weapons"
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
         TabIndex        =   122
         Top             =   6960
         Width           =   1455
      End
      Begin VB.CheckBox chkSB 
         Caption         =   "Small Blades"
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
         Left            =   1440
         TabIndex        =   121
         Top             =   6960
         Width           =   1215
      End
      Begin VB.CheckBox chkLB 
         Caption         =   "Large Blades"
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
         Left            =   240
         TabIndex        =   120
         Top             =   6960
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Item Rarity"
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
         Left            =   3480
         TabIndex        =   119
         Top             =   0
         Width           =   855
      End
      Begin VB.Frame fraRarity 
         Caption         =   "Item Rarity"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   7080
         TabIndex        =   109
         Top             =   360
         Visible         =   0   'False
         Width           =   3135
         Begin VB.CommandButton cmdRarNormal 
            Caption         =   "Normal"
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
            TabIndex        =   116
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton cmdRarCommon 
            Caption         =   "Common"
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
            Left            =   1080
            TabIndex        =   115
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton cmdRarRare 
            Caption         =   "Rare"
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
            Left            =   2040
            TabIndex        =   114
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton cmdRarEpic 
            Caption         =   "Epic"
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
            TabIndex        =   113
            Top             =   600
            Width           =   975
         End
         Begin VB.CommandButton cmdRarLegendary 
            Caption         =   "Legendary"
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
            Left            =   1080
            TabIndex        =   112
            Top             =   600
            Width           =   975
         End
         Begin VB.CommandButton cmdRarArtifact 
            Caption         =   "Artifact"
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
            Left            =   2040
            TabIndex        =   111
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox txtColor 
            Height          =   390
            Left            =   1320
            TabIndex        =   110
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label lblSelCol 
            Caption         =   "Selected Rarity:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Left            =   120
            TabIndex        =   117
            Top             =   1080
            Width           =   1095
         End
      End
      Begin VB.HScrollBar scrlStamRemove 
         Height          =   135
         Left            =   4680
         Max             =   500
         TabIndex        =   105
         Top             =   5640
         Width           =   1935
      End
      Begin VB.HScrollBar scrlElement 
         Height          =   255
         Left            =   3600
         Max             =   1000
         TabIndex        =   102
         Top             =   6000
         Value           =   1
         Width           =   2055
      End
      Begin VB.CheckBox chkStackable 
         Caption         =   "Stackable"
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
         Left            =   5880
         TabIndex        =   101
         Top             =   5880
         Width           =   1095
      End
      Begin VB.CheckBox chkBound 
         Caption         =   "Bound"
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
         Left            =   5880
         TabIndex        =   100
         Top             =   6120
         Width           =   855
      End
      Begin VB.Frame fraScript 
         Caption         =   "Scripted Item"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   96
         Top             =   6000
         Visible         =   0   'False
         Width           =   3255
         Begin VB.HScrollBar scrlScript 
            Height          =   255
            Left            =   240
            Max             =   255
            TabIndex        =   97
            Top             =   480
            Width           =   2775
         End
         Begin VB.Label Label38 
            Alignment       =   1  'Right Justify
            Caption         =   "Script Number"
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
            TabIndex        =   99
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblScript 
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
            Left            =   1200
            TabIndex        =   98
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.TextBox txtPrice 
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
         Left            =   7080
         TabIndex        =   94
         Top             =   6240
         Width           =   3135
      End
      Begin VB.Frame fraBow 
         Caption         =   "Bow"
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
         Left            =   840
         TabIndex        =   71
         Top             =   4320
         Visible         =   0   'False
         Width           =   2535
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   540
            Left            =   120
            ScaleHeight     =   34
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   34
            TabIndex        =   74
            Top             =   960
            Width           =   540
            Begin VB.PictureBox Picture3 
               BackColor       =   &H00404040&
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
               Left            =   15
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   75
               Top             =   15
               Width           =   480
               Begin VB.PictureBox picBow 
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
                  Left            =   -960
                  ScaleHeight     =   32
                  ScaleMode       =   3  'Pixel
                  ScaleWidth      =   128
                  TabIndex        =   76
                  Top             =   0
                  Width           =   1920
               End
            End
         End
         Begin VB.ComboBox cmbBow 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmItemEditor.frx":018B
            Left            =   120
            List            =   "frmItemEditor.frx":018D
            Style           =   2  'Dropdown List
            TabIndex        =   73
            Top             =   600
            Width           =   2295
         End
         Begin VB.CheckBox chkBow 
            Caption         =   "Bow"
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
            Left            =   120
            TabIndex        =   72
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblName 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   720
            TabIndex        =   78
            Top             =   1150
            Width           =   1665
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Height          =   165
            Left            =   720
            TabIndex        =   77
            Top             =   960
            Width           =   465
         End
      End
      Begin VB.TextBox txtDesc 
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
         Left            =   7080
         MaxLength       =   150
         TabIndex        =   69
         Top             =   5640
         Width           =   3135
      End
      Begin VB.Frame fraAttributes 
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
         Height          =   4935
         Left            =   7080
         TabIndex        =   44
         Top             =   360
         Visible         =   0   'False
         Width           =   3135
         Begin VB.HScrollBar scrlAttackSpeed 
            Height          =   230
            Left            =   360
            Max             =   5000
            Min             =   1
            TabIndex        =   79
            Top             =   4440
            Value           =   1000
            Width           =   2655
         End
         Begin VB.HScrollBar scrlAddEXP 
            Height          =   230
            Left            =   360
            Max             =   100
            TabIndex        =   66
            Top             =   3960
            Width           =   2655
         End
         Begin VB.HScrollBar scrlAddSP 
            Height          =   230
            Left            =   360
            Max             =   10000
            TabIndex        =   64
            Top             =   1560
            Width           =   2655
         End
         Begin VB.HScrollBar scrlAddSpeed 
            Height          =   230
            Left            =   360
            Max             =   10000
            TabIndex        =   56
            Top             =   3480
            Width           =   2655
         End
         Begin VB.HScrollBar scrlAddMagi 
            Height          =   230
            Left            =   360
            Max             =   10000
            TabIndex        =   55
            Top             =   3000
            Width           =   2655
         End
         Begin VB.HScrollBar scrlAddDef 
            Height          =   230
            Left            =   360
            Max             =   10000
            TabIndex        =   54
            Top             =   2520
            Width           =   2655
         End
         Begin VB.HScrollBar scrlAddStr 
            Height          =   230
            Left            =   360
            Max             =   10000
            TabIndex        =   53
            Top             =   2040
            Width           =   2655
         End
         Begin VB.HScrollBar scrlAddMP 
            Height          =   230
            Left            =   360
            Max             =   10000
            TabIndex        =   52
            Top             =   1080
            Width           =   2655
         End
         Begin VB.HScrollBar scrlAddHP 
            Height          =   230
            Left            =   360
            Max             =   10000
            TabIndex        =   51
            Top             =   600
            Width           =   2655
         End
         Begin VB.Label lblAttackSpeed 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1000 Milleseconds"
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
            Left            =   1200
            TabIndex        =   81
            Top             =   4200
            Width           =   1110
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Attack Speed :"
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
            TabIndex        =   80
            Top             =   4200
            Width           =   975
         End
         Begin VB.Label lblAddEXP 
            Alignment       =   1  'Right Justify
            Caption         =   "0%"
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
            Left            =   1080
            TabIndex        =   68
            Top             =   3720
            Width           =   495
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Add EXP :"
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
            TabIndex        =   67
            Top             =   3720
            Width           =   855
         End
         Begin VB.Label lblAddSP 
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
            Height          =   375
            Left            =   1080
            TabIndex        =   65
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            Caption         =   "Add SP :"
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
            TabIndex        =   63
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label lblAddSpeed 
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
            Height          =   375
            Left            =   1080
            TabIndex        =   62
            Top             =   3240
            Width           =   495
         End
         Begin VB.Label lblAddMagi 
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
            Height          =   375
            Left            =   1080
            TabIndex        =   61
            Top             =   2760
            Width           =   495
         End
         Begin VB.Label lblAddDef 
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
            Height          =   375
            Left            =   1080
            TabIndex        =   60
            Top             =   2280
            Width           =   495
         End
         Begin VB.Label lblAddStr 
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
            Height          =   375
            Left            =   1080
            TabIndex        =   59
            Top             =   1800
            Width           =   495
         End
         Begin VB.Label lblAddMP 
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
            Height          =   375
            Left            =   1080
            TabIndex        =   58
            Top             =   840
            Width           =   495
         End
         Begin VB.Label lblAddHP 
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
            Height          =   375
            Left            =   1080
            TabIndex        =   57
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Add Speed :"
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
            TabIndex        =   50
            Top             =   3240
            Width           =   855
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            Caption         =   "Add Magi :"
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
            TabIndex        =   49
            Top             =   2760
            Width           =   735
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            Caption         =   "Add Def :"
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
            TabIndex        =   48
            Top             =   2280
            Width           =   735
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            Caption         =   "Add Str :"
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
            TabIndex        =   47
            Top             =   1800
            Width           =   735
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            Caption         =   "Add MP :"
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
            TabIndex        =   46
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "Add HP :"
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
            TabIndex        =   45
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   2520
         Left            =   3120
         Max             =   464
         TabIndex        =   26
         Top             =   1680
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   240
         ScaleHeight     =   510
         ScaleWidth      =   510
         TabIndex        =   36
         Top             =   4440
         Width           =   540
      End
      Begin VB.Label Label42 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Stamina Cost:"
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
         Left            =   4110
         TabIndex        =   118
         Top             =   5400
         Width           =   885
      End
      Begin VB.Label lblStamRemove 
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
         Left            =   6360
         TabIndex        =   106
         Top             =   5400
         Width           =   495
      End
      Begin VB.Label lblElement 
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
         Height          =   165
         Left            =   4200
         TabIndex        =   104
         Top             =   5760
         Width           =   1410
      End
      Begin VB.Label Label39 
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
         Height          =   165
         Left            =   3600
         TabIndex        =   103
         Top             =   5760
         Width           =   555
      End
      Begin VB.Label Label36 
         Caption         =   "Sell price:"
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
         Left            =   7200
         TabIndex        =   95
         Top             =   6000
         Width           =   735
      End
      Begin VB.Label Label26 
         Caption         =   "Description :"
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
         Left            =   7200
         TabIndex        =   70
         Top             =   5400
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Item Name :"
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
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Item Sprite :"
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
         TabIndex        =   23
         Top             =   1440
         Width           =   855
      End
   End
   Begin VB.Label Label41 
      Caption         =   "SP Remove :"
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
      Left            =   4800
      TabIndex        =   108
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label Label40 
      Caption         =   "Item Name :"
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
      TabIndex        =   107
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmItemEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Copyright (c) 2006 Chaos Engine Source. All rights reserved.
' This code is licensed under the Chaos Engine General License.

Option Explicit

Private Sub chkBow_Click()
Dim I As Long
    If chkBow.Value = Unchecked Then
        cmbBow.Clear
        cmbBow.AddItem "None", 0
        cmbBow.ListIndex = 0
        cmbBow.Enabled = False
        lblName.Caption = ""
    Else
        cmbBow.Clear
        For I = 1 To MAX_ARROWS
            cmbBow.AddItem I & ": " & Arrows(I).Name
        Next I
        cmbBow.ListIndex = 0
        cmbBow.Enabled = True
    End If
End Sub

Private Sub cmbBow_Click()
    lblName.Caption = Arrows(cmbBow.ListIndex + 1).Name
    picBow.Top = (Arrows(cmbBow.ListIndex + 1).pic * 32) * -1
End Sub

Private Sub cmdItemRarity_Click()
If fraRarity.Visible = False Then
fraRarity.Visible = True
Else: fraRarity.Visible = False
End If
End Sub

Private Sub cmdOk_Click()
    Call ItemEditorOk
End Sub

Private Sub cmdCancel_Click()
    Call ItemEditorCancel
End Sub

Private Sub cmbType_Click()
    If (cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (cmbType.ListIndex <= ITEM_TYPE_SHIELD) Or (cmbType.ListIndex >= ITEM_TYPE_LEGS) Or (cmbType.ListIndex <= ITEM_TYPE_BOOTS) Or (cmbType.ListIndex <= ITEM_TYPE_GLOVES) Or (cmbType.ListIndex <= ITEM_TYPE_RING1) Or (cmbType.ListIndex <= ITEM_TYPE_RING2) Or (cmbType.ListIndex <= ITEM_TYPE_AMULET) Then
        If cmbType.ListIndex = ITEM_TYPE_WEAPON Then
            Label3.Caption = "Damage :"
        Else
            Label3.Caption = "Defence :"
        End If
        fraEquipment.Visible = True
        fraPet.Visible = False
        fraAttributes.Visible = True
        fraBow.Visible = True
    Else
        fraEquipment.Visible = False
        fraAttributes.Visible = False
        fraBow.Visible = False
    End If
    
        
    If (cmbType.ListIndex >= ITEM_TYPE_POTIONADDHP) And (cmbType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Then
        fraVitals.Visible = True
        fraPet.Visible = False
        fraAttributes.Visible = False
        fraEquipment.Visible = False
        fraBow.Visible = False
    Else
        fraVitals.Visible = False
    End If
    
    If (cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        fraSpell.Visible = True
        fraPet.Visible = False
        fraAttributes.Visible = False
        fraEquipment.Visible = False
        fraBow.Visible = False
    Else
        fraSpell.Visible = False
    End If
    
    If (cmbType.ListIndex = ITEM_TYPE_PET) Then
        fraSpell.Visible = False
        fraPet.Visible = True
        fraAttributes.Visible = False
        fraEquipment.Visible = False
        fraBow.Visible = False
    Else
        fraPet.Visible = False
    End If
    
    If (cmbType.ListIndex = ITEM_TYPE_SCRIPTED) Then
       fraScript.Visible = True
       fraAttributes.Visible = False
       fraEquipment.Visible = False
       fraBow.Visible = False
       fraSpell.Visible = False
    Else
       fraScript.Visible = False
End If

    If (cmbType.ListIndex = ITEM_TYPE_GUILDDEED) Then
       fraScript.Visible = False
       fraAttributes.Visible = False
       fraEquipment.Visible = False
       fraBow.Visible = False
       fraSpell.Visible = False
    Else
       fraScript.Visible = False
End If
End Sub

Private Sub Command1_Click()
If fraRarity.Visible = False Then
fraRarity.Visible = True
Else
fraRarity.Visible = False
End If
End Sub

Private Sub Command2_Click()
FrameSkillsReq.Visible = True
End Sub

Private Sub Command3_Click()
FrameSkillsReq.Visible = False
End Sub

Private Sub form_load()
Dim sDc As Long
    scrlElement.Max = MAX_ELEMENTS
    picItems.Height = 320 * PIC_Y
    Call BitBlt(picSelect.hDC, 0, 0, PIC_X, PIC_Y, picItems.hDC, EditorItemX * PIC_X, EditorItemY * PIC_Y, SRCCOPY)

    sDc = DD_ArrowAnim.GetDC
    With picBow
        .Cls
        .Width = DDSD_ArrowAnim.lWidth
        .Height = DDSD_ArrowAnim.lHeight
        Call BitBlt(.hDC, 0, 0, .Width, .Height, sDc, 0, 0, SRCCOPY)
    End With
    Call DD_ArrowAnim.ReleaseDC(sDc)

End Sub

Private Sub picItems_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim sDc As Long

    If Button = 1 Then
        EditorItemX = Int(x / PIC_X)
        EditorItemY = Int(y / PIC_Y)
    End If
    sDc = DD_ItemSurf.GetDC
    Call BitBlt(picSelect.hDC, 0, 0, PIC_X, PIC_Y, sDc, EditorItemX * PIC_X, EditorItemY * PIC_Y, SRCCOPY)
    Call DD_ItemSurf.ReleaseDC(sDc)
End Sub

Private Sub picItems_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim sDc As Long

    If Button = 1 Then
        EditorItemX = Int(x / PIC_X)
        EditorItemY = Int(y / PIC_Y)
    End If
    sDc = DD_ItemSurf.GetDC
    Call BitBlt(picSelect.hDC, 0, 0, PIC_X, PIC_Y, sDc, EditorItemX * PIC_X, EditorItemY * PIC_Y, SRCCOPY)
    Call DD_ItemSurf.ReleaseDC(sDc)
End Sub

Private Sub scrlAccessReq_Change()
    With scrlAccessReq
        Select Case .Value
            Case 0
                Label17.Caption = "0 - Anyone"
            Case 1
                Label17.Caption = "1 - Moniters"
            Case 2
                Label17.Caption = "2 - Mappers"
            Case 3
                Label17.Caption = "3 - Developers"
            Case 4
                Label17.Caption = "4 - Admins"
            End Select
    End With
End Sub

Private Sub scrlAddDef_Change()
    lblAddDef.Caption = scrlAddDef.Value
End Sub

Private Sub scrlAddEXP_Change()
    lblAddEXP.Caption = scrlAddEXP.Value & "%"
End Sub

Private Sub scrlAddHP_Change()
    lblAddHP.Caption = scrlAddHP.Value
End Sub

Private Sub scrlAddMagi_Change()
    lblAddMagi.Caption = scrlAddMagi.Value
End Sub

Private Sub scrlAddMP_Change()
    lblAddMP.Caption = scrlAddMP.Value
End Sub

Private Sub scrlAddSP_Change()
    lblAddSP.Caption = scrlAddSP.Value
End Sub

Private Sub scrlAddSpeed_Change()
    lblAddSpeed.Caption = scrlAddSpeed.Value
End Sub

Private Sub scrlAddStr_Change()
    lblAddStr.Caption = scrlAddStr.Value
End Sub

Private Sub scrlAttackSpeed_Change()
    lblAttackSpeed.Caption = scrlAttackSpeed.Value & " Milleseconds"
End Sub

Private Sub scrlAxesReq_Change()
lblAxesReq.Caption = STR(scrlAxesReq.Value)
End Sub

Private Sub scrlBluntWeaponsReq_Change()
lblBluntWeaponsReq.Caption = STR(scrlBluntWeaponsReq.Value)
End Sub

Private Sub scrlBowsReq_Change()
lblBowsReq.Caption = STR(scrlBowsReq.Value)
End Sub

Private Sub scrlClassReq_Change()
If scrlClassReq.Value = 0 Then
    Label16.Caption = "None"
Else
    Label16.Caption = scrlClassReq.Value & " - " & Trim(Class(scrlClassReq.Value).Name)
End If
End Sub

Private Sub scrlDefReq_Change()
    Label12.Caption = scrlDefReq.Value
End Sub

Private Sub scrlElement_Change()
lblElement.Caption = Element(scrlElement.Value).Name
End Sub

Private Sub scrlLargeBladesReq_Change()
lblLargeBladesReq.Caption = STR(scrlLargeBladesReq.Value)
End Sub

Private Sub scrlLevelReq_Change()
  lblLevelReq.Caption = STR(scrlLevelReq.Value)
End Sub

Private Sub scrlMagicReq_Change()
    Label30.Caption = scrlMagicReq.Value
End Sub

Private Sub scrlPet_Change()
    Label34.Caption = scrlPet.Value
End Sub

Private Sub scrlPetLevel_Change()
    Label35.Caption = scrlPetLevel.Value
End Sub

Private Sub scrlPoleArmsReq_Change()
lblPoleArmsReq.Caption = STR(scrlPoleArmsReq.Value)
End Sub

Private Sub scrlSmallBladesReq_Change()
lblSmallBladesReq.Caption = STR(scrlSmallBladesReq.Value)
End Sub

Private Sub scrlSpeedReq_Change()
    Label13.Caption = scrlSpeedReq.Value
End Sub

Private Sub scrlStamRemove_Change()
lblStamRemove.Caption = scrlStamRemove.Value
End Sub

Private Sub scrlStrReq_Change()
    Label11.Caption = scrlStrReq.Value
End Sub

Private Sub scrlThrownReq_Change()
lblThrownReq.Caption = STR(scrlThrownReq.Value)
End Sub

Private Sub scrlVitalMod_Change()
    lblVitalMod.Caption = STR(scrlVitalMod.Value)
End Sub

Private Sub scrlDurability_Change()
    lblDurability.Caption = STR(scrlDurability.Value)
    If STR(scrlDurability.Value) = 0 Then
        lblDurability.Caption = "Ind."
    End If
End Sub

Private Sub scrlStrength_Change()
    lblStrength.Caption = STR(scrlStrength.Value)
End Sub

Private Sub scrlSpell_Change()
    lblSpellName.Caption = Trim(Spell(scrlSpell.Value).Name)
    lblSpell.Caption = STR(scrlSpell.Value)
End Sub

Private Sub scrlXbowsReq_Change()
lblXbowsReq.Caption = STR(scrlXbowsReq.Value)
End Sub

Private Sub Timer1_Timer()
    Call BitBlt(picSelect.hDC, 0, 0, PIC_X, PIC_Y, picItems.hDC, EditorItemX * PIC_X, EditorItemY * PIC_Y, SRCCOPY)
End Sub

Private Sub VScroll1_Change()
    picItems.Top = (VScroll1.Value * PIC_Y) * -1
End Sub

Private Sub scrlScript_Change()
lblScript.Caption = scrlScript.Value
End Sub

Private Sub cmdRarArtifact_Click()
txtColor.Text = &HFF&
lblSelCol.ForeColor = &HFF&
End Sub

Private Sub cmdRarCommon_Click()
txtColor.Text = &HFF00&
lblSelCol.ForeColor = &HFF00&
End Sub

Private Sub cmdRarEpic_Click()
txtColor.Text = &HFF0000
lblSelCol.ForeColor = &HFF0000
End Sub

Private Sub cmdRarLegendary_Click()
txtColor.Text = &HC000C0
lblSelCol.ForeColor = &HC000C0
End Sub

Private Sub cmdRarNormal_Click()
txtColor.Text = &HFFFFFF
lblSelCol.ForeColor = &HFFFFFF
End Sub

Private Sub cmdRarRare_Click()
txtColor.Text = &HFFFF&
lblSelCol.ForeColor = &HFFFF&
End Sub
