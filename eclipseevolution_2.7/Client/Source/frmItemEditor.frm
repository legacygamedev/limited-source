VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmItemEditor 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Editor"
   ClientHeight    =   7125
   ClientLeft      =   330
   ClientTop       =   510
   ClientWidth     =   10725
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
   Icon            =   "frmItemEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   475
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   715
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
      Left            =   510
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   26
      Top             =   4470
      Width           =   480
   End
   Begin VB.Frame fraEquipment 
      Caption         =   "Main Stats And Requirements"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   3840
      TabIndex        =   1
      Top             =   600
      Width           =   3135
      Begin VB.HScrollBar scrlMagicReq 
         Height          =   230
         Left            =   120
         Max             =   10000
         TabIndex        =   93
         Top             =   3480
         Width           =   2895
      End
      Begin VB.HScrollBar scrlAccessReq 
         Height          =   230
         Left            =   120
         Max             =   4
         TabIndex        =   40
         Top             =   5280
         Width           =   2895
      End
      Begin VB.HScrollBar scrlClassReq 
         Height          =   230
         Left            =   120
         Max             =   1
         Min             =   -1
         TabIndex        =   39
         Top             =   4080
         Value           =   -1
         Width           =   2895
      End
      Begin VB.HScrollBar scrlSpeedReq 
         Height          =   230
         Left            =   120
         Max             =   10000
         TabIndex        =   29
         Top             =   2880
         Width           =   2895
      End
      Begin VB.HScrollBar scrlDefReq 
         Height          =   230
         Left            =   120
         Max             =   10000
         TabIndex        =   28
         Top             =   2280
         Width           =   2895
      End
      Begin VB.HScrollBar scrlStrReq 
         Height          =   230
         Left            =   120
         Max             =   10000
         TabIndex        =   27
         Top             =   1680
         Width           =   2895
      End
      Begin VB.HScrollBar scrlStrength 
         Height          =   230
         Left            =   120
         Max             =   10000
         TabIndex        =   6
         Top             =   1080
         Value           =   1
         Width           =   2895
      End
      Begin VB.HScrollBar scrlDurability 
         Height          =   230
         Left            =   120
         Max             =   10000
         Min             =   -1
         TabIndex        =   4
         Top             =   480
         Value           =   -1
         Width           =   2895
      End
      Begin VB.Label Label33 
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
         Left            =   2520
         TabIndex        =   95
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Magic Required :"
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
         TabIndex        =   94
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
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
         Left            =   1440
         TabIndex        =   42
         Top             =   5040
         Width           =   1575
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
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
         Left            =   2640
         TabIndex        =   41
         Top             =   3840
         Width           =   330
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Minimum Access :"
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
         TabIndex        =   38
         Top             =   5040
         Width           =   1215
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Class Required :"
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
         TabIndex        =   37
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Speed Required :"
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
         TabIndex        =   36
         Top             =   2640
         Width           =   1215
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
         Left            =   2520
         TabIndex        =   34
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
         Left            =   2520
         TabIndex        =   33
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
         Left            =   2520
         TabIndex        =   32
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Defense Required :"
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
         TabIndex        =   31
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Strength Required :"
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
         TabIndex        =   30
         Top             =   1440
         Width           =   1215
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
         Left            =   2520
         TabIndex        =   7
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
         Left            =   2520
         TabIndex        =   5
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Damage Bonus :"
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
         TabIndex        =   3
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label2 
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
         Left            =   120
         TabIndex        =   2
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
      Height          =   2160
      Left            =   480
      ScaleHeight     =   144
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   192
      TabIndex        =   14
      Top             =   2040
      Width           =   2880
      Begin VB.PictureBox picItems 
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
         TabIndex        =   24
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   13
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Save Item"
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
      Left            =   8760
      TabIndex        =   12
      Top             =   6360
      Width           =   1575
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
      Left            =   480
      TabIndex        =   0
      Top             =   840
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
      Left            =   3840
      TabIndex        =   8
      Top             =   600
      Visible         =   0   'False
      Width           =   3135
      Begin VB.HScrollBar scrlVitalMod 
         Height          =   255
         Left            =   240
         Max             =   10000
         TabIndex        =   10
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
         TabIndex        =   11
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
         TabIndex        =   9
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
      Left            =   3840
      TabIndex        =   15
      Top             =   600
      Visible         =   0   'False
      Width           =   3135
      Begin VB.HScrollBar scrlSpell 
         Height          =   255
         Left            =   240
         Max             =   255
         Min             =   1
         TabIndex        =   16
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
         Top             =   960
         Width           =   495
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6945
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   12250
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
      TabPicture(0)   =   "frmItemEditor.frx":0FC2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label26"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label29"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Picture1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "VScroll1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fraAttributes"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtDesc"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "fraBow"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "HScroll1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtPrice"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "fraScript"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "chkStackable"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "chkBound"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmbType"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
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
         ItemData        =   "frmItemEditor.frx":0FDE
         Left            =   360
         List            =   "frmItemEditor.frx":101B
         Style           =   2  'Dropdown List
         TabIndex        =   91
         Top             =   1080
         Width           =   3135
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
         Left            =   2520
         TabIndex        =   90
         Top             =   6240
         Width           =   855
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
         Left            =   1200
         TabIndex        =   89
         Top             =   6240
         Width           =   1215
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
         Height          =   3375
         Left            =   3720
         TabIndex        =   84
         Top             =   480
         Visible         =   0   'False
         Width           =   3135
         Begin VB.HScrollBar scrlScript 
            Height          =   255
            Left            =   240
            Max             =   500
            TabIndex        =   85
            Top             =   1200
            Width           =   2775
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
            TabIndex        =   88
            Top             =   960
            Width           =   495
         End
         Begin VB.Label Label32 
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
            TabIndex        =   87
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label30 
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
            TabIndex        =   86
            Top             =   600
            Width           =   2760
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
         Left            =   4680
         TabIndex        =   83
         Top             =   6240
         Width           =   2175
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   360
         TabIndex        =   81
         Top             =   1680
         Width           =   2895
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
         Height          =   1935
         Left            =   960
         TabIndex        =   70
         Top             =   4200
         Visible         =   0   'False
         Width           =   2535
         Begin VB.CheckBox chkGrapple 
            Caption         =   "Grapplehook"
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
            Left            =   720
            TabIndex        =   92
            Top             =   240
            Width           =   1095
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   540
            Left            =   120
            ScaleHeight     =   34
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   34
            TabIndex        =   73
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
               TabIndex        =   74
               Top             =   15
               Width           =   480
               Begin VB.PictureBox picBow 
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
                  Left            =   -240
                  ScaleHeight     =   32
                  ScaleMode       =   3  'Pixel
                  ScaleWidth      =   128
                  TabIndex        =   75
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
            ItemData        =   "frmItemEditor.frx":10E7
            Left            =   120
            List            =   "frmItemEditor.frx":10E9
            Style           =   2  'Dropdown List
            TabIndex        =   72
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
            TabIndex        =   71
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
            TabIndex        =   77
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
            TabIndex        =   76
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
         Left            =   2520
         MaxLength       =   150
         TabIndex        =   68
         Top             =   6600
         Width           =   7695
      End
      Begin VB.Frame fraAttributes 
         Caption         =   "Item Bonuses"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5655
         Left            =   7080
         TabIndex        =   43
         Top             =   480
         Width           =   3135
         Begin VB.HScrollBar scrlAttackSpeed 
            Height          =   230
            Left            =   120
            Max             =   5000
            Min             =   1
            TabIndex        =   78
            Top             =   5280
            Value           =   1000
            Width           =   2895
         End
         Begin VB.HScrollBar scrlAddEXP 
            Height          =   230
            Left            =   120
            Max             =   100
            TabIndex        =   65
            Top             =   4680
            Width           =   2895
         End
         Begin VB.HScrollBar scrlAddSP 
            Height          =   230
            Left            =   120
            Max             =   10000
            TabIndex        =   63
            Top             =   1680
            Width           =   2895
         End
         Begin VB.HScrollBar scrlAddSpeed 
            Height          =   230
            Left            =   120
            Max             =   10000
            TabIndex        =   55
            Top             =   4080
            Width           =   2895
         End
         Begin VB.HScrollBar scrlAddMagi 
            Height          =   230
            Left            =   120
            Max             =   10000
            TabIndex        =   54
            Top             =   3480
            Width           =   2895
         End
         Begin VB.HScrollBar scrlAddDef 
            Height          =   230
            Left            =   120
            Max             =   10000
            TabIndex        =   53
            Top             =   2880
            Width           =   2895
         End
         Begin VB.HScrollBar scrlAddStr 
            Height          =   230
            Left            =   120
            Max             =   10000
            TabIndex        =   52
            Top             =   2280
            Width           =   2895
         End
         Begin VB.HScrollBar scrlAddMP 
            Height          =   230
            Left            =   120
            Max             =   10000
            TabIndex        =   51
            Top             =   1080
            Width           =   2895
         End
         Begin VB.HScrollBar scrlAddHP 
            Height          =   230
            Left            =   120
            Max             =   10000
            TabIndex        =   50
            Top             =   480
            Width           =   2895
         End
         Begin VB.Label lblAttackSpeed 
            Alignment       =   1  'Right Justify
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
            Left            =   1920
            TabIndex        =   80
            Top             =   5040
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
            TabIndex        =   79
            Top             =   5040
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
            Left            =   2520
            TabIndex        =   67
            Top             =   4440
            Width           =   495
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "Experience Modifier :"
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
            TabIndex        =   66
            Top             =   4440
            Width           =   1575
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
            Left            =   2520
            TabIndex        =   64
            Top             =   1440
            Width           =   495
         End
         Begin VB.Label Label24 
            Caption         =   "Stamina Bonus:"
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
            TabIndex        =   62
            Top             =   1440
            Width           =   1215
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
            Left            =   2520
            TabIndex        =   61
            Top             =   3840
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
            Left            =   2520
            TabIndex        =   60
            Top             =   3240
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
            Left            =   2520
            TabIndex        =   59
            Top             =   2640
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
            Left            =   2520
            TabIndex        =   58
            Top             =   2040
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
            Left            =   2520
            TabIndex        =   57
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
            Left            =   2520
            TabIndex        =   56
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Bonus Speed :"
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
            TabIndex        =   49
            Top             =   3840
            Width           =   1455
         End
         Begin VB.Label Label22 
            Caption         =   "Bonus Magic :"
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
            TabIndex        =   48
            Top             =   3240
            Width           =   1455
         End
         Begin VB.Label Label21 
            Caption         =   "Bonus Defence :"
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
            TabIndex        =   47
            Top             =   2640
            Width           =   1335
         End
         Begin VB.Label Label20 
            Caption         =   "Bonus Strength :"
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
            TabIndex        =   46
            Top             =   2040
            Width           =   1335
         End
         Begin VB.Label Label19 
            Caption         =   "Mana Bonus :"
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
            TabIndex        =   45
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label18 
            Caption         =   "Health Bonus :"
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
            TabIndex        =   44
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   2400
         Left            =   3240
         Max             =   464
         TabIndex        =   25
         Top             =   1680
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   360
         ScaleHeight     =   510
         ScaleWidth      =   510
         TabIndex        =   35
         Top             =   4320
         Width           =   540
      End
      Begin VB.Label Label29 
         Caption         =   "Sell Price :"
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
         Left            =   3840
         TabIndex        =   82
         Top             =   6300
         Width           =   1455
      End
      Begin VB.Label Label26 
         Caption         =   "Item Description :"
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
         Left            =   960
         TabIndex        =   69
         Top             =   6645
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Item Name :"
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
         Left            =   360
         TabIndex        =   23
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
         Left            =   360
         TabIndex        =   22
         Top             =   1440
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmItemEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkBow_Click()
    Dim i As Long
    If chkBow.Value = Unchecked Then
        cmbBow.Clear
        cmbBow.addItem "None", 0
        cmbBow.ListIndex = 0
        cmbBow.Enabled = False
        lblName.Caption = vbNullString
    Else
        cmbBow.Clear
        For i = 1 To MAX_ARROWS
            cmbBow.addItem i & ": " & Arrows(i).Name
        Next i
        cmbBow.ListIndex = 0
        cmbBow.Enabled = True
    End If
End Sub

Private Sub cmbBow_Click()
    lblName.Caption = Arrows(cmbBow.ListIndex + 1).Name
    picBow.Top = (Arrows(frmItemEditor.cmbBow.ListIndex + 1).Pic * 32) * -1
    picBow.Left = 0
End Sub

Private Sub cmdOk_Click()
    Call ItemEditorOk
End Sub

Private Sub cmdCancel_Click()
    Call ItemEditorCancel
End Sub

Private Sub cmbType_Click()
    If (cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (cmbType.ListIndex <= ITEM_TYPE_NECKLACE) Then
        If cmbType.ListIndex = ITEM_TYPE_WEAPON Or cmbType.ListIndex = ITEM_TYPE_TWO_HAND Or cmbType.ListIndex = ITEM_TYPE_RING Or cmbType.ListIndex = ITEM_TYPE_NECKLACE Then
            Label3.Caption = "Damage :"
        Else
            Label3.Caption = STAT2 & " :"
        End If
        fraEquipment.Visible = True
        fraAttributes.Visible = True
        If cmbType.ListIndex = ITEM_TYPE_WEAPON Then
            fraBow.Visible = True
        End If
        chkStackable.Visible = False
    Else
        chkStackable.Visible = True
        fraEquipment.Visible = False
        fraAttributes.Visible = False
        fraBow.Visible = False
    End If

    If (cmbType.ListIndex >= ITEM_TYPE_POTIONADDHP) And (cmbType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Then
        fraVitals.Visible = True
        fraAttributes.Visible = False
        fraEquipment.Visible = False
        fraBow.Visible = False
    Else
        fraVitals.Visible = False
    End If

    If (cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        fraSpell.Visible = True
        fraAttributes.Visible = False
        fraEquipment.Visible = False
        fraBow.Visible = False
        chkStackable.Visible = False
    Else
        fraSpell.Visible = False
    End If

    If (cmbType.ListIndex = ITEM_TYPE_SCRIPTED) Then
        fraScript.Visible = True
        fraAttributes.Visible = False
        fraEquipment.Visible = False
        fraBow.Visible = False
        fraSpell.Visible = False
        chkStackable.Visible = True
    Else
        fraScript.Visible = False
    End If

End Sub

Private Sub Form_Load()
    picItems.Height = 320 * PIC_Y

    Call BitBlt(picSelect.hDC, 0, 0, PIC_X, PIC_Y, picItems.hDC, EditorItemX * PIC_X, EditorItemY * PIC_Y, SRCCOPY)

    frmItemEditor.picBow.Picture = LoadPicture(App.Path & "\GFX\Arrows.bmp")

    HScroll1.Max = Int(picItems.Width / 32)
    VScroll1.Max = Int(picItems.Height / 32)

    Label8.Caption = STAT1 & " Req:"
    Label9.Caption = STAT2 & " Req:"
    Label10.Caption = STAT3 & " Req:"
    Label20.Caption = "Add " & STAT1 & ":"
    Label21.Caption = "Add " & STAT2 & ":"
    Label22.Caption = "Add " & STAT3 & ":"
    Label23.Caption = "Add " & STAT4 & ":"
End Sub

Private Sub HScroll1_Change()
    picItems.Left = (HScroll1.Value * PIC_X) * -1
End Sub

Private Sub picItems_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = 1 Then
        EditorItemX = Int(X / PIC_X)
        EditorItemY = Int(y / PIC_Y)
    End If
    Call BitBlt(picSelect.hDC, 0, 0, PIC_X, PIC_Y, picItems.hDC, EditorItemX * PIC_X, EditorItemY * PIC_Y, SRCCOPY)
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

Private Sub scrlClassReq_Change()
    If scrlClassReq.Value = -1 Then
        Label16.Caption = "None"
    Else
        Label16.Caption = scrlClassReq.Value & " - " & Trim$(Class(scrlClassReq.Value).Name)
    End If
End Sub

Private Sub scrlDefReq_Change()
    Label12.Caption = STR(scrlDefReq.Value)
End Sub

Private Sub scrlMagicReq_Change()
    Label33.Caption = STR(scrlMagicReq.Value)
End Sub

Private Sub scrlSpeedReq_Change()
    Label13.Caption = STR(scrlSpeedReq.Value)
End Sub

Private Sub scrlStrReq_Change()
    Label11.Caption = STR(scrlStrReq.Value)
End Sub

Private Sub scrlVitalMod_Change()
    lblVitalMod.Caption = STR(scrlVitalMod.Value)
End Sub

Private Sub scrlDurability_Change()
    lblDurability.Caption = STR(scrlDurability.Value)
    If STR(scrlDurability.Value) = -1 Then
        lblDurability.Caption = "Ind."
    End If
End Sub

Private Sub scrlStrength_Change()
    lblStrength.Caption = STR(scrlStrength.Value)
End Sub

Private Sub scrlSpell_Change()
    lblSpellName.Caption = Trim$(Spell(scrlSpell.Value).Name)
    lblSpell.Caption = STR(scrlSpell.Value)
End Sub

Private Sub Timer1_Timer()
    Call BitBlt(picSelect.hDC, 0, 0, PIC_X, PIC_Y, picItems.hDC, EditorItemX * PIC_X, EditorItemY * PIC_Y, SRCCOPY)
End Sub

Private Sub VScroll1_Change()
    picItems.Top = (VScroll1.Value * PIC_Y) * -1
End Sub
Private Sub scrlScript_Change()
    lblScript.Caption = STR(scrlScript.Value)
End Sub

