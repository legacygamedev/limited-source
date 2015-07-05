VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form frmItemEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Editor"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10380
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
   ScaleHeight     =   534
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   692
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
      Left            =   3720
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   27
      Top             =   3510
      Width           =   480
   End
   Begin VB.Frame fraEquipment 
      Caption         =   "Equipment Data"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   3720
      TabIndex        =   2
      Top             =   600
      Width           =   3375
      Visible         =   0   'False
      Begin VB.HScrollBar ScrlLvlReq 
         Height          =   135
         Left            =   960
         Max             =   100
         TabIndex        =   110
         Top             =   2280
         Width           =   1215
      End
      Begin VB.HScrollBar scrlMagiReq 
         Height          =   135
         Left            =   960
         Max             =   1000
         TabIndex        =   82
         Top             =   1560
         Width           =   1215
      End
      Begin VB.HScrollBar scrlAccessReq 
         Height          =   135
         Left            =   960
         Max             =   4
         TabIndex        =   41
         Top             =   2040
         Width           =   1215
      End
      Begin VB.HScrollBar scrlClassReq 
         Height          =   135
         Left            =   960
         Max             =   1
         Min             =   -1
         TabIndex        =   40
         Top             =   1800
         Value           =   -1
         Width           =   1215
      End
      Begin VB.HScrollBar scrlSpeedReq 
         Height          =   135
         Left            =   960
         Max             =   1000
         TabIndex        =   30
         Top             =   1320
         Width           =   1215
      End
      Begin VB.HScrollBar scrlDefReq 
         Height          =   135
         Left            =   960
         Max             =   1000
         TabIndex        =   29
         Top             =   1080
         Width           =   1215
      End
      Begin VB.HScrollBar scrlStrReq 
         Height          =   135
         Left            =   960
         Max             =   1000
         TabIndex        =   28
         Top             =   840
         Width           =   1215
      End
      Begin VB.HScrollBar scrlStrength 
         Height          =   135
         Left            =   960
         Max             =   1000
         TabIndex        =   7
         Top             =   600
         Value           =   1
         Width           =   1215
      End
      Begin VB.HScrollBar scrlDurability 
         Height          =   135
         Left            =   960
         Max             =   5000
         Min             =   -1
         TabIndex        =   5
         Top             =   360
         Value           =   -1
         Width           =   1215
      End
      Begin VB.Label LblLvlReq 
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
         Left            =   2280
         TabIndex        =   111
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         Caption         =   "Level Req:"
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
         TabIndex        =   109
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label30 
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
         Left            =   2280
         TabIndex        =   84
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "INT Req:"
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
         TabIndex        =   83
         Top             =   1560
         Width           =   735
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
         Height          =   255
         Left            =   2280
         TabIndex        =   43
         Top             =   2040
         Width           =   1095
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
         Left            =   2280
         TabIndex        =   42
         Top             =   1800
         Width           =   525
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
         Height          =   255
         Left            =   -120
         TabIndex        =   39
         Top             =   2040
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
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DEX Req :"
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
         TabIndex        =   37
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label13 
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
         Left            =   2280
         TabIndex        =   35
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label12 
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
         Left            =   2280
         TabIndex        =   34
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label11 
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
         Left            =   2280
         TabIndex        =   33
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CON Req :"
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
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "STR Req :"
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
         TabIndex        =   31
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblStrength 
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
         Left            =   2280
         TabIndex        =   8
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblDurability 
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
         Height          =   135
         Left            =   2280
         TabIndex        =   6
         Top             =   360
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
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   615
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
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
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
      Height          =   2400
      Left            =   360
      ScaleHeight     =   160
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
      Left            =   8640
      TabIndex        =   14
      Top             =   6600
      Width           =   1455
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
      Left            =   8640
      TabIndex        =   13
      Top             =   6240
      Width           =   1455
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
      List            =   "frmItemEditor.frx":0043
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
      Height          =   855
      Left            =   3720
      TabIndex        =   9
      Top             =   600
      Width           =   3135
      Visible         =   0   'False
      Begin VB.HScrollBar scrlVitalMod 
         Height          =   135
         Left            =   240
         Max             =   1000
         TabIndex        =   11
         Top             =   600
         Value           =   1
         Width           =   2655
      End
      Begin VB.Label lblVitalMod 
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
         Left            =   960
         TabIndex        =   12
         Top             =   360
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
         Top             =   360
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
      Height          =   1455
      Left            =   3720
      TabIndex        =   16
      Top             =   600
      Width           =   3135
      Visible         =   0   'False
      Begin VB.HScrollBar scrlSpell 
         Height          =   135
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   6915
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   10155
      _ExtentX        =   17912
      _ExtentY        =   12197
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
      TabPicture(0)   =   "frmItemEditor.frx":0111
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label26"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label31"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label32"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label33"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label34"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label35"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "LblWaterSTR"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label36"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "LblWaterDEF"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label37"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "LblEarthSTR"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label38"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "LblEarthDEF"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label39"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label40"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "LblAirSTR"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "LblAirDEF"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label42"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label43"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label44"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label45"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "LblHeatSTR"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "LblColdSTR"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "LblLightSTR"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "LblDarkSTR"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label46"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label47"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Label48"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Label49"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "LblHeatDEF"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "LblColdDEF"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "LblLightDEF"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "LblDarkDEF"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Picture1"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "VScroll1"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "fraAttributes"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "txtDesc"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "fraBow"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "ScrlFireSTR"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "ScrlFireDEF"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "ScrlWaterSTR"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "ScrlWaterDEF"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "ScrlEarthSTR"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "ScrlEarthDEF"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "ScrlAirSTR"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "ScrlAirDEF"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "ScrlHeatSTR"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "ScrlColdSTR"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "ScrlLightSTR"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "ScrlDarkSTR"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "ScrlHeatDEF"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "ScrlColdDEF"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "ScrlLightDEF"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "ScrlDarkDEF"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).ControlCount=   56
      Begin VB.HScrollBar ScrlDarkDEF 
         Height          =   135
         Left            =   5160
         Max             =   1000
         TabIndex        =   131
         Top             =   6600
         Width           =   2415
      End
      Begin VB.HScrollBar ScrlLightDEF 
         Height          =   135
         Left            =   5160
         Max             =   1000
         TabIndex        =   130
         Top             =   6360
         Width           =   2415
      End
      Begin VB.HScrollBar ScrlColdDEF 
         Height          =   135
         Left            =   5160
         Max             =   1000
         TabIndex        =   129
         Top             =   6120
         Width           =   2415
      End
      Begin VB.HScrollBar ScrlHeatDEF 
         Height          =   135
         Left            =   5160
         Max             =   1000
         TabIndex        =   128
         Top             =   5880
         Width           =   2415
      End
      Begin VB.HScrollBar ScrlDarkSTR 
         Height          =   135
         Left            =   960
         Max             =   1000
         TabIndex        =   119
         Top             =   6600
         Width           =   2415
      End
      Begin VB.HScrollBar ScrlLightSTR 
         Height          =   135
         Left            =   960
         Max             =   1000
         TabIndex        =   118
         Top             =   6360
         Width           =   2415
      End
      Begin VB.HScrollBar ScrlColdSTR 
         Height          =   135
         Left            =   960
         Max             =   1000
         TabIndex        =   117
         Top             =   6120
         Width           =   2415
      End
      Begin VB.HScrollBar ScrlHeatSTR 
         Height          =   135
         Left            =   960
         Max             =   1000
         TabIndex        =   116
         Top             =   5880
         Width           =   2415
      End
      Begin VB.HScrollBar ScrlAirDEF 
         Height          =   135
         Left            =   5160
         Max             =   1000
         TabIndex        =   106
         Top             =   5640
         Width           =   2415
      End
      Begin VB.HScrollBar ScrlAirSTR 
         Height          =   135
         Left            =   960
         Max             =   1000
         TabIndex        =   105
         Top             =   5640
         Width           =   2415
      End
      Begin VB.HScrollBar ScrlEarthDEF 
         Height          =   135
         Left            =   5160
         Max             =   1000
         TabIndex        =   101
         Top             =   5400
         Width           =   2415
      End
      Begin VB.HScrollBar ScrlEarthSTR 
         Height          =   135
         Left            =   960
         Max             =   1000
         TabIndex        =   98
         Top             =   5400
         Width           =   2415
      End
      Begin VB.HScrollBar ScrlWaterDEF 
         Height          =   135
         Left            =   5160
         Max             =   1000
         TabIndex        =   94
         Top             =   5160
         Width           =   2415
      End
      Begin VB.HScrollBar ScrlWaterSTR 
         Height          =   135
         Left            =   960
         Max             =   1000
         TabIndex        =   91
         Top             =   5160
         Width           =   2415
      End
      Begin VB.HScrollBar ScrlFireDEF 
         Height          =   135
         Left            =   5160
         Max             =   1000
         TabIndex        =   88
         Top             =   4920
         Width           =   2415
      End
      Begin VB.HScrollBar ScrlFireSTR 
         Height          =   135
         Left            =   960
         Max             =   1000
         TabIndex        =   85
         Top             =   4920
         Width           =   2415
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
         Left            =   4440
         TabIndex        =   71
         Top             =   3240
         Width           =   2535
         Visible         =   0   'False
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
            ItemData        =   "frmItemEditor.frx":012D
            Left            =   120
            List            =   "frmItemEditor.frx":012F
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
         Height          =   255
         Left            =   7080
         MaxLength       =   150
         TabIndex        =   69
         Top             =   4080
         Width           =   2895
      End
      Begin VB.Frame fraAttributes 
         Caption         =   "Attributes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   7080
         TabIndex        =   44
         Top             =   480
         Width           =   2895
         Visible         =   0   'False
         Begin VB.HScrollBar scrlAttackSpeed 
            Height          =   135
            Left            =   1080
            Max             =   5000
            Min             =   1
            TabIndex        =   79
            Top             =   2520
            Value           =   1000
            Width           =   1215
         End
         Begin VB.HScrollBar scrlAddEXP 
            Height          =   135
            Left            =   1080
            Max             =   100
            TabIndex        =   66
            Top             =   2040
            Width           =   1215
         End
         Begin VB.HScrollBar scrlAddSP 
            Height          =   135
            Left            =   1080
            Max             =   1000
            Min             =   -100
            TabIndex        =   64
            Top             =   840
            Width           =   1215
         End
         Begin VB.HScrollBar scrlAddSpeed 
            Height          =   135
            Left            =   1080
            Max             =   1000
            Min             =   -100
            TabIndex        =   56
            Top             =   1560
            Width           =   1215
         End
         Begin VB.HScrollBar scrlAddMagi 
            Height          =   135
            Left            =   1080
            Max             =   1000
            Min             =   -100
            TabIndex        =   55
            Top             =   1800
            Width           =   1215
         End
         Begin VB.HScrollBar scrlAddDef 
            Height          =   135
            Left            =   1080
            Max             =   1000
            Min             =   -100
            TabIndex        =   54
            Top             =   1320
            Width           =   1215
         End
         Begin VB.HScrollBar scrlAddStr 
            Height          =   135
            Left            =   1080
            Max             =   1000
            Min             =   -100
            TabIndex        =   53
            Top             =   1080
            Width           =   1215
         End
         Begin VB.HScrollBar scrlAddMP 
            Height          =   135
            Left            =   1080
            Max             =   1000
            Min             =   -100
            TabIndex        =   52
            Top             =   600
            Width           =   1215
         End
         Begin VB.HScrollBar scrlAddHP 
            Height          =   135
            Left            =   1080
            Max             =   1000
            Min             =   -100
            TabIndex        =   51
            Top             =   360
            Width           =   1215
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
            Left            =   1080
            TabIndex        =   81
            Top             =   2280
            Width           =   1110
         End
         Begin VB.Label Label28 
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
            Height          =   255
            Left            =   120
            TabIndex        =   80
            Top             =   2280
            Width           =   975
         End
         Begin VB.Label lblAddEXP 
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
            Height          =   135
            Left            =   2400
            TabIndex        =   68
            Top             =   2040
            Width           =   375
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
            Height          =   255
            Left            =   240
            TabIndex        =   67
            Top             =   2040
            Width           =   735
         End
         Begin VB.Label lblAddSP 
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
            Left            =   2400
            TabIndex        =   65
            Top             =   840
            Width           =   375
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
            Height          =   255
            Left            =   360
            TabIndex        =   63
            Top             =   840
            Width           =   615
         End
         Begin VB.Label lblAddSpeed 
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
            Left            =   2400
            TabIndex        =   62
            Top             =   1560
            Width           =   375
         End
         Begin VB.Label lblAddMagi 
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
            Left            =   2400
            TabIndex        =   61
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label lblAddDef 
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
            Left            =   2400
            TabIndex        =   60
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label lblAddStr 
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
            Left            =   2400
            TabIndex        =   59
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label lblAddMP 
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
            Left            =   2400
            TabIndex        =   58
            Top             =   600
            Width           =   375
         End
         Begin VB.Label lblAddHP 
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
            Left            =   2400
            TabIndex        =   57
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Add DEX :"
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
            TabIndex        =   50
            Top             =   1560
            Width           =   855
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            Caption         =   "Add INT :"
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
            TabIndex        =   49
            Top             =   1800
            Width           =   735
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            Caption         =   "Add CON :"
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
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            Caption         =   "Add STR :"
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
            Top             =   1080
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
            Height          =   255
            Left            =   240
            TabIndex        =   46
            Top             =   600
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
            Height          =   255
            Left            =   360
            TabIndex        =   45
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   2400
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
         Left            =   3570
         ScaleHeight     =   510
         ScaleWidth      =   510
         TabIndex        =   36
         Top             =   3360
         Width           =   540
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
         Height          =   135
         Left            =   7680
         TabIndex        =   135
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
         Height          =   135
         Left            =   7680
         TabIndex        =   134
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
         Height          =   135
         Left            =   7680
         TabIndex        =   133
         Top             =   6120
         Width           =   615
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
         Height          =   135
         Left            =   7680
         TabIndex        =   132
         Top             =   5880
         Width           =   615
      End
      Begin VB.Label Label49 
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
         Left            =   4320
         TabIndex        =   127
         Top             =   6600
         Width           =   735
      End
      Begin VB.Label Label48 
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
         Left            =   4320
         TabIndex        =   126
         Top             =   6360
         Width           =   735
      End
      Begin VB.Label Label47 
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
         Left            =   4320
         TabIndex        =   125
         Top             =   6120
         Width           =   735
      End
      Begin VB.Label Label46 
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
         Left            =   4320
         TabIndex        =   124
         Top             =   5880
         Width           =   735
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
         Height          =   135
         Left            =   3480
         TabIndex        =   123
         Top             =   6600
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
         Height          =   135
         Left            =   3480
         TabIndex        =   122
         Top             =   6360
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
         Height          =   135
         Left            =   3480
         TabIndex        =   121
         Top             =   6120
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
         Height          =   135
         Left            =   3480
         TabIndex        =   120
         Top             =   5880
         Width           =   615
      End
      Begin VB.Label Label45 
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
         Left            =   120
         TabIndex        =   115
         Top             =   6600
         Width           =   735
      End
      Begin VB.Label Label44 
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
         Left            =   120
         TabIndex        =   114
         Top             =   6360
         Width           =   735
      End
      Begin VB.Label Label43 
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
         Left            =   120
         TabIndex        =   113
         Top             =   6120
         Width           =   735
      End
      Begin VB.Label Label42 
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
         Left            =   120
         TabIndex        =   112
         Top             =   5880
         Width           =   735
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
         Height          =   135
         Left            =   7680
         TabIndex        =   108
         Top             =   5640
         Width           =   615
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
         Height          =   135
         Left            =   3480
         TabIndex        =   107
         Top             =   5640
         Width           =   615
      End
      Begin VB.Label Label40 
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
         Height          =   135
         Left            =   4440
         TabIndex        =   104
         Top             =   5640
         Width           =   615
      End
      Begin VB.Label Label39 
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
         Height          =   135
         Left            =   120
         TabIndex        =   103
         Top             =   5640
         Width           =   735
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
         Height          =   135
         Left            =   7680
         TabIndex        =   102
         Top             =   5400
         Width           =   615
      End
      Begin VB.Label Label38 
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
         Height          =   135
         Left            =   4320
         TabIndex        =   100
         Top             =   5400
         Width           =   735
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
         Height          =   135
         Left            =   3480
         TabIndex        =   99
         Top             =   5400
         Width           =   615
      End
      Begin VB.Label Label37 
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
         Height          =   135
         Left            =   120
         TabIndex        =   97
         Top             =   5400
         Width           =   735
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
         Height          =   135
         Left            =   7680
         TabIndex        =   96
         Top             =   5160
         Width           =   615
      End
      Begin VB.Label Label36 
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
         Height          =   135
         Left            =   4320
         TabIndex        =   95
         Top             =   5160
         Width           =   735
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
         Height          =   135
         Left            =   3480
         TabIndex        =   93
         Top             =   5160
         Width           =   615
      End
      Begin VB.Label Label35 
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
         Height          =   135
         Left            =   120
         TabIndex        =   92
         Top             =   5160
         Width           =   735
      End
      Begin VB.Label Label34 
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
         Left            =   7680
         TabIndex        =   90
         Top             =   4920
         Width           =   615
      End
      Begin VB.Label Label33 
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
         Height          =   135
         Left            =   4440
         TabIndex        =   89
         Top             =   4920
         Width           =   615
      End
      Begin VB.Label Label32 
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
         Left            =   3480
         TabIndex        =   87
         Top             =   4920
         Width           =   615
      End
      Begin VB.Label Label31 
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
         Height          =   135
         Left            =   120
         TabIndex        =   86
         Top             =   4920
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
         Left            =   7080
         TabIndex        =   70
         Top             =   3840
         Width           =   855
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
        cmbBow.AddItem "None", 0
        cmbBow.ListIndex = 0
        cmbBow.Enabled = False
        lblName.Caption = ""
    Else
        cmbBow.Clear
        For i = 1 To MAX_ARROWS
            cmbBow.AddItem i & ": " & Arrows(i).Name
        Next i
        cmbBow.ListIndex = 0
        cmbBow.Enabled = True
    End If
End Sub

Private Sub cmbBow_Click()
    lblName.Caption = Arrows(cmbBow.ListIndex + 1).Name
    picBow.Top = (Arrows(cmbBow.ListIndex + 1).Pic * 32) * -1
End Sub

Private Sub cmdOk_Click()
    Call ItemEditorOk
End Sub

Private Sub cmdCancel_Click()
    Call ItemEditorCancel
End Sub

Private Sub cmbType_Click()
    If (cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (cmbType.ListIndex <= ITEM_TYPE_SHIELD) Or (cmbType.ListIndex >= ITEM_TYPE_GLOVES) And (cmbType.ListIndex <= ITEM_TYPE_AMULET) Then
        If cmbType.ListIndex = ITEM_TYPE_WEAPON Or cmbType.ListIndex = ITEM_TYPE_TWO_HAND Then
        
            Label3.Caption = "Damage :"
        Else
            Label3.Caption = "Defence :"
        End If
        fraEquipment.Visible = True
        fraAttributes.Visible = True
        fraBow.Visible = True
    Else
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
    Else
        fraSpell.Visible = False
    End If
End Sub

Private Sub Form_Load()
    picItems.Height = 320 * PIC_Y
    Call BitBlt(picSelect.hDC, 0, 0, PIC_X, PIC_Y, picItems.hDC, EditorItemX * PIC_X, EditorItemY * PIC_Y, SRCCOPY)
    picBow.Picture = LoadPicture(App.Path & "\GFX\arrows.bmp")
End Sub

Private Sub picItems_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        EditorItemX = Int(x / PIC_X)
        EditorItemY = Int(y / PIC_Y)
    End If
    Call BitBlt(picSelect.hDC, 0, 0, PIC_X, PIC_Y, picItems.hDC, EditorItemX * PIC_X, EditorItemY * PIC_Y, SRCCOPY)
End Sub

Private Sub picItems_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        EditorItemX = Int(x / PIC_X)
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
    Label16.Caption = scrlClassReq.Value & " - " & Trim(Class(scrlClassReq.Value).Name)
End If
End Sub

Private Sub scrlDefReq_Change()
    Label12.Caption = scrlDefReq.Value
End Sub

Private Sub ScrlEarthDEF_Change()
    LblEarthDEF.Caption = ScrlEarthDEF.Value
End Sub

Private Sub ScrlEarthSTR_Change()
    LblEarthSTR.Caption = ScrlEarthSTR.Value
End Sub
Private Sub ScrlAirDEF_Change()
    LblAirDEF.Caption = ScrlAirDEF.Value
End Sub

Private Sub ScrlAirSTR_Change()
    LblAirSTR.Caption = ScrlAirSTR.Value
End Sub
Private Sub ScrlFireDEF_Change()
    Label34.Caption = ScrlFireDEF.Value
End Sub

Private Sub ScrlFireSTR_Change()
    Label32.Caption = ScrlFireSTR.Value
End Sub
Private Sub ScrlHeatDEF_Change()
    LblHeatDEF.Caption = ScrlHeatDEF.Value
End Sub

Private Sub ScrlHeatSTR_Change()
    LblHeatSTR.Caption = ScrlHeatSTR.Value
End Sub
Private Sub ScrlColdDEF_Change()
    LblColdDEF.Caption = ScrlColdDEF.Value
End Sub

Private Sub ScrlColdSTR_Change()
    LblColdSTR.Caption = ScrlColdSTR.Value
End Sub
Private Sub ScrlLightDEF_Change()
    LblLightDEF.Caption = ScrlLightDEF.Value
End Sub
Private Sub ScrlLightSTR_Change()
    LblLightSTR.Caption = ScrlLightSTR.Value
End Sub
Private Sub ScrlDarkDEF_Change()
    LblDarkDEF.Caption = ScrlDarkDEF.Value
End Sub
Private Sub ScrlDarkSTR_Change()
    LblDarkSTR.Caption = ScrlDarkSTR.Value
End Sub
Private Sub scrlSpeedReq_Change()
    Label13.Caption = scrlSpeedReq.Value
End Sub
Private Sub scrlMagiReq_Change()
    Label30.Caption = scrlMagiReq.Value
End Sub
Private Sub scrlLvlReq_Change()
    LblLvlReq.Caption = ScrlLvlReq.Value
End Sub
Private Sub scrlStrReq_Change()
    Label11.Caption = scrlStrReq.Value
End Sub

Private Sub ScrlWaterDEF_Change()
    LblWaterDEF.Caption = STR(ScrlWaterDEF.Value)
End Sub

Private Sub ScrlWaterSTR_Change()
    LblWaterSTR.Caption = STR(ScrlWaterSTR.Value)
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
    lblSpellName.Caption = Trim(Spell(scrlSpell.Value).Name)
    lblSpell.Caption = STR(scrlSpell.Value)
End Sub

Private Sub Timer1_Timer()
    Call BitBlt(picSelect.hDC, 0, 0, PIC_X, PIC_Y, picItems.hDC, EditorItemX * PIC_X, EditorItemY * PIC_Y, SRCCOPY)
End Sub

Private Sub VScroll1_Change()
    picItems.Top = (VScroll1.Value * PIC_Y) * -1
End Sub
