VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmSpellEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editor de Hechizos"
   ClientHeight    =   6015
   ClientLeft      =   135
   ClientTop       =   420
   ClientWidth     =   8115
   ControlBox      =   0   'False
   DrawMode        =   14  'Copy Pen
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
   ScaleHeight     =   401
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   541
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   120
      Top             =   120
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   10398
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   397
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
      TabCaption(0)   =   "Hechizo"
      TabPicture(0)   =   "frmSpellEditor.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblSound"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblVitalMod"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label7"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblRange"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblSpellAnim"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblSpellTime"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblSpellDone"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "LblFireSTR"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label9"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label10"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label11"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label12"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label13"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label14"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label15"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label16"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "LblWaterSTR"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "LblEarthSTR"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "LblAirSTR"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "LblHeatSTR"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "LblColdSTR"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "LblLightSTR"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "LblDarkSTR"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "chkArea"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "info"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Frame1"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "scrlSound"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "scrlVitalMod"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "cmbType"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "cmdCancel"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "cmdOk"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "scrlRange"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "scrlSpellAnim"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "scrlSpellTime"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "scrlSpellDone"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "picSpell"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Command1"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "ScrlFireSTR"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "ScrlWaterSTR"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "ScrlEarthSTR"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "ScrlAirSTR"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "ScrlHeatSTR"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "ScrlColdSTR"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "ScrlLightSTR"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "ScrlDarkSTR"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).ControlCount=   48
      Begin VB.HScrollBar ScrlDarkSTR 
         Height          =   135
         Left            =   960
         Max             =   1000
         TabIndex        =   54
         Top             =   5400
         Width           =   2175
      End
      Begin VB.HScrollBar ScrlLightSTR 
         Height          =   135
         Left            =   960
         Max             =   1000
         TabIndex        =   53
         Top             =   5160
         Width           =   2175
      End
      Begin VB.HScrollBar ScrlColdSTR 
         Height          =   135
         Left            =   960
         Max             =   1000
         TabIndex        =   52
         Top             =   4920
         Width           =   2175
      End
      Begin VB.HScrollBar ScrlHeatSTR 
         Height          =   135
         Left            =   960
         Max             =   1000
         TabIndex        =   51
         Top             =   4680
         Width           =   2175
      End
      Begin VB.HScrollBar ScrlAirSTR 
         Height          =   135
         Left            =   960
         Max             =   1000
         TabIndex        =   48
         Top             =   4440
         Width           =   2175
      End
      Begin VB.HScrollBar ScrlEarthSTR 
         Height          =   135
         Left            =   960
         Max             =   1000
         TabIndex        =   47
         Top             =   4200
         Width           =   2175
      End
      Begin VB.HScrollBar ScrlWaterSTR 
         Height          =   135
         Left            =   960
         Max             =   1000
         TabIndex        =   45
         Top             =   3960
         Width           =   2175
      End
      Begin VB.HScrollBar ScrlFireSTR 
         Height          =   135
         Left            =   960
         Max             =   1000
         TabIndex        =   35
         Top             =   3720
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Actualizar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   220
         Left            =   6480
         TabIndex        =   33
         Top             =   2880
         Width           =   1095
      End
      Begin VB.PictureBox picSpell 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   7080
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   32
         Top             =   4800
         Width           =   480
      End
      Begin VB.HScrollBar scrlSpellDone 
         Height          =   135
         Left            =   4080
         Max             =   10
         Min             =   1
         TabIndex        =   31
         Top             =   4440
         Value           =   1
         Width           =   3495
      End
      Begin VB.HScrollBar scrlSpellTime 
         Height          =   135
         Left            =   4080
         Max             =   500
         Min             =   40
         TabIndex        =   30
         Top             =   3840
         Value           =   40
         Width           =   3495
      End
      Begin VB.HScrollBar scrlSpellAnim 
         Height          =   135
         Left            =   4080
         Max             =   2000
         TabIndex        =   29
         Top             =   3240
         Value           =   1
         Width           =   3495
      End
      Begin VB.HScrollBar scrlRange 
         Height          =   135
         Left            =   240
         Max             =   30
         Min             =   1
         TabIndex        =   25
         Top             =   3360
         Value           =   1
         Width           =   3495
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
         Height          =   285
         Left            =   6600
         TabIndex        =   22
         Top             =   5520
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5280
         TabIndex        =   21
         Top             =   5520
         Width           =   1230
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
         ItemData        =   "frmSpellEditor.frx":001C
         Left            =   4080
         List            =   "frmSpellEditor.frx":0035
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2400
         Width           =   3495
      End
      Begin VB.HScrollBar scrlVitalMod 
         Height          =   135
         Left            =   240
         Max             =   1000
         TabIndex        =   14
         Top             =   2400
         Width           =   3495
      End
      Begin VB.HScrollBar scrlSound 
         Height          =   135
         Left            =   240
         Max             =   100
         TabIndex        =   13
         Top             =   2880
         Width           =   3495
      End
      Begin VB.Frame Frame1 
         Caption         =   "Cualidades"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   4080
         TabIndex        =   6
         Top             =   480
         Width           =   3615
         Begin VB.HScrollBar scrlLevelReq 
            Height          =   135
            Left            =   120
            Max             =   500
            TabIndex        =   8
            Top             =   600
            Value           =   1
            Width           =   3375
         End
         Begin VB.HScrollBar scrlCost 
            Height          =   135
            Left            =   120
            Max             =   1000
            TabIndex        =   7
            Top             =   1200
            Width           =   3375
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nivel Requerido:"
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
            Left            =   165
            TabIndex        =   12
            Top             =   360
            Width           =   1035
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Coste de MP:"
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
            TabIndex        =   11
            Top             =   960
            Width           =   855
         End
         Begin VB.Label lblLevelReq 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hechizo de Admins"
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
            TabIndex        =   10
            Top             =   360
            Width           =   1185
         End
         Begin VB.Label lblCost 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   1320
            TabIndex        =   9
            Top             =   960
            Width           =   75
         End
      End
      Begin VB.Frame info 
         Caption         =   "Informacion"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   3615
         Begin VB.ComboBox cmbClassReq 
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
            ItemData        =   "frmSpellEditor.frx":0074
            Left            =   120
            List            =   "frmSpellEditor.frx":0076
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1080
            Width           =   3345
         End
         Begin VB.TextBox txtName 
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
            TabIndex        =   2
            Top             =   480
            Width           =   3315
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Clase Requerida"
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
            TabIndex        =   5
            Top             =   840
            Width           =   1005
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre"
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
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.CheckBox chkArea 
         Alignment       =   1  'Right Justify
         Caption         =   "Efecto de Area"
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
         TabIndex        =   34
         Top             =   4920
         Width           =   1455
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
         Left            =   3240
         TabIndex        =   58
         Top             =   5400
         Width           =   495
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
         Left            =   3240
         TabIndex        =   57
         Top             =   5160
         Width           =   495
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
         Left            =   3240
         TabIndex        =   56
         Top             =   4920
         Width           =   495
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
         Left            =   3240
         TabIndex        =   55
         Top             =   4680
         Width           =   495
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
         Left            =   3240
         TabIndex        =   50
         Top             =   4440
         Width           =   495
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
         Left            =   3240
         TabIndex        =   49
         Top             =   4200
         Width           =   495
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
         Left            =   3240
         TabIndex        =   46
         Top             =   3960
         Width           =   495
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Oscuro + :"
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
         Top             =   5400
         Width           =   735
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Luz + :"
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
         Top             =   5160
         Width           =   735
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Hielo + :"
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
         Top             =   4920
         Width           =   735
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Rayo + :"
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
         Top             =   4680
         Width           =   735
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Aire + :"
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
         TabIndex        =   40
         Top             =   4440
         Width           =   615
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Tierra + :"
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
         Top             =   4200
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Agua + :"
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
         Top             =   3960
         Width           =   735
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Fuego + :"
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
         Top             =   3720
         Width           =   735
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
         Height          =   135
         Left            =   3240
         TabIndex        =   36
         Top             =   3720
         Width           =   495
      End
      Begin VB.Label lblSpellDone 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ciclo de Animacion 1 Tiempo"
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
         Left            =   4080
         TabIndex        =   28
         Top             =   4200
         Width           =   1815
      End
      Begin VB.Label lblSpellTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tiempo: 40"
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
         Left            =   4080
         TabIndex        =   27
         Top             =   3600
         Width           =   705
      End
      Begin VB.Label lblSpellAnim 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Anim: 0"
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
         Left            =   4080
         TabIndex        =   26
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   720
         TabIndex        =   24
         Top             =   3120
         Width           =   75
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Rango:"
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
         Top             =   3120
         Width           =   780
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Hechizo"
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
         TabIndex        =   20
         Top             =   2160
         Width           =   1200
      End
      Begin VB.Label lblVitalMod 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   1200
         TabIndex        =   19
         Top             =   2160
         Width           =   75
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Modo Vital:"
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
         Top             =   2160
         Width           =   780
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Sonido:"
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
         TabIndex        =   17
         Top             =   2640
         Width           =   780
      End
      Begin VB.Label lblSound 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sin Sonido"
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
         Left            =   840
         TabIndex        =   16
         Top             =   2640
         Width           =   660
      End
   End
End
Attribute VB_Name = "frmSpellEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Done As Long
Private Time As Long
Private SpellVar As Long

Private Sub Command1_Click()
    Done = 0
End Sub

Private Sub scrlCost_Change()
    lblCost.Caption = STR(scrlCost.Value)
End Sub
Private Sub ScrlFireSTR_Change()
    LblFireSTR.Caption = STR(ScrlFireSTR.Value)
End Sub
Private Sub ScrlWaterSTR_Change()
    LblWaterSTR.Caption = STR(ScrlWaterSTR.Value)
End Sub
Private Sub ScrlEarthSTR_Change()
    LblEarthSTR.Caption = STR(ScrlEarthSTR.Value)
End Sub
Private Sub ScrlAirSTR_Change()
    LblAirSTR.Caption = STR(ScrlAirSTR.Value)
End Sub
Private Sub ScrlHeatSTR_Change()
    LblHeatSTR.Caption = STR(ScrlHeatSTR.Value)
End Sub
Private Sub ScrlColdSTR_Change()
    LblColdSTR.Caption = STR(ScrlColdSTR.Value)
End Sub
Private Sub ScrlLightSTR_Change()
    LblLightSTR.Caption = STR(ScrlLightSTR.Value)
End Sub
Private Sub ScrlDarkSTR_Change()
    LblDarkSTR.Caption = STR(ScrlDarkSTR.Value)
End Sub

Private Sub scrlLevelReq_Change()
    If STR(scrlLevelReq.Value) = 0 Then
        lblLevelReq.Caption = "Hechizo de Admins"
    Else
        lblLevelReq.Caption = STR(scrlLevelReq.Value)
    End If
End Sub

Private Sub scrlRange_Change()
    lblRange.Caption = STR(scrlRange.Value)
End Sub

Private Sub scrlSound_Change()
If STR(scrlSound.Value) = 0 Then
    lblSound.Caption = "Sin Sonido"
Else
    lblSound.Caption = STR(scrlSound.Value)
    Call PlaySound("magic" & scrlSound.Value & ".wav")
End If
End Sub

Private Sub scrlSpellAnim_Change()
    lblSpellAnim.Caption = "Anim: " & scrlSpellAnim.Value
    Done = 0
End Sub

Private Sub scrlSpellDone_Change()
Dim String2 As String
    String2 = "Tiempos"
    If scrlSpellDone.Value = 1 Then String2 = "Tiempo"
    lblSpellDone.Caption = "Ciclo de Animacion " & scrlSpellDone.Value & " " & String2
    Done = 0
End Sub

Private Sub scrlSpellTime_Change()
    lblSpellTime.Caption = "Tiempo: " & scrlSpellTime.Value
    Done = 0
End Sub

Private Sub scrlVitalMod_Change()
    If (cmbType.ListIndex = SPELL_TYPE_GIVEITEM) Then
            Label4.Caption = "Dar Item:"
        Else
            Label4.Caption = "Modo Vital:"
        End If
    lblVitalMod.Caption = STR(scrlVitalMod.Value)
End Sub

Private Sub cmdOk_Click()
    Call SpellEditorOk
End Sub

Private Sub cmdCancel_Click()
    Call SpellEditorCancel
End Sub

Private Sub Timer1_Timer()
Dim sRECT As RECT
Dim dRECT As RECT
Dim SpellDone As Long
Dim SpellAnim As Long
Dim SpellTime As Long

SpellDone = scrlSpellDone.Value
SpellAnim = scrlSpellAnim.Value
SpellTime = scrlSpellTime.Value

If SpellAnim <= 0 Then Exit Sub
If Done = SpellDone Then Exit Sub

    With dRECT
        .Top = 0
        .Bottom = PIC_Y
        .Left = 0
        .Right = PIC_X
    End With

    If SpellVar > 10 Then
        Done = Done + 1
        SpellVar = 0
    End If
    If GetTickCount > Time + SpellTime Then
        Time = GetTickCount
        SpellVar = SpellVar + 1
    End If

    If DD_SpellAnim Is Nothing Then
    Else
        With sRECT
            .Top = SpellAnim * PIC_Y
            .Bottom = .Top + PIC_Y
            .Left = SpellVar * PIC_X
            .Right = .Left + PIC_X
        End With
        
        Call DD_SpellAnim.BltToDC(picSpell.hDC, sRECT, dRECT)
        picSpell.Refresh
    End If
End Sub
