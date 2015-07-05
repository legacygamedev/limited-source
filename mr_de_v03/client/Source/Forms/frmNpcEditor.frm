VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmNpcEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Npc Editor"
   ClientHeight    =   8985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5385
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
   ScaleHeight     =   8985
   ScaleWidth      =   5385
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraShopKeeper 
      Height          =   975
      Left            =   240
      TabIndex        =   68
      Top             =   3120
      Visible         =   0   'False
      Width           =   4935
      Begin VB.ComboBox cmbShop 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   425
         Width           =   2415
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "Shop:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   70
         Top             =   420
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   8520
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   13
      Top             =   8520
      Width           =   975
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
      Left            =   120
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   9360
      Width           =   480
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8775
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   15478
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   370
      TabMaxWidth     =   1764
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Info."
      TabPicture(0)   =   "frmNpcEditor.frx":0E42
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraNpc"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Frame Frame1 
         Height          =   2655
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4935
         Begin VB.ComboBox cmbMovementFrequency 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmNpcEditor.frx":0E5E
            Left            =   960
            List            =   "frmNpcEditor.frx":0E6C
            Style           =   2  'Dropdown List
            TabIndex        =   71
            Top             =   2200
            Width           =   3855
         End
         Begin VB.HScrollBar scrlMovementSpeed 
            Height          =   255
            Left            =   960
            Max             =   3
            Min             =   1
            TabIndex        =   61
            Top             =   1440
            Value           =   1
            Width           =   2655
         End
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   570
            Left            =   4245
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   15
            Top             =   1125
            Width           =   570
            Begin VB.PictureBox picSprite 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00FFFFFF&
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
               Left            =   30
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   16
               Top             =   30
               Width           =   480
            End
         End
         Begin VB.ComboBox cmbBehavior 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmNpcEditor.frx":0E8B
            Left            =   960
            List            =   "frmNpcEditor.frx":0E9E
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1800
            Width           =   3855
         End
         Begin VB.HScrollBar scrlSprite 
            Height          =   255
            Left            =   960
            Max             =   255
            TabIndex        =   5
            Top             =   1080
            Width           =   2655
         End
         Begin VB.TextBox txtName 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   960
            TabIndex        =   4
            Top             =   240
            Width           =   3855
         End
         Begin VB.TextBox txtAttackSay 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   960
            TabIndex        =   3
            Top             =   720
            Width           =   3855
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Frequency:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   72
            Top             =   2200
            Width           =   855
         End
         Begin VB.Label lblMovementSpeed 
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3720
            TabIndex        =   62
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Speed:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   60
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Behaviour:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   12
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label lblSprite 
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3720
            TabIndex        =   9
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Sprite:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   8
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Name:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   7
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "Spoken:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   6
            Top             =   720
            Width           =   615
         End
      End
      Begin VB.Frame fraNpc 
         Height          =   5415
         Left            =   120
         TabIndex        =   10
         Top             =   2880
         Visible         =   0   'False
         Width           =   4935
         Begin VB.HScrollBar scrlLevel 
            Height          =   255
            Left            =   1320
            Max             =   255
            TabIndex        =   91
            Top             =   1560
            Width           =   2655
         End
         Begin VB.TextBox txtEXP 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1320
            TabIndex        =   90
            Text            =   "Text1"
            Top             =   2040
            Width           =   2655
         End
         Begin VB.TextBox txtHP 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1320
            TabIndex        =   89
            Text            =   "Text1"
            Top             =   1800
            Width           =   2655
         End
         Begin VB.CommandButton cmdCalc 
            Caption         =   "Calc"
            Height          =   495
            Left            =   120
            TabIndex        =   88
            Top             =   1560
            Width           =   615
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   0
            Left            =   1200
            Max             =   5000
            TabIndex        =   77
            Top             =   240
            Value           =   1
            Width           =   2775
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   1
            Left            =   1200
            Max             =   5000
            TabIndex        =   76
            Top             =   480
            Value           =   1
            Width           =   2775
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   2
            Left            =   1200
            Max             =   5000
            TabIndex        =   75
            Top             =   720
            Value           =   1
            Width           =   2775
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   3
            Left            =   1200
            Max             =   5000
            TabIndex        =   74
            Top             =   960
            Value           =   1
            Width           =   2775
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   4
            Left            =   1200
            Max             =   5000
            TabIndex        =   73
            Top             =   1200
            Value           =   1
            Width           =   2775
         End
         Begin VB.HScrollBar scrlRange 
            Height          =   255
            Left            =   1080
            Max             =   255
            TabIndex        =   67
            Top             =   2790
            Value           =   1
            Width           =   2895
         End
         Begin VB.TextBox txtSpawnSecs 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2520
            TabIndex        =   58
            Text            =   "0"
            Top             =   2400
            Width           =   1815
         End
         Begin TabDlg.SSTab SSTab2 
            Height          =   2175
            Left            =   120
            TabIndex        =   17
            Top             =   3120
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   3836
            _Version        =   393216
            Tabs            =   4
            TabsPerRow      =   4
            TabHeight       =   370
            TabMaxWidth     =   1587
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "Drop 1"
            TabPicture(0)   =   "frmNpcEditor.frx":0EEB
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "lblNum"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Label9"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "Label11"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "lblItemName"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "Label7"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "lblValue"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "Label3"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "scrlNum"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "scrlValue"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "txtChance"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).ControlCount=   10
            TabCaption(1)   =   "Drop 2"
            TabPicture(1)   =   "frmNpcEditor.frx":0F07
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "txtChance2"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).Control(1)=   "scrlValue2"
            Tab(1).Control(1).Enabled=   0   'False
            Tab(1).Control(2)=   "scrlNum2"
            Tab(1).Control(2).Enabled=   0   'False
            Tab(1).Control(3)=   "Label21"
            Tab(1).Control(3).Enabled=   0   'False
            Tab(1).Control(4)=   "lblValue2"
            Tab(1).Control(4).Enabled=   0   'False
            Tab(1).Control(5)=   "Label19"
            Tab(1).Control(5).Enabled=   0   'False
            Tab(1).Control(6)=   "lblItemName2"
            Tab(1).Control(6).Enabled=   0   'False
            Tab(1).Control(7)=   "Label17"
            Tab(1).Control(7).Enabled=   0   'False
            Tab(1).Control(8)=   "Label15"
            Tab(1).Control(8).Enabled=   0   'False
            Tab(1).Control(9)=   "lblNum2"
            Tab(1).Control(9).Enabled=   0   'False
            Tab(1).ControlCount=   10
            TabCaption(2)   =   "Drop 3"
            TabPicture(2)   =   "frmNpcEditor.frx":0F23
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "txtChance3"
            Tab(2).Control(0).Enabled=   0   'False
            Tab(2).Control(1)=   "scrlValue3"
            Tab(2).Control(1).Enabled=   0   'False
            Tab(2).Control(2)=   "scrlNum3"
            Tab(2).Control(2).Enabled=   0   'False
            Tab(2).Control(3)=   "Label29"
            Tab(2).Control(3).Enabled=   0   'False
            Tab(2).Control(4)=   "lblValue3"
            Tab(2).Control(4).Enabled=   0   'False
            Tab(2).Control(5)=   "Label27"
            Tab(2).Control(5).Enabled=   0   'False
            Tab(2).Control(6)=   "lblItemName3"
            Tab(2).Control(6).Enabled=   0   'False
            Tab(2).Control(7)=   "Label25"
            Tab(2).Control(7).Enabled=   0   'False
            Tab(2).Control(8)=   "Label24"
            Tab(2).Control(8).Enabled=   0   'False
            Tab(2).Control(9)=   "lblNum3"
            Tab(2).Control(9).Enabled=   0   'False
            Tab(2).ControlCount=   10
            TabCaption(3)   =   "Drop 4"
            TabPicture(3)   =   "frmNpcEditor.frx":0F3F
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "txtChance4"
            Tab(3).Control(0).Enabled=   0   'False
            Tab(3).Control(1)=   "scrlValue4"
            Tab(3).Control(1).Enabled=   0   'False
            Tab(3).Control(2)=   "scrlNum4"
            Tab(3).Control(2).Enabled=   0   'False
            Tab(3).Control(3)=   "Label37"
            Tab(3).Control(3).Enabled=   0   'False
            Tab(3).Control(4)=   "lblValue4"
            Tab(3).Control(4).Enabled=   0   'False
            Tab(3).Control(5)=   "Label35"
            Tab(3).Control(5).Enabled=   0   'False
            Tab(3).Control(6)=   "lblItemName4"
            Tab(3).Control(6).Enabled=   0   'False
            Tab(3).Control(7)=   "Label33"
            Tab(3).Control(7).Enabled=   0   'False
            Tab(3).Control(8)=   "Label32"
            Tab(3).Control(8).Enabled=   0   'False
            Tab(3).Control(9)=   "lblNum4"
            Tab(3).Control(9).Enabled=   0   'False
            Tab(3).ControlCount=   10
            Begin VB.TextBox txtChance4 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   -72480
               TabIndex        =   50
               Text            =   "0"
               Top             =   1680
               Width           =   1815
            End
            Begin VB.HScrollBar scrlValue4 
               Height          =   255
               Left            =   -74280
               Max             =   1000
               TabIndex        =   49
               Top             =   1200
               Value           =   1
               Width           =   3375
            End
            Begin VB.HScrollBar scrlNum4 
               Height          =   255
               Left            =   -74280
               Max             =   255
               TabIndex        =   48
               Top             =   720
               Value           =   1
               Width           =   3375
            End
            Begin VB.TextBox txtChance3 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   -72480
               TabIndex        =   40
               Text            =   "0"
               Top             =   1680
               Width           =   1815
            End
            Begin VB.HScrollBar scrlValue3 
               Height          =   255
               Left            =   -74280
               Max             =   1000
               TabIndex        =   39
               Top             =   1200
               Value           =   1
               Width           =   3375
            End
            Begin VB.HScrollBar scrlNum3 
               Height          =   255
               Left            =   -74280
               Max             =   255
               TabIndex        =   38
               Top             =   720
               Value           =   1
               Width           =   3375
            End
            Begin VB.TextBox txtChance2 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   -72480
               TabIndex        =   30
               Text            =   "0"
               Top             =   1680
               Width           =   1815
            End
            Begin VB.HScrollBar scrlValue2 
               Height          =   255
               Left            =   -74280
               Max             =   1000
               TabIndex        =   29
               Top             =   1200
               Value           =   1
               Width           =   3375
            End
            Begin VB.HScrollBar scrlNum2 
               Height          =   255
               Left            =   -74280
               Max             =   255
               TabIndex        =   28
               Top             =   720
               Value           =   1
               Width           =   3375
            End
            Begin VB.TextBox txtChance 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   2520
               TabIndex        =   20
               Text            =   "0"
               Top             =   1680
               Width           =   1815
            End
            Begin VB.HScrollBar scrlValue 
               Height          =   255
               Left            =   720
               Max             =   1000
               TabIndex        =   19
               Top             =   1200
               Value           =   1
               Width           =   3375
            End
            Begin VB.HScrollBar scrlNum 
               Height          =   255
               Left            =   720
               Max             =   255
               TabIndex        =   18
               Top             =   720
               Value           =   1
               Width           =   3375
            End
            Begin VB.Label Label37 
               Alignment       =   1  'Right Justify
               Caption         =   "Drop Item Chance (%):"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   -74760
               TabIndex        =   57
               Top             =   1680
               Width           =   1935
            End
            Begin VB.Label lblValue4 
               Alignment       =   1  'Right Justify
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   -70920
               TabIndex        =   56
               Top             =   1200
               Width           =   495
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               Caption         =   "Value:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   -74880
               TabIndex        =   55
               Top             =   1200
               Width           =   495
            End
            Begin VB.Label lblItemName4 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   -74280
               TabIndex        =   54
               Top             =   360
               Width           =   3855
            End
            Begin VB.Label Label33 
               Alignment       =   1  'Right Justify
               Caption         =   "Item:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   -74760
               TabIndex        =   53
               Top             =   360
               Width           =   375
            End
            Begin VB.Label Label32 
               Alignment       =   1  'Right Justify
               Caption         =   "Num:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   -74880
               TabIndex        =   52
               Top             =   720
               Width           =   495
            End
            Begin VB.Label lblNum4 
               Alignment       =   1  'Right Justify
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   -70920
               TabIndex        =   51
               Top             =   720
               Width           =   495
            End
            Begin VB.Label Label29 
               Alignment       =   1  'Right Justify
               Caption         =   "Drop Item Chance (%):"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   -74760
               TabIndex        =   47
               Top             =   1680
               Width           =   1935
            End
            Begin VB.Label lblValue3 
               Alignment       =   1  'Right Justify
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   -70920
               TabIndex        =   46
               Top             =   1200
               Width           =   495
            End
            Begin VB.Label Label27 
               Alignment       =   1  'Right Justify
               Caption         =   "Value:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   -74880
               TabIndex        =   45
               Top             =   1200
               Width           =   495
            End
            Begin VB.Label lblItemName3 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   -74280
               TabIndex        =   44
               Top             =   360
               Width           =   3855
            End
            Begin VB.Label Label25 
               Alignment       =   1  'Right Justify
               Caption         =   "Item:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   -74760
               TabIndex        =   43
               Top             =   360
               Width           =   375
            End
            Begin VB.Label Label24 
               Alignment       =   1  'Right Justify
               Caption         =   "Num:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   -74880
               TabIndex        =   42
               Top             =   720
               Width           =   495
            End
            Begin VB.Label lblNum3 
               Alignment       =   1  'Right Justify
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   -70920
               TabIndex        =   41
               Top             =   720
               Width           =   495
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               Caption         =   "Drop Item Chance (%):"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   -74760
               TabIndex        =   37
               Top             =   1680
               Width           =   1935
            End
            Begin VB.Label lblValue2 
               Alignment       =   1  'Right Justify
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   -70920
               TabIndex        =   36
               Top             =   1200
               Width           =   495
            End
            Begin VB.Label Label19 
               Alignment       =   1  'Right Justify
               Caption         =   "Value:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   -74880
               TabIndex        =   35
               Top             =   1200
               Width           =   495
            End
            Begin VB.Label lblItemName2 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   -74280
               TabIndex        =   34
               Top             =   360
               Width           =   3855
            End
            Begin VB.Label Label17 
               Alignment       =   1  'Right Justify
               Caption         =   "Item:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   -74760
               TabIndex        =   33
               Top             =   360
               Width           =   375
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               Caption         =   "Num:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   -74880
               TabIndex        =   32
               Top             =   720
               Width           =   495
            End
            Begin VB.Label lblNum2 
               Alignment       =   1  'Right Justify
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   -70920
               TabIndex        =   31
               Top             =   720
               Width           =   495
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               Caption         =   "Drop Item Chance (%):"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   27
               Top             =   1680
               Width           =   1935
            End
            Begin VB.Label lblValue 
               Alignment       =   1  'Right Justify
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4080
               TabIndex        =   26
               Top             =   1200
               Width           =   495
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               Caption         =   "Value:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   25
               Top             =   1200
               Width           =   495
            End
            Begin VB.Label lblItemName 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   720
               TabIndex        =   24
               Top             =   340
               Width           =   3855
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               Caption         =   "Item:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   225
               TabIndex        =   23
               Top             =   360
               Width           =   375
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               Caption         =   "Num:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   22
               Top             =   720
               Width           =   495
            End
            Begin VB.Label lblNum 
               Alignment       =   1  'Right Justify
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4080
               TabIndex        =   21
               Top             =   720
               Width           =   495
            End
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "Level"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   720
            TabIndex        =   93
            Top             =   1560
            Width           =   495
         End
         Begin VB.Label lblLevel 
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4080
            TabIndex        =   92
            Top             =   1560
            Width           =   495
         End
         Begin VB.Label lblStat 
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   4080
            TabIndex        =   87
            Top             =   240
            Width           =   375
         End
         Begin VB.Label lblStatName 
            Alignment       =   1  'Right Justify
            Caption         =   "Stat Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   86
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lblStat 
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   4080
            TabIndex        =   85
            Top             =   480
            Width           =   375
         End
         Begin VB.Label lblStatName 
            Alignment       =   1  'Right Justify
            Caption         =   "Stat Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   84
            Top             =   480
            Width           =   855
         End
         Begin VB.Label lblStat 
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   4080
            TabIndex        =   83
            Top             =   720
            Width           =   375
         End
         Begin VB.Label lblStatName 
            Alignment       =   1  'Right Justify
            Caption         =   "Stat Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   82
            Top             =   720
            Width           =   855
         End
         Begin VB.Label lblStat 
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   4080
            TabIndex        =   81
            Top             =   960
            Width           =   375
         End
         Begin VB.Label lblStatName 
            Alignment       =   1  'Right Justify
            Caption         =   "Stat Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   80
            Top             =   960
            Width           =   855
         End
         Begin VB.Label lblStatName 
            Alignment       =   1  'Right Justify
            Caption         =   "Stat Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   79
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label lblStat 
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   4080
            TabIndex        =   78
            Top             =   1200
            Width           =   375
         End
         Begin VB.Label Label4 
            Caption         =   "Range:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   450
            TabIndex        =   66
            Top             =   2820
            Width           =   585
         End
         Begin VB.Label lblRange 
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4080
            TabIndex        =   65
            Top             =   2760
            Width           =   495
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "HP:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   840
            TabIndex        =   64
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "EXP:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   855
            TabIndex        =   63
            Top             =   2040
            Width           =   330
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "Spawn Rate (in seconds):"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   59
            Top             =   2400
            Width           =   2055
         End
      End
   End
End
Attribute VB_Name = "frmNpcEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function NpcLevel() As Long
Dim i As Long
Dim StatCount As Long

    For i = 1 To Stats.Stat_Count - 1
        StatCount = StatCount + scrlStat(i - 1).Value
    Next
    NpcLevel = Clamp(StatCount \ 3.75, 0, scrlLevel.Max)
End Function

Private Sub cmdCalc_Click()
Dim MaxEXP As Long
Dim MaxHP As Long
Dim tempLevel As Long
Dim tempHp As Long
Dim Result As Boolean
    
    Result = MsgBox("This will change the Level, HP, and EXP. Do you want to continue?", vbOKCancel)
    
    If Result Then
        scrlLevel.Value = NpcLevel
        
        tempLevel = (scrlLevel.Value * 3) * 5
        tempHp = ((scrlStat(Stats.Strength - 1).Value \ 2) + scrlStat(Stats.Vitality - 1).Value) + (scrlLevel.Value * 5)
        
        MaxHP = Rand(tempHp * 0.9, tempHp * 1.1)
        MaxEXP = Rand(tempLevel * 0.9, tempLevel * 1.1)
        
        txtHP.Text = MaxHP
        txtEXP.Text = MaxEXP
    End If
End Sub

Private Sub scrlLevel_Change()
    lblLevel.Caption = scrlLevel.Value
End Sub

Private Sub scrlStat_Change(Index As Integer)
    lblStat(Index) = scrlStat(Index).Value
End Sub

Private Sub cmbBehavior_Click()
    If (cmbBehavior.ListIndex = NPC_BEHAVIOR_SHOPKEEPER) Then
        fraShopKeeper.Visible = True
        fraNpc.Visible = False
    Else
        fraShopKeeper.Visible = False
        fraNpc.Visible = True
    End If
End Sub

Private Sub scrlValue_Change()
    lblValue.Caption = CStr(scrlValue.Value)
End Sub

Private Sub scrlValue2_Change()
    lblValue2.Caption = Str$(scrlValue2.Value)
End Sub

Private Sub scrlValue3_Change()
    lblValue3.Caption = Str$(scrlValue3.Value)
End Sub

Private Sub scrlValue4_Change()
    lblValue4.Caption = Str$(scrlValue4.Value)
End Sub


Private Sub scrlNum_Change()
    lblNum.Caption = scrlNum.Value
    If scrlNum.Value > 0 Then
        lblItemName.Caption = Trim$(Item(scrlNum.Value).Name)
    End If
End Sub

Private Sub scrlNum2_Change()
    lblNum2.Caption = Str$(scrlNum2.Value)
    If scrlNum2.Value > 0 Then
        lblItemName2.Caption = Trim$(Item(scrlNum2.Value).Name)
    End If
End Sub

Private Sub scrlNum3_Change()
    lblNum3.Caption = scrlNum3.Value
    If scrlNum3.Value > 0 Then
        lblItemName3.Caption = Trim$(Item(scrlNum3.Value).Name)
    End If
End Sub

Private Sub scrlNum4_Change()
    lblNum4.Caption = scrlNum4.Value
    If scrlNum4.Value > 0 Then
        lblItemName4.Caption = Trim$(Item(scrlNum4.Value).Name)
    End If
End Sub



Private Sub scrlSprite_Change()
    NpcEditorBltSprite
    lblSprite.Caption = scrlSprite.Value
End Sub

Private Sub scrlRange_Change()
    lblRange.Caption = scrlRange.Value
End Sub

Private Sub cmdOk_Click()
    Call NpcEditorOk
End Sub

Private Sub cmdCancel_Click()
    Call NpcEditorCancel
End Sub

Private Sub scrlMovementSpeed_Change()
    lblMovementSpeed.Caption = scrlMovementSpeed.Value
End Sub

Private Sub NpcEditorBltSprite()
Dim rec As RECT
Dim drec As RECT
    
    With rec
        .Top = scrlSprite.Value * PIC_Y
        .Bottom = .Top + PIC_Y
        .Left = 4 * PIC_X
        .Right = .Left + PIC_X
    End With
    
    With drec
        .Top = 0
        .Bottom = PIC_Y
        .Left = 0
        .Right = PIC_X
    End With
    
    DD_SpriteSurf.BltToDC picSprite.hdc, rec, drec
    picSprite.Refresh
End Sub

Private Sub txtEXP_Change()
    If Not IsNumeric(txtEXP.Text) Then txtEXP.Text = Npc(EditorIndex).MaxEXP
    If txtEXP.Text > MAX_LONG \ 2 Then txtEXP.Text = MAX_LONG \ 2
    If txtEXP.Text < 0 Then txtEXP.Text = 0
End Sub

Private Sub txtHP_Change()
    If Not IsNumeric(txtHP.Text) Then txtHP.Text = Npc(EditorIndex).MaxHP
    If txtHP.Text > MAX_LONG \ 2 Then txtHP.Text = MAX_LONG \ 2
    If txtHP.Text < 0 Then txtHP.Text = 0
End Sub
