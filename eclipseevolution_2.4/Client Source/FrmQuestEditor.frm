VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmQuestEditor 
   Caption         =   "Edit Quest"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11070
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   5535
   ScaleWidth      =   11070
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5355
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10845
      _ExtentX        =   19129
      _ExtentY        =   9446
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
      TabCaption(0)   =   "Edit Quest"
      TabPicture(0)   =   "FrmQuestEditor.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdOk"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Command1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Name"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.Frame Frame2 
         Caption         =   "Location"
         Height          =   2055
         Left            =   240
         TabIndex        =   44
         Top             =   3120
         Width           =   4215
         Begin VB.HScrollBar HScroll11 
            Height          =   255
            Left            =   2280
            TabIndex        =   54
            Top             =   1440
            Width           =   1575
         End
         Begin VB.HScrollBar HScroll10 
            Height          =   255
            Left            =   2280
            TabIndex        =   51
            Top             =   720
            Width           =   1575
         End
         Begin VB.HScrollBar HScroll9 
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   1440
            Width           =   1815
         End
         Begin VB.HScrollBar HScroll8 
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Y:"
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
            TabIndex        =   56
            Top             =   1200
            Width           =   135
         End
         Begin VB.Label Label28 
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
            Left            =   3960
            TabIndex        =   55
            Top             =   1440
            Width           =   75
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Npc:"
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
            TabIndex        =   53
            Top             =   480
            Width           =   285
         End
         Begin VB.Label Label24 
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
            Left            =   3960
            TabIndex        =   52
            Top             =   720
            Width           =   75
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "X:"
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
            TabIndex        =   50
            Top             =   1200
            Width           =   120
         End
         Begin VB.Label Label20 
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
            Left            =   2040
            TabIndex        =   49
            Top             =   1440
            Width           =   315
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Map:"
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
            TabIndex        =   47
            Top             =   480
            Width           =   315
         End
         Begin VB.Label Label3 
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
            Left            =   2040
            TabIndex        =   46
            Top             =   720
            Width           =   315
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Quest stages"
         Height          =   4335
         Left            =   4560
         TabIndex        =   17
         Top             =   480
         Width           =   6135
         Begin VB.HScrollBar HScroll12 
            Height          =   255
            Left            =   3480
            TabIndex        =   57
            Top             =   3480
            Width           =   2175
         End
         Begin VB.ComboBox cmbItem 
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
            Index           =   1
            ItemData        =   "FrmQuestEditor.frx":001C
            Left            =   360
            List            =   "FrmQuestEditor.frx":001E
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   600
            Width           =   2175
         End
         Begin VB.HScrollBar HScroll3 
            Height          =   255
            Left            =   360
            TabIndex        =   28
            Top             =   1200
            Width           =   2175
         End
         Begin VB.ComboBox cmbItem 
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
            Index           =   2
            ItemData        =   "FrmQuestEditor.frx":0020
            Left            =   360
            List            =   "FrmQuestEditor.frx":0022
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   2040
            Width           =   2175
         End
         Begin VB.HScrollBar HScroll4 
            Height          =   255
            Left            =   360
            TabIndex        =   26
            Top             =   2640
            Width           =   2175
         End
         Begin VB.ComboBox cmbItem 
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
            Index           =   3
            ItemData        =   "FrmQuestEditor.frx":0024
            Left            =   3480
            List            =   "FrmQuestEditor.frx":0026
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   600
            Width           =   2175
         End
         Begin VB.HScrollBar HScroll5 
            Height          =   255
            Left            =   3480
            TabIndex        =   24
            Top             =   1200
            Width           =   2175
         End
         Begin VB.ComboBox cmbItem 
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
            Index           =   4
            ItemData        =   "FrmQuestEditor.frx":0028
            Left            =   3480
            List            =   "FrmQuestEditor.frx":002A
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   2040
            Width           =   2175
         End
         Begin VB.HScrollBar HScroll6 
            Height          =   255
            Left            =   3480
            TabIndex        =   22
            Top             =   2640
            Width           =   2175
         End
         Begin VB.HScrollBar HScroll7 
            Height          =   255
            Left            =   360
            TabIndex        =   21
            Top             =   3480
            Width           =   2175
         End
         Begin VB.CommandButton Command3 
            Caption         =   ">"
            Height          =   255
            Left            =   5640
            TabIndex        =   19
            Top             =   3960
            Width           =   255
         End
         Begin VB.CommandButton Command2 
            Caption         =   "<"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   3960
            Width           =   255
         End
         Begin VB.Label Label31 
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
            Left            =   5760
            TabIndex        =   59
            Top             =   3480
            Width           =   315
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Script to run :"
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
            Left            =   3480
            TabIndex        =   58
            Top             =   3240
            Width           =   840
         End
         Begin VB.Label Label2 
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
            Left            =   5760
            TabIndex        =   43
            Top             =   2640
            Width           =   315
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Item to take :"
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
            Left            =   360
            TabIndex        =   42
            Top             =   360
            Width           =   870
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Number to take :"
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
            Left            =   360
            TabIndex        =   41
            Top             =   960
            Width           =   1065
         End
         Begin VB.Label Label11 
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
            Left            =   2640
            TabIndex        =   40
            Top             =   1200
            Width           =   315
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Item to take :"
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
            Left            =   360
            TabIndex        =   39
            Top             =   1800
            Width           =   870
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Number to take :"
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
            Left            =   360
            TabIndex        =   38
            Top             =   2400
            Width           =   1065
         End
         Begin VB.Label Label14 
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
            Left            =   2640
            TabIndex        =   37
            Top             =   2640
            Width           =   315
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Item to give :"
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
            Left            =   3480
            TabIndex        =   36
            Top             =   360
            Width           =   870
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Number to give :"
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
            Left            =   3480
            TabIndex        =   35
            Top             =   960
            Width           =   1065
         End
         Begin VB.Label Label17 
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
            Left            =   5760
            TabIndex        =   34
            Top             =   1200
            Width           =   315
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Item to give :"
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
            Left            =   3480
            TabIndex        =   33
            Top             =   1800
            Width           =   870
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Number to give :"
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
            Left            =   3480
            TabIndex        =   32
            Top             =   2400
            Width           =   1065
         End
         Begin VB.Label Label25 
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
            Left            =   2640
            TabIndex        =   31
            Top             =   3480
            Width           =   315
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Exp to give :"
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
            Left            =   360
            TabIndex        =   30
            Top             =   3240
            Width           =   795
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Stage 1"
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
            Left            =   2880
            TabIndex        =   20
            Top             =   4080
            Width           =   480
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Icon"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Width           =   4215
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   540
            Left            =   240
            ScaleHeight     =   34
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   34
            TabIndex        =   7
            Top             =   600
            Width           =   540
            Begin VB.PictureBox picEmoticon 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   480
               Left            =   15
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   8
               Top             =   15
               Width           =   480
               Begin VB.PictureBox iconn 
                  AutoRedraw      =   -1  'True
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  Height          =   480
                  Left            =   0
                  ScaleHeight     =   32
                  ScaleMode       =   3  'Pixel
                  ScaleWidth      =   32
                  TabIndex        =   9
                  Top             =   0
                  Width           =   480
               End
            End
         End
         Begin VB.HScrollBar HScroll2 
            Height          =   255
            Left            =   1200
            TabIndex        =   6
            Top             =   1080
            Width           =   2415
         End
         Begin VB.HScrollBar HScroll1 
            Height          =   255
            Left            =   1200
            TabIndex        =   5
            Top             =   480
            Width           =   2415
         End
         Begin VB.Label Label9 
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
            Left            =   3720
            TabIndex        =   14
            Top             =   1080
            Width           =   315
         End
         Begin VB.Label Label8 
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
            Left            =   3720
            TabIndex        =   13
            Top             =   480
            Width           =   315
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Left:"
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
            TabIndex        =   12
            Top             =   840
            Width           =   270
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Icon:"
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
            Top             =   360
            Width           =   315
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Top:"
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
            TabIndex        =   10
            Top             =   240
            Width           =   285
         End
      End
      Begin VB.Frame Name 
         Caption         =   "Name"
         Height          =   1095
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   4215
         Begin VB.TextBox txtName 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   600
            TabIndex        =   15
            Top             =   480
            Width           =   2775
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Height          =   165
            Left            =   600
            TabIndex        =   16
            Top             =   240
            Width           =   420
         End
      End
      Begin VB.CommandButton Command1 
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
         Left            =   9480
         TabIndex        =   2
         Top             =   4920
         Width           =   1215
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
         TabIndex        =   1
         Top             =   4920
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FrmQuestEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

    Call QuestEditorCancel

End Sub

Private Sub cmdOk_Click()

    Call savesheet
    Call QuestEditorOk

End Sub

Private Sub Command2_Click()

    If currentsheet > 0 Then
        Call savesheet
        currentsheet = currentsheet - 1
        Call loadsheet
    End If

End Sub

Private Sub Command3_Click()

    If currentsheet < MAX_QUEST_LENGHT Then
        Call savesheet
        currentsheet = currentsheet + 1
        Call loadsheet
    End If

End Sub

Private Sub Form_Load()

  Dim i As Long
  Dim Ending As String

    For i = 1 To 5
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"
        If i = 4 Then Ending = ".bmp"

        If FileExist("GFX\Icons" & Ending) Then iconn.Picture = LoadPicture(App.Path & "\GFX\Icons" & Ending)
    Next i

    iconn.Left = -val(HScroll2.Value * PIC_X)
    iconn.Top = -val(HScroll1.Value * PIC_Y)

End Sub

Private Sub HScroll10_Change()

    Label24.Caption = HScroll10.Value

End Sub

Private Sub HScroll11_Change()

    Label28.Caption = HScroll11.Value

End Sub

Private Sub HScroll12_Change()

    Label31.Caption = HScroll12.Value

End Sub

Private Sub HScroll1_Change()

    iconn.Left = -val(HScroll2.Value * PIC_X)
    iconn.Top = -val(HScroll1.Value * PIC_Y)
    Label8.Caption = HScroll1.Value

End Sub

Private Sub HScroll2_Change()

    iconn.Left = -val(HScroll2.Value * PIC_X)
    iconn.Top = -val(HScroll1.Value * PIC_Y)
    Label9.Caption = HScroll2.Value

End Sub

Private Sub HScroll3_Change()

    Label11.Caption = HScroll3.Value

End Sub

Private Sub HScroll4_Change()

    Label14.Caption = HScroll4.Value

End Sub

Private Sub HScroll5_Change()

    Label17.Caption = HScroll5.Value

End Sub

Private Sub HScroll6_Change()

    Label2.Caption = HScroll6.Value

End Sub

Private Sub HScroll7_Change()

    Label25.Caption = HScroll7.Value

End Sub

Private Sub HScroll8_Change()

    Label3.Caption = HScroll8.Value

End Sub

Private Sub HScroll9_Change()

    Label20.Caption = HScroll9.Value

End Sub

Sub loadsheet()

    HScroll8.Value = Q_Map(currentsheet)
    HScroll9.Value = Q_X(currentsheet)
    HScroll10.Value = Q_Npc(currentsheet)
    HScroll11.Value = Q_Y(currentsheet)
    HScroll12.Value = Q_Script(currentsheet)
    HScroll3.Value = Q_ItemTake1val(currentsheet)
    HScroll4.Value = Q_ItemTake2val(currentsheet)
    HScroll5.Value = Q_ItemGive1val(currentsheet)
    HScroll6.Value = Q_ItemGive2val(currentsheet)
    HScroll7.Value = Q_ExpGiven(currentsheet)

    Label11.Caption = HScroll3.Value
    Label14.Caption = HScroll4.Value
    Label17.Caption = HScroll5.Value
    Label2.Caption = HScroll6.Value
    Label25.Caption = HScroll7.Value
    Label3.Caption = HScroll8.Value
    Label20.Caption = HScroll9.Value
    Label24.Caption = HScroll10.Value
    Label28.Caption = HScroll11.Value
    Label31.Caption = HScroll12.Value

    If val(Q_ItemTake1num(currentsheet)) <> 0 Then
        cmbItem(1).ListIndex = Q_ItemTake1num(currentsheet)
     Else
        cmbItem(1).ListIndex = 0
    End If

    If val(Q_ItemTake2num(currentsheet)) <> 0 Then
        cmbItem(2).ListIndex = Q_ItemTake2num(currentsheet)
     Else
        cmbItem(2).ListIndex = 0
    End If

    If val(Q_ItemGive1num(currentsheet)) <> 0 Then
        cmbItem(3).ListIndex = Q_ItemGive1num(currentsheet)
     Else
        cmbItem(3).ListIndex = 0
    End If

    If val(Q_ItemGive2num(currentsheet)) <> 0 Then
        cmbItem(4).ListIndex = Q_ItemGive2num(currentsheet)
     Else
        cmbItem(4).ListIndex = 0
    End If

End Sub

Sub savesheet()

    Q_Map(currentsheet) = HScroll8.Value
    Q_X(currentsheet) = HScroll9.Value
    Q_Npc(currentsheet) = HScroll10.Value
    Q_Y(currentsheet) = HScroll11.Value
    Q_Script(currentsheet) = HScroll12.Value
    Q_ItemTake1num(currentsheet) = cmbItem(1).ListIndex
    Q_ItemTake2num(currentsheet) = cmbItem(2).ListIndex
    Q_ItemGive1num(currentsheet) = cmbItem(3).ListIndex
    Q_ItemGive2num(currentsheet) = cmbItem(4).ListIndex
    Q_ItemTake1val(currentsheet) = HScroll3.Value
    Q_ItemTake2val(currentsheet) = HScroll4.Value
    Q_ItemGive1val(currentsheet) = HScroll5.Value
    Q_ItemGive2val(currentsheet) = HScroll6.Value
    Q_ExpGiven(currentsheet) = HScroll7.Value

End Sub

