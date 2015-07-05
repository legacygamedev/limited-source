VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmSkillEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit skill"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12015
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   12015
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   9551
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   397
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
      TabCaption(0)   =   "Edit skill"
      TabPicture(0)   =   "frmSkillEditor.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label30"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdOk"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdCancel"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.Frame Frame4 
         Caption         =   "Items"
         Height          =   4455
         Left            =   5760
         TabIndex        =   34
         Top             =   480
         Width           =   5895
         Begin VB.HScrollBar HScroll8 
            Height          =   255
            Left            =   3240
            TabIndex        =   54
            Top             =   3720
            Width           =   2175
         End
         Begin VB.HScrollBar HScroll7 
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   3720
            Width           =   2175
         End
         Begin VB.ComboBox cmbLevel 
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
            ItemData        =   "frmSkillEditor.frx":001C
            Left            =   3240
            List            =   "frmSkillEditor.frx":001E
            Style           =   2  'Dropdown List
            TabIndex        =   49
            Top             =   480
            Width           =   2175
         End
         Begin VB.CommandButton Command2 
            Caption         =   "<"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   4080
            Width           =   255
         End
         Begin VB.CommandButton Command1 
            Caption         =   ">"
            Height          =   255
            Left            =   5520
            TabIndex        =   16
            Top             =   4080
            Width           =   255
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
            Index           =   0
            ItemData        =   "frmSkillEditor.frx":0020
            Left            =   120
            List            =   "frmSkillEditor.frx":0022
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   480
            Width           =   2175
         End
         Begin VB.HScrollBar HScroll6 
            Height          =   255
            Left            =   3240
            TabIndex        =   14
            Top             =   3120
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
            ItemData        =   "frmSkillEditor.frx":0024
            Left            =   3240
            List            =   "frmSkillEditor.frx":0026
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   2520
            Width           =   2175
         End
         Begin VB.HScrollBar HScroll5 
            Height          =   255
            Left            =   3240
            TabIndex        =   10
            Top             =   1800
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
            ItemData        =   "frmSkillEditor.frx":0028
            Left            =   3240
            List            =   "frmSkillEditor.frx":002A
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1200
            Width           =   2175
         End
         Begin VB.HScrollBar HScroll4 
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   3120
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
            ItemData        =   "frmSkillEditor.frx":002C
            Left            =   120
            List            =   "frmSkillEditor.frx":002E
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   2520
            Width           =   2175
         End
         Begin VB.HScrollBar HScroll3 
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   1800
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
            ItemData        =   "frmSkillEditor.frx":0030
            Left            =   120
            List            =   "frmSkillEditor.frx":0032
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1200
            Width           =   2175
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
            Left            =   5520
            TabIndex        =   56
            Top             =   3720
            Width           =   315
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Base chance :"
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
            TabIndex        =   55
            Top             =   3480
            Width           =   855
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
            Left            =   120
            TabIndex        =   53
            Top             =   3480
            Width           =   795
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
            Left            =   2400
            TabIndex        =   52
            Top             =   3720
            Width           =   315
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Level required:"
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
            TabIndex        =   50
            Top             =   240
            Width           =   915
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Sheet 1"
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
            TabIndex        =   48
            Top             =   4200
            Width           =   480
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Item equiped:"
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
            Top             =   240
            Width           =   870
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
            Left            =   5520
            TabIndex        =   46
            Top             =   3120
            Width           =   315
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
            Left            =   3240
            TabIndex        =   45
            Top             =   2880
            Width           =   1065
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
            Left            =   3240
            TabIndex        =   44
            Top             =   2280
            Width           =   870
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
            Left            =   5520
            TabIndex        =   43
            Top             =   1800
            Width           =   315
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
            Left            =   3240
            TabIndex        =   42
            Top             =   1560
            Width           =   1065
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
            Left            =   3240
            TabIndex        =   41
            Top             =   960
            Width           =   870
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
            Left            =   2400
            TabIndex        =   40
            Top             =   3120
            Width           =   315
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
            Left            =   120
            TabIndex        =   39
            Top             =   2880
            Width           =   1065
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
            Left            =   120
            TabIndex        =   38
            Top             =   2280
            Width           =   870
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
            Left            =   2400
            TabIndex        =   37
            Top             =   1800
            Width           =   315
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
            Left            =   120
            TabIndex        =   36
            Top             =   1560
            Width           =   1065
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
            Left            =   120
            TabIndex        =   35
            Top             =   960
            Width           =   870
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
         Height          =   2175
         Left            =   120
         TabIndex        =   25
         Top             =   2760
         Width           =   5535
         Begin VB.HScrollBar HScroll1 
            Height          =   255
            Left            =   1560
            TabIndex        =   4
            Top             =   1320
            Width           =   2415
         End
         Begin VB.HScrollBar HScroll2 
            Height          =   255
            Left            =   1560
            TabIndex        =   5
            Top             =   720
            Width           =   2415
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   540
            Left            =   240
            ScaleHeight     =   34
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   34
            TabIndex        =   26
            Top             =   840
            Width           =   540
            Begin VB.PictureBox picEmoticon 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   480
               Left            =   15
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   27
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
                  TabIndex        =   28
                  Top             =   0
                  Width           =   480
               End
            End
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
            Left            =   1560
            TabIndex        =   33
            Top             =   480
            Width           =   285
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
            TabIndex        =   32
            Top             =   600
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
            Left            =   1560
            TabIndex        =   31
            Top             =   1080
            Width           =   270
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
            Left            =   4080
            TabIndex        =   30
            Top             =   720
            Width           =   315
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
            Left            =   4080
            TabIndex        =   29
            Top             =   1320
            Width           =   315
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Messages"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   5535
         Begin VB.TextBox txtAttempt 
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
            Left            =   2880
            MaxLength       =   50
            TabIndex        =   57
            Top             =   1080
            Width           =   2415
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
            Height          =   285
            Left            =   120
            MaxLength       =   50
            TabIndex        =   0
            Top             =   480
            Width           =   2415
         End
         Begin VB.TextBox txtAction 
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
            Left            =   2880
            MaxLength       =   50
            TabIndex        =   1
            Top             =   480
            Width           =   2415
         End
         Begin VB.TextBox txtFail 
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
            Left            =   120
            MaxLength       =   50
            TabIndex        =   3
            Top             =   1680
            Width           =   2415
         End
         Begin VB.TextBox txtSucces 
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
            Left            =   120
            MaxLength       =   50
            TabIndex        =   2
            Top             =   1080
            Width           =   2415
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "On Attempt Message:"
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
            TabIndex        =   58
            Top             =   840
            Width           =   1380
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   420
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Action:"
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
            TabIndex        =   23
            Top             =   240
            Width           =   435
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "On fail message:"
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
            TabIndex        =   22
            Top             =   1440
            Width           =   1050
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "On Success message:"
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
            TabIndex        =   21
            Top             =   840
            Width           =   1350
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
         Left            =   10320
         TabIndex        =   18
         Top             =   5040
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
         Left            =   9000
         TabIndex        =   17
         Top             =   5040
         Width           =   1215
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "NOTE: If you edit a skill, you must re-place all your skill tiles to make the changes. Keep this in mind when making your skills!"
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
         TabIndex        =   59
         Top             =   5040
         Width           =   7875
      End
   End
End
Attribute VB_Name = "frmskilleditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Sub cleanform()

    If 0 + itemequiped(currentsheet) <> 0 Then
        frmskilleditor.cmbItem(0).ListIndex = 0
        frmskilleditor.cmbItem(1).ListIndex = 0
        frmskilleditor.cmbItem(2).ListIndex = 0
        frmskilleditor.cmbItem(3).ListIndex = 0
        frmskilleditor.cmbItem(4).ListIndex = 0
        frmskilleditor.cmbLevel.ListIndex = 0
        frmskilleditor.HScroll3.Value = 0
        frmskilleditor.HScroll4.Value = 0
        frmskilleditor.HScroll5.Value = 0
        frmskilleditor.HScroll6.Value = 0
        frmskilleditor.HScroll7.Value = 0
        frmskilleditor.HScroll8.Value = 0
        frmskilleditor.Label11.Caption = 0
        frmskilleditor.Label14.Caption = 0
        frmskilleditor.Label17.Caption = 0
        frmskilleditor.Label20.Caption = 0
        frmskilleditor.Label25.Caption = 0
        frmskilleditor.Label28.Caption = 0
    End If

End Sub

Private Sub cmdCancel_Click()

    Call SkillEditorCancel

End Sub

Private Sub cmdOk_Click()

    Call savesheet
    Call SkillEditorOk

End Sub

Private Sub Command1_Click()

    Call savesheet
    If currentsheet = MAX_SKILLS_SHEETS Then Exit Sub
    Call cleanform
    currentsheet = currentsheet + 1

    If itemequiped(currentsheet) <> 0 Then
        frmskilleditor.cmbItem(0).ListIndex = itemequiped(currentsheet)
     Else
        frmskilleditor.cmbItem(0).ListIndex = 1
    End If

    If ItemTake1num(currentsheet) <> 0 Then
        frmskilleditor.cmbItem(1).ListIndex = ItemTake1num(currentsheet)
     Else
        frmskilleditor.cmbItem(1).ListIndex = 1
    End If

    If ItemTake2num(currentsheet) <> 0 Then
        frmskilleditor.cmbItem(2).ListIndex = ItemTake2num(currentsheet)
     Else
        frmskilleditor.cmbItem(2).ListIndex = 1
    End If

    If ItemGive1num(currentsheet) <> 0 Then
        frmskilleditor.cmbItem(3).ListIndex = ItemGive1num(currentsheet)
     Else
        frmskilleditor.cmbItem(3).ListIndex = 1
    End If

    If ItemGive2num(currentsheet) <> 0 Then
        frmskilleditor.cmbItem(4).ListIndex = ItemGive2num(currentsheet)
     Else
        frmskilleditor.cmbItem(4).ListIndex = 1
    End If

    frmskilleditor.HScroll3.Value = ItemTake1val(currentsheet)
    frmskilleditor.HScroll4.Value = ItemTake2val(currentsheet)
    frmskilleditor.HScroll5.Value = ItemGive1val(currentsheet)
    frmskilleditor.HScroll6.Value = ItemGive2val(currentsheet)
    frmskilleditor.HScroll7.Value = ExpGiven(currentsheet)
    frmskilleditor.HScroll8.Value = base_chance(currentsheet)
    frmskilleditor.Label11.Caption = ItemTake1val(currentsheet)
    frmskilleditor.Label14.Caption = ItemTake2val(currentsheet)
    frmskilleditor.Label17.Caption = ItemGive1val(currentsheet)
    frmskilleditor.Label20.Caption = ItemGive2val(currentsheet)
    frmskilleditor.Label25.Caption = ExpGiven(currentsheet)
    frmskilleditor.Label28.Caption = base_chance(currentsheet)
    frmskilleditor.Label10.Caption = "Sheet " & val(currentsheet + 1)

End Sub

Private Sub Command2_Click()

    Call savesheet
    If currentsheet = 1 Then Exit Sub
    Call cleanform
    currentsheet = currentsheet - 1

    If itemequiped(currentsheet) <> 0 Then frmskilleditor.cmbItem(0).ListIndex = itemequiped(currentsheet)
    If ItemTake1num(currentsheet) <> 0 Then frmskilleditor.cmbItem(1).ListIndex = ItemTake1num(currentsheet)
    If ItemTake2num(currentsheet) <> 0 Then frmskilleditor.cmbItem(2).ListIndex = ItemTake2num(currentsheet)
    If ItemGive1num(currentsheet) <> 0 Then frmskilleditor.cmbItem(3).ListIndex = ItemGive1num(currentsheet)
    If ItemGive2num(currentsheet) <> 0 Then frmskilleditor.cmbItem(4).ListIndex = ItemGive2num(currentsheet)

    frmskilleditor.HScroll3.Value = ItemTake1val(currentsheet)
    frmskilleditor.HScroll4.Value = ItemTake2val(currentsheet)
    frmskilleditor.HScroll5.Value = ItemGive1val(currentsheet)
    frmskilleditor.HScroll6.Value = ItemGive2val(currentsheet)
    frmskilleditor.HScroll7.Value = ExpGiven(currentsheet)
    frmskilleditor.HScroll8.Value = base_chance(currentsheet)
    frmskilleditor.Label11.Caption = ItemTake1val(currentsheet)
    frmskilleditor.Label14.Caption = ItemTake2val(currentsheet)
    frmskilleditor.Label17.Caption = ItemGive1val(currentsheet)
    frmskilleditor.Label20.Caption = ItemGive2val(currentsheet)
    frmskilleditor.Label25.Caption = ExpGiven(currentsheet)
    frmskilleditor.Label28.Caption = base_chance(currentsheet)
    frmskilleditor.Label10.Caption = "Sheet " & val(currentsheet + 1)

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

    iconn.Left = -val(HScroll1.Value * PIC_X)
    iconn.Top = -val(HScroll2.Value * PIC_Y)

End Sub

Private Sub HScroll1_Change()

    iconn.Left = -val(HScroll1.Value * PIC_X)
    iconn.Top = -val(HScroll2.Value * PIC_Y)
    Label9.Caption = HScroll1.Value

End Sub

Private Sub HScroll2_Change()

    iconn.Left = -val(HScroll1.Value * PIC_X)
    iconn.Top = -val(HScroll2.Value * PIC_Y)
    Label8.Caption = HScroll2.Value

End Sub

Private Sub HScroll3_Change()

    Label11.Caption = val(HScroll3.Value)

End Sub

Private Sub HScroll4_Change()

    Label14.Caption = val(HScroll4.Value)

End Sub

Private Sub HScroll5_Change()

    Label17.Caption = val(HScroll5.Value)

End Sub

Private Sub HScroll6_Change()

    Label20.Caption = val(HScroll6.Value)

End Sub

Private Sub HScroll7_Change()

    Label25.Caption = HScroll7.Value

End Sub

Private Sub HScroll8_Change()

    Label28.Caption = HScroll8.Value

End Sub

Sub savesheet()

    itemequiped(currentsheet) = frmskilleditor.cmbItem(0).ListIndex
    ItemTake1num(currentsheet) = frmskilleditor.cmbItem(1).ListIndex
    ItemTake2num(currentsheet) = frmskilleditor.cmbItem(2).ListIndex
    ItemGive1num(currentsheet) = frmskilleditor.cmbItem(3).ListIndex
    ItemGive2num(currentsheet) = frmskilleditor.cmbItem(4).ListIndex
    minlevel(currentsheet) = frmskilleditor.cmbLevel.ListIndex
    ItemTake1val(currentsheet) = frmskilleditor.HScroll3.Value
    ItemTake2val(currentsheet) = frmskilleditor.HScroll4.Value
    ItemGive1val(currentsheet) = frmskilleditor.HScroll5.Value
    ItemGive2val(currentsheet) = frmskilleditor.HScroll6.Value
    ExpGiven(currentsheet) = frmskilleditor.HScroll7.Value
    base_chance(currentsheet) = frmskilleditor.HScroll8.Value
    ItemTake1val(currentsheet) = frmskilleditor.Label11.Caption
    ItemTake2val(currentsheet) = frmskilleditor.Label14.Caption
    ItemGive1val(currentsheet) = frmskilleditor.Label17.Caption
    ItemGive2val(currentsheet) = frmskilleditor.Label20.Caption

End Sub

