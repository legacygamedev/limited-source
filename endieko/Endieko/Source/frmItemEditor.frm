VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmEditItem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Editor"
   ClientHeight    =   6465
   ClientLeft      =   795
   ClientTop       =   705
   ClientWidth     =   6255
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
   ScaleHeight     =   431
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   417
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6345
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   11192
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
      TabPicture(0)   =   "frmItemEditor.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label26"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraSpell"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraVitals"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "VScroll1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "picPic"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Picture1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "fraAttributes"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtDesc"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "fraBow"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdSelectSprite"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtName"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmbType"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "picSelect"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmdCancel"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmdOk"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "fraEquipment"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
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
         Height          =   3135
         Left            =   240
         TabIndex        =   30
         Top             =   2040
         Visible         =   0   'False
         Width           =   3135
         Begin RichTextLib.RichTextBox txtMagiReq 
            Height          =   255
            Left            =   1320
            TabIndex        =   69
            Top             =   1520
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            _Version        =   393217
            Enabled         =   -1  'True
            MaxLength       =   4
            Appearance      =   0
            TextRTF         =   $"frmItemEditor.frx":001C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin RichTextLib.RichTextBox txtSpeedReq 
            Height          =   255
            Left            =   1320
            TabIndex        =   59
            Top             =   1280
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            _Version        =   393217
            Enabled         =   -1  'True
            MaxLength       =   4
            Appearance      =   0
            TextRTF         =   $"frmItemEditor.frx":0096
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin RichTextLib.RichTextBox txtDefReq 
            Height          =   255
            Left            =   1320
            TabIndex        =   58
            Top             =   1040
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            _Version        =   393217
            Enabled         =   -1  'True
            MaxLength       =   4
            Appearance      =   0
            TextRTF         =   $"frmItemEditor.frx":0110
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin RichTextLib.RichTextBox txtDamage 
            Height          =   255
            Left            =   1320
            TabIndex        =   57
            Top             =   560
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            _Version        =   393217
            Enabled         =   -1  'True
            MaxLength       =   4
            Appearance      =   0
            TextRTF         =   $"frmItemEditor.frx":018A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin RichTextLib.RichTextBox txtStrReq 
            Height          =   255
            Left            =   1320
            TabIndex        =   56
            Top             =   800
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            _Version        =   393217
            Enabled         =   -1  'True
            MaxLength       =   4
            Appearance      =   0
            TextRTF         =   $"frmItemEditor.frx":0204
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.CheckBox chkInd 
            Caption         =   "Ind."
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
            Left            =   2040
            TabIndex        =   55
            Top             =   320
            Width           =   615
         End
         Begin VB.HScrollBar scrlClassReq 
            Height          =   255
            Left            =   240
            Max             =   1
            Min             =   -1
            TabIndex        =   32
            Top             =   2160
            Value           =   -1
            Width           =   2775
         End
         Begin VB.HScrollBar scrlAccessReq 
            Height          =   255
            Left            =   240
            Max             =   4
            TabIndex        =   31
            Top             =   2760
            Width           =   2775
         End
         Begin RichTextLib.RichTextBox txtDurability 
            Height          =   255
            Left            =   1320
            TabIndex        =   54
            Top             =   315
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            _Version        =   393217
            Enabled         =   -1  'True
            MaxLength       =   4
            Appearance      =   0
            TextRTF         =   $"frmItemEditor.frx":027E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Magi Req :"
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
            TabIndex        =   68
            Top             =   1560
            Width           =   975
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
            Height          =   255
            Left            =   240
            TabIndex        =   41
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label3 
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
            TabIndex        =   40
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label8 
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
            Height          =   255
            Left            =   240
            TabIndex        =   39
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label9 
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
            Height          =   255
            Left            =   240
            TabIndex        =   38
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label10 
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
            Height          =   255
            Left            =   240
            TabIndex        =   37
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label14 
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
            Left            =   240
            TabIndex        =   36
            Top             =   1920
            Width           =   975
         End
         Begin VB.Label Label15 
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
            Left            =   240
            TabIndex        =   35
            Top             =   2520
            Width           =   975
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
            TabIndex        =   34
            Top             =   1920
            Width           =   330
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
            TabIndex        =   33
            Top             =   2520
            Width           =   1695
         End
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
         Left            =   360
         TabIndex        =   53
         Top             =   5880
         Width           =   1455
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
         Left            =   1920
         TabIndex        =   52
         Top             =   5880
         Width           =   1455
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
         Left            =   2910
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   29
         Top             =   1470
         Width           =   480
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
         ItemData        =   "frmItemEditor.frx":02F8
         Left            =   240
         List            =   "frmItemEditor.frx":032F
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   1080
         Width           =   3135
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
         Left            =   240
         MaxLength       =   20
         TabIndex        =   27
         Top             =   720
         Width           =   3135
      End
      Begin VB.CommandButton cmdSelectSprite 
         Caption         =   "Select Sprite"
         Height          =   270
         Left            =   240
         TabIndex        =   24
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Frame fraBow 
         Caption         =   "Bow"
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
         Left            =   3480
         TabIndex        =   16
         Top             =   480
         Visible         =   0   'False
         Width           =   2535
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
            TabIndex        =   21
            Top             =   240
            Width           =   615
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
            TabIndex        =   18
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
               TabIndex        =   19
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
                  TabIndex        =   20
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
            ItemData        =   "frmItemEditor.frx":03E0
            Left            =   120
            List            =   "frmItemEditor.frx":03E2
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   600
            Width           =   2295
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
            Height          =   225
            Left            =   720
            TabIndex        =   23
            Top             =   1155
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
            TabIndex        =   22
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
         Left            =   240
         MaxLength       =   150
         TabIndex        =   14
         Top             =   5520
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
         Height          =   3135
         Left            =   3360
         TabIndex        =   5
         Top             =   2040
         Width           =   2655
         Begin VB.CheckBox chkDrop 
            Caption         =   "Does not drop on death"
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
            TabIndex        =   71
            Top             =   2760
            Width           =   2295
         End
         Begin VB.CheckBox chkFix 
            Caption         =   "Cannot be Repaired"
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
            TabIndex        =   70
            Top             =   2520
            Width           =   1935
         End
         Begin RichTextLib.RichTextBox txtAddExp 
            Height          =   255
            Left            =   1200
            TabIndex        =   67
            Top             =   2000
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            _Version        =   393217
            Enabled         =   -1  'True
            MaxLength       =   3
            Appearance      =   0
            TextRTF         =   $"frmItemEditor.frx":03E4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin RichTextLib.RichTextBox txtAddSpeed 
            Height          =   255
            Left            =   1200
            TabIndex        =   66
            Top             =   1760
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            _Version        =   393217
            Enabled         =   -1  'True
            Appearance      =   0
            TextRTF         =   $"frmItemEditor.frx":045E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin RichTextLib.RichTextBox txtAddMagi 
            Height          =   255
            Left            =   1200
            TabIndex        =   65
            Top             =   1520
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            _Version        =   393217
            Enabled         =   -1  'True
            Appearance      =   0
            TextRTF         =   $"frmItemEditor.frx":04D8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin RichTextLib.RichTextBox txtAddDef 
            Height          =   255
            Left            =   1200
            TabIndex        =   64
            Top             =   1280
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            _Version        =   393217
            Enabled         =   -1  'True
            Appearance      =   0
            TextRTF         =   $"frmItemEditor.frx":0552
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin RichTextLib.RichTextBox txtAddStr 
            Height          =   255
            Left            =   1200
            TabIndex        =   63
            Top             =   1040
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            _Version        =   393217
            Enabled         =   -1  'True
            Appearance      =   0
            TextRTF         =   $"frmItemEditor.frx":05CC
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin RichTextLib.RichTextBox txtAddSP 
            Height          =   255
            Left            =   1200
            TabIndex        =   62
            Top             =   800
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            _Version        =   393217
            Enabled         =   -1  'True
            Appearance      =   0
            TextRTF         =   $"frmItemEditor.frx":0646
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin RichTextLib.RichTextBox txtAddMP 
            Height          =   255
            Left            =   1200
            TabIndex        =   61
            Top             =   560
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            _Version        =   393217
            Enabled         =   -1  'True
            Appearance      =   0
            TextRTF         =   $"frmItemEditor.frx":06C0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin RichTextLib.RichTextBox txtAddHP 
            Height          =   255
            Left            =   1200
            TabIndex        =   60
            Top             =   320
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            _Version        =   393217
            Enabled         =   -1  'True
            Appearance      =   0
            TextRTF         =   $"frmItemEditor.frx":073A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "Add EXP% :"
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
            TabIndex        =   13
            Top             =   2040
            Width           =   855
         End
         Begin VB.Label Label24 
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
            Left            =   240
            TabIndex        =   12
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label23 
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
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label Label22 
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
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label Label21 
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
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label20 
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
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label19 
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
            TabIndex        =   7
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label18 
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
            Left            =   240
            TabIndex        =   6
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   2880
         ScaleHeight     =   510
         ScaleWidth      =   510
         TabIndex        =   4
         Top             =   1440
         Width           =   540
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
         Left            =   240
         ScaleHeight     =   160
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   192
         TabIndex        =   25
         Top             =   2160
         Visible         =   0   'False
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
            TabIndex        =   26
            Top             =   0
            Width           =   2880
         End
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   2400
         Left            =   3120
         Max             =   464
         TabIndex        =   3
         Top             =   2160
         Visible         =   0   'False
         Width           =   255
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
         Height          =   2895
         Left            =   240
         TabIndex        =   42
         Top             =   2040
         Visible         =   0   'False
         Width           =   3135
         Begin VB.HScrollBar scrlVitalMod 
            Height          =   255
            Left            =   240
            Max             =   255
            TabIndex        =   43
            Top             =   1080
            Value           =   1
            Width           =   2655
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
            TabIndex        =   45
            Top             =   840
            Width           =   735
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
            TabIndex        =   44
            Top             =   840
            Width           =   495
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
         Height          =   2895
         Left            =   240
         TabIndex        =   46
         Top             =   2040
         Visible         =   0   'False
         Width           =   3135
         Begin VB.HScrollBar scrlSpell 
            Height          =   255
            Left            =   240
            Max             =   255
            Min             =   1
            TabIndex        =   47
            Top             =   1200
            Value           =   1
            Width           =   2775
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
            TabIndex        =   51
            Top             =   960
            Width           =   495
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
            TabIndex        =   50
            Top             =   960
            Width           =   1095
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
            TabIndex        =   49
            Top             =   360
            Width           =   1095
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
            TabIndex        =   48
            Top             =   600
            Width           =   2760
         End
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
         Left            =   240
         TabIndex        =   15
         Top             =   5280
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
         TabIndex        =   2
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
         TabIndex        =   1
         Top             =   2040
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmEditItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SpriteSelect As Boolean

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

Private Sub chkInd_Click()
    If chkInd.Value = 0 Then
        txtDurability.Enabled = True
    Else
        txtDurability.Enabled = False
        txtDurability.Text = -1
    End If
End Sub

Private Sub cmbBow_Change()
    lblName.Caption = Arrows(cmbBow.ListIndex + 1).Name
    picBow.Top = (Arrows(cmbBow.ListIndex + 1).Pic * 32) * -1
End Sub

Private Sub cmdOk_Click()
If txtDurability.Text = vbNullString Then txtDurability.Text = 0
If txtDamage.Text = vbNullString Then txtDamage.Text = 0
If txtStrReq.Text = vbNullString Then txtStrReq.Text = 0
If txtDefReq.Text = vbNullString Then txtDefReq.Text = 0
If txtSpeedReq.Text = vbNullString Then txtSpeedReq.Text = 0
If txtMagiReq.Text = vbNullString Then txtMagiReq.Text = 0
If txtAddHP.Text = vbNullString Then txtAddHP.Text = 0
If txtAddMP.Text = vbNullString Then txtAddMP.Text = 0
If txtAddSP.Text = vbNullString Then txtAddSP.Text = 0
If txtAddStr.Text = vbNullString Then txtAddStr.Text = 0
If txtAddDef.Text = vbNullString Then txtAddDef.Text = 0
If txtAddMagi.Text = vbNullString Then txtAddMagi.Text = 0
If txtAddSpeed.Text = vbNullString Then txtAddSpeed.Text = 0
If txtAddExp.Text = vbNullString Then txtAddExp.Text = 0


    Call ItemEditorOk
End Sub

Private Sub cmdCancel_Click()
    Call ItemEditorCancel
End Sub

Private Sub cmbType_Click()
    If (cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (cmbType.ListIndex <= ITEM_TYPE_BOOTS) Then
        If cmbType.ListIndex = ITEM_TYPE_WEAPON Then
            Label3.Caption = "Damage :"
        Else
            Label3.Caption = "Defence :"
        End If
        fraEquipment.Visible = True
        fraAttributes.Visible = True
        frmEditItem.SSTab1.Width = 403
        frmEditItem.Width = 6345
    Else
        fraEquipment.Visible = False
        fraAttributes.Visible = False
    End If
    
    If (cmbType.ListIndex >= ITEM_TYPE_POTIONADDHP) And (cmbType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Then
        fraVitals.Visible = True
        fraAttributes.Visible = False
        fraAttributes.Visible = False
        SSTab1.Width = 235
        frmEditItem.Width = 3825
    Else
        fraVitals.Visible = False
    End If
    
    If (cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        fraSpell.Visible = True
        fraAttributes.Visible = False
        fraAttributes.Visible = False
        SSTab1.Width = 235
        frmEditItem.Width = 3825
    Else
        fraSpell.Visible = False
    End If
    
    If (cmbType.ListIndex = ITEM_TYPE_CURRENCY) Then
        fraAttributes.Visible = False
        SSTab1.Width = 235
        frmEditItem.Width = 3825
    End If
End Sub

Private Sub cmdSelectSprite_Click()
    If Not SpriteSelect Then
        SpriteSelect = True
        frmEditItem.Width = 3825
        Label5.Visible = True
        Picpic.Visible = True
        VScroll1.Visible = True
        fraAttributes.Visible = False
        fraBow.Visible = False
        fraEquipment.Visible = False
        fraSpell.Visible = False
        fraVitals.Visible = False
        SSTab1.Width = 235
    Else
        SpriteSelect = False
        If (cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (cmbType.ListIndex <= ITEM_TYPE_BOOTS) Then
            If cmbType.ListIndex = ITEM_TYPE_WEAPON Then
                Label3.Caption = "Damage :"
            Else
                Label3.Caption = "Defence :"
            End If
            fraEquipment.Visible = True
            fraAttributes.Visible = True
            SSTab1.Width = 403
            frmEditItem.Width = 6345
        Else
            fraEquipment.Visible = False
            fraAttributes.Visible = False
        End If
    
        If (cmbType.ListIndex >= ITEM_TYPE_POTIONADDHP) And (cmbType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Then
            fraVitals.Visible = True
            fraAttributes.Visible = False
            SSTab1.Width = 235
            frmEditItem.Width = 3825
        Else
            fraVitals.Visible = False
        End If
    
        If (cmbType.ListIndex = ITEM_TYPE_SPELL) Then
            fraSpell.Visible = True
            fraAttributes.Visible = False
            SSTab1.Width = 235
            frmEditItem.Width = 3825
        Else
            fraSpell.Visible = False
        End If
        
        If (cmbType.ListIndex = ITEM_TYPE_CURRENCY) Then
            fraAttributes.Visible = False
            SSTab1.Width = 235
            frmEditItem.Width = 3825
        End If

        Label5.Visible = False
        Picpic.Visible = False
        VScroll1.Visible = False
    End If
End Sub

Private Sub Form_Load()
    picBow.Picture = LoadPicture(App.Path & "\graphics\arrows.bmp")
    picItems.Height = 320 * PIC_Y
    Call BitBlt(picSelect.hdc, 0, 0, PIC_X, PIC_Y, picItems.hdc, EditorItemX * PIC_X, EditorItemY * PIC_Y, SRCCOPY)
End Sub

Private Sub picItems_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        EditorItemX = Int(X / PIC_X)
        EditorItemY = Int(Y / PIC_Y)
    End If
    Call BitBlt(picSelect.hdc, 0, 0, PIC_X, PIC_Y, picItems.hdc, EditorItemX * PIC_X, EditorItemY * PIC_Y, SRCCOPY)
End Sub

Private Sub picItems_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        EditorItemX = Int(X / PIC_X)
        EditorItemY = Int(Y / PIC_Y)
    End If
    Call BitBlt(picSelect.hdc, 0, 0, PIC_X, PIC_Y, picItems.hdc, EditorItemX * PIC_X, EditorItemY * PIC_Y, SRCCOPY)
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

Private Sub scrlClassReq_Change()
If scrlClassReq.Value = -1 Then
    Label16.Caption = "None"
Else
    Label16.Caption = scrlClassReq.Value & " - " & Trim$(Class(scrlClassReq.Value).Name)
End If
End Sub

Private Sub scrlSpell_Change()
    lblSpellName.Caption = Trim$(Spell(scrlSpell.Value).Name)
    lblSpell.Caption = STR$(scrlSpell.Value)
End Sub

Private Sub Timer1_Timer()
    Call BitBlt(picSelect.hdc, 0, 0, PIC_X, PIC_Y, picItems.hdc, EditorItemX * PIC_X, EditorItemY * PIC_Y, SRCCOPY)
End Sub

Private Sub txtAddDef_Change()
    If IsNumeric(txtAddDef.Text) Then
        If txtAddDef.Text > 1000 Then
            txtAddDef.Text = 1000
        End If
    Else
        txtAddDef.Text = 0
    End If
End Sub

Private Sub txtAddExp_Change()
    If IsNumeric(txtAddExp.Text) Then
        If txtAddExp.Text > 100 Then
            txtAddExp.Text = 100
        End If
    Else
        txtAddExp.Text = 0
    End If
End Sub

Private Sub txtAddHP_Change()
    If IsNumeric(txtAddHP.Text) Then
        If txtAddHP.Text > 1000 Then
            txtAddHP.Text = 1000
        End If
    Else
        txtAddHP.Text = 0
    End If
End Sub

Private Sub txtAddMagi_Change()
    If IsNumeric(txtAddMagi.Text) Then
        If txtAddMagi.Text > 1000 Then
            txtAddMagi.Text = 1000
        End If
    Else
        txtAddMagi.Text = 0
    End If
End Sub

Private Sub txtAddMP_Change()
    If IsNumeric(txtAddMP.Text) Then
        If txtAddMP.Text > 1000 Then
            txtAddMP.Text = 1000
        End If
    Else
        txtAddMP.Text = 0
    End If
End Sub

Private Sub txtAddSP_Change()
    If IsNumeric(txtAddSP.Text) Then
        If txtAddSP.Text > 1000 Then
            txtAddSP.Text = 1000
        End If
    Else
        txtAddSP.Text = 0
    End If
End Sub

Private Sub txtAddSpeed_Change()
    If IsNumeric(txtAddSpeed.Text) Then
        If txtAddSpeed.Text > 1000 Then
            txtAddSpeed.Text = 1000
        End If
    Else
        txtAddSpeed.Text = 0
    End If
End Sub

Private Sub txtAddStr_Change()
    If IsNumeric(txtAddStr.Text) Then
        If txtAddStr.Text > 1000 Then
            txtAddStr.Text = 1000
        End If
    Else
        txtAddStr.Text = 0
    End If
End Sub

Private Sub txtDamage_Change()
    If IsNumeric(txtDamage.Text) Then
        If txtDamage.Text > 1000 Then
            txtDamage.Text = 1000
        End If
    Else
        txtDamage.Text = 0
    End If
End Sub

Private Sub txtDefReq_Change()
    If IsNumeric(txtDefReq.Text) Then
        If txtDefReq.Text > 1000 Then
            txtDefReq.Text = 1000
        End If
    Else
        txtDefReq.Text = 0
    End If
End Sub

Private Sub txtDurability_Change()
    If IsNumeric(txtDurability.Text) Then
        If txtDurability.Text > 1000 Then
            txtDurability.Text = 1000
        End If
    Else
        txtDurability.Text = 0
    End If
End Sub

Private Sub txtMagiReq_Change()
    If IsNumeric(txtMagiReq.Text) Then
        If txtMagiReq.Text > 1000 Then
            txtMagiReq.Text = 1000
        End If
    Else
        txtMagiReq.Text = 0
    End If
End Sub

Private Sub txtSpeedReq_Change()
    If IsNumeric(txtSpeedReq.Text) Then
        If txtSpeedReq.Text > 1000 Then
            txtSpeedReq.Text = 1000
        End If
    Else
        txtSpeedReq.Text = 0
    End If
End Sub

Private Sub txtStrReq_Change()
    If IsNumeric(txtStrReq.Text) Then
        If txtStrReq.Text > 1000 Then
            txtStrReq.Text = 1000
        End If
    Else
        txtStrReq.Text = 0
    End If
End Sub

Private Sub VScroll1_Change()
    picItems.Top = (VScroll1.Value * PIC_Y) * -1
End Sub
