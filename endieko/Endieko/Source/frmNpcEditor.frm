VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmEditNpc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Npc Editor"
   ClientHeight    =   6975
   ClientLeft      =   165
   ClientTop       =   390
   ClientWidth     =   5175
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
   ScaleHeight     =   465
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar scrlDropItem 
      Height          =   255
      Left            =   1200
      Max             =   5
      Min             =   1
      TabIndex        =   35
      Top             =   4560
      Value           =   1
      Width           =   3255
   End
   Begin VB.HScrollBar scrlNum 
      Height          =   255
      Left            =   1200
      Max             =   500
      TabIndex        =   30
      Top             =   5280
      Value           =   1
      Width           =   3255
   End
   Begin VB.HScrollBar scrlValue 
      Height          =   255
      Left            =   1200
      Max             =   10000
      TabIndex        =   27
      Top             =   5640
      Value           =   1
      Width           =   3255
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
      Left            =   2880
      TabIndex        =   25
      Text            =   "0"
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   15
      Top             =   3480
      Width           =   4695
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
         TabIndex        =   17
         Top             =   240
         Width           =   3615
      End
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
         TabIndex        =   16
         Text            =   "0"
         Top             =   600
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
         TabIndex        =   19
         Top             =   240
         Width           =   735
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
         Left            =   720
         TabIndex        =   18
         Top             =   600
         Width           =   1815
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
      Height          =   2535
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   4695
      Begin RichTextLib.RichTextBox txtKarma 
         Height          =   255
         Left            =   3720
         TabIndex        =   46
         Top             =   2000
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         _Version        =   393217
         Enabled         =   -1  'True
         Appearance      =   0
         TextRTF         =   $"frmNpcEditor.frx":005C
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
      Begin VB.VScrollBar scrlSprite 
         Height          =   1575
         Left            =   240
         Max             =   500
         TabIndex        =   45
         Top             =   720
         Width           =   255
      End
      Begin RichTextLib.RichTextBox txtExp 
         Height          =   255
         Left            =   3720
         TabIndex        =   44
         Top             =   1755
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         _Version        =   393217
         Enabled         =   -1  'True
         Appearance      =   0
         TextRTF         =   $"frmNpcEditor.frx":00D6
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
      Begin RichTextLib.RichTextBox txtStartingHP 
         Height          =   255
         Left            =   3720
         TabIndex        =   43
         Top             =   1515
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         _Version        =   393217
         Enabled         =   -1  'True
         Appearance      =   0
         TextRTF         =   $"frmNpcEditor.frx":0150
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
      Begin RichTextLib.RichTextBox txtMagi 
         Height          =   255
         Left            =   3720
         TabIndex        =   42
         Top             =   1275
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         _Version        =   393217
         Enabled         =   -1  'True
         Appearance      =   0
         TextRTF         =   $"frmNpcEditor.frx":01CA
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
      Begin RichTextLib.RichTextBox txtSPEED 
         Height          =   255
         Left            =   3720
         TabIndex        =   41
         Top             =   1035
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         _Version        =   393217
         Enabled         =   -1  'True
         Appearance      =   0
         TextRTF         =   $"frmNpcEditor.frx":0244
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
      Begin RichTextLib.RichTextBox txtDef 
         Height          =   255
         Left            =   3720
         TabIndex        =   40
         Top             =   795
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         _Version        =   393217
         Enabled         =   -1  'True
         Appearance      =   0
         TextRTF         =   $"frmNpcEditor.frx":02BE
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
      Begin RichTextLib.RichTextBox txtStrength 
         Height          =   255
         Left            =   3720
         TabIndex        =   39
         Top             =   555
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         _Version        =   393217
         Enabled         =   -1  'True
         Appearance      =   0
         TextRTF         =   $"frmNpcEditor.frx":0338
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
      Begin RichTextLib.RichTextBox txtRange 
         Height          =   255
         Left            =   3720
         TabIndex        =   38
         Top             =   315
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         _Version        =   393217
         Enabled         =   -1  'True
         Appearance      =   0
         TextRTF         =   $"frmNpcEditor.frx":03B2
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
         Left            =   1320
         TabIndex        =   22
         Top             =   360
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
         Left            =   4080
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   20
         Top             =   120
         Visible         =   0   'False
         Width           =   480
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
         TabIndex        =   7
         Top             =   1040
         Width           =   480
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
         TabIndex        =   21
         Top             =   720
         Width           =   1125
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Karma :"
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
         TabIndex        =   47
         Top             =   2040
         Width           =   855
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
         Left            =   2760
         TabIndex        =   24
         Top             =   1800
         Width           =   855
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
         Left            =   2760
         TabIndex        =   23
         Top             =   1560
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
         Left            =   840
         TabIndex        =   14
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
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Range :"
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
         Left            =   3000
         TabIndex        =   12
         Top             =   360
         Width           =   615
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
         Left            =   3000
         TabIndex        =   11
         Top             =   600
         Width           =   615
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
         Left            =   3000
         TabIndex        =   10
         Top             =   840
         Width           =   615
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
         Left            =   3000
         TabIndex        =   9
         Top             =   1080
         Width           =   615
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
         Left            =   3000
         TabIndex        =   8
         Top             =   1320
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
      Left            =   4560
      Top             =   0
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
      Left            =   2760
      TabIndex        =   3
      Top             =   6600
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
      Left            =   360
      TabIndex        =   2
      Top             =   6600
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
      Left            =   360
      TabIndex        =   37
      Top             =   4560
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
      Left            =   4560
      TabIndex        =   36
      Top             =   4560
      Width           =   75
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
      Left            =   360
      TabIndex        =   34
      Top             =   4920
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
      Left            =   1200
      TabIndex        =   33
      Top             =   4920
      Width           =   3495
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
      Left            =   360
      TabIndex        =   32
      Top             =   5280
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
      Left            =   4560
      TabIndex        =   31
      Top             =   5280
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
      Left            =   360
      TabIndex        =   29
      Top             =   5640
      Width           =   735
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
      Left            =   4560
      TabIndex        =   28
      Top             =   5640
      Width           =   75
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
      Left            =   960
      TabIndex        =   26
      Top             =   6000
      Width           =   1815
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
Attribute VB_Name = "frmEditNpc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BigNpc_Click()
frmEditNpc.ScaleMode = 3
    If BigNpc.Value = Checked Then
        frmEditNpc.picSprites.Picture = LoadPicture(App.Path & "\Graphics\bigsprites.bmp")
        picSprite.Width = 960
        picSprite.Height = 960
        picSprite.Top = 800
        picSprite.Left = 1170
    Else
        frmEditNpc.picSprites.Picture = LoadPicture(App.Path & "\Graphics\sprites.bmp")
        picSprite.Width = 480
        picSprite.Height = 480
        picSprite.Top = 1040
        picSprite.Left = 1400
    End If
End Sub

Private Sub Form_Load()
    scrlDropItem.Max = MAX_NPC_DROPS
    picSprites.Picture = LoadPicture(App.Path & "\Graphics\sprites.bmp")
End Sub

Private Sub scrlDropItem_Change()
    txtChance.Text = Npc(EditorIndex).ItemNPC(scrlDropItem.Value).Chance
    scrlNum.Value = Npc(EditorIndex).ItemNPC(scrlDropItem.Value).ItemNum
    scrlValue.Value = Npc(EditorIndex).ItemNPC(scrlDropItem.Value).ItemValue
    lblDropItem.Caption = scrlDropItem.Value
End Sub

Private Sub scrlSprite_Change()
    lblSprite.Caption = STR$(scrlSprite.Value)
End Sub

Private Sub scrlNum_Change()
    lblNum.Caption = STR$(scrlNum.Value)
    lblItemName.Caption = ""
    If scrlNum.Value > 0 Then
        lblItemName.Caption = Trim$(Item(scrlNum.Value).Name)
    End If
    Npc(EditorIndex).ItemNPC(scrlDropItem.Value).ItemNum = scrlNum.Value
End Sub

Private Sub scrlValue_Change()
    Npc(EditorIndex).ItemNPC(scrlDropItem.Value).ItemValue = scrlValue.Value
    lblValue.Caption = STR$(scrlValue.Value)
End Sub

Private Sub cmdOk_Click()

If txtDef.Text = vbNullString Then txtDef.Text = 0
If txtExp.Text = vbNullString Then txtDef.Text = 0
If txtMagi.Text = vbNullString Then txtMagi.Text = 0
If txtRange.Text = vbNullString Then txtRange.Text = 0
If txtSPEED.Text = vbNullString Then txtSPEED.Text = 0
If txtStartingHP.Text = vbNullString Then txtStartingHP.Text = 0
If txtStrength.Text = vbNullString Then txtStrength.Text = 0


    Call NpcEditorOk
End Sub

Private Sub cmdCancel_Click()
    Call NpcEditorCancel
End Sub

Private Sub tmrSprite_Timer()
    Call NpcEditorBltSprite
End Sub

Private Sub txtChance_Change()
    Npc(EditorIndex).ItemNPC(scrlDropItem.Value).Chance = Val(txtChance.Text)
End Sub

Private Sub txtDef_Change()
    If IsNumeric(txtDef.Text) Then
        If txtDef.Text > MAX_MAPX Then
            txtDef.Text = MAX_MAPX
        ElseIf txtDef.Text < 0 Then
            txtDef.Text = 0
        End If
    Else
        txtDef.Text = 0
    End If
End Sub

Private Sub txtExp_Change()
    If IsNumeric(txtExp.Text) Then
        If txtExp.Text > MAX_MAPX Then
            txtExp.Text = MAX_MAPX
        ElseIf txtExp.Text < 0 Then
            txtExp.Text = 0
        End If
    Else
        txtExp.Text = 0
    End If
End Sub

Private Sub txtKarma_Change()
    If IsNumeric(txtKarma.Text) Then
        If txtKarma.Text > 1000 Then
            txtKarma.Text = 1000
        ElseIf txtKarma.Text < -1000 Then
            txtKarma.Text = -1000
        End If
    Else
        txtKarma.Text = 0
    End If
End Sub

Private Sub txtMagi_Change()
    If IsNumeric(txtMagi.Text) Then
        If txtMagi.Text > MAX_MAPX Then
            txtMagi.Text = MAX_MAPX
        ElseIf txtMagi.Text < 0 Then
            txtMagi.Text = 0
        End If
    Else
        txtMagi.Text = 0
    End If
End Sub

Private Sub txtRange_Change()
    If IsNumeric(txtRange.Text) Then
        If txtRange.Text > MAX_MAPX Then
            txtRange.Text = MAX_MAPX
        ElseIf txtRange.Text < 0 Then
            txtRange.Text = 0
        End If
    Else
        txtRange.Text = 0
    End If
End Sub

Private Sub txtSPEED_Change()
    If IsNumeric(txtSPEED.Text) Then
        If txtSPEED.Text > MAX_MAPX Then
            txtSPEED.Text = MAX_MAPX
        ElseIf txtSPEED.Text < 0 Then
            txtSPEED.Text = 0
        End If
    Else
        txtSPEED.Text = 0
    End If
End Sub

Private Sub txtStartingHP_Change()
    If IsNumeric(txtStartingHP.Text) Then
        If txtStartingHP.Text > MAX_MAPX Then
            txtStartingHP.Text = MAX_MAPX
        ElseIf txtStartingHP.Text < 0 Then
            txtStartingHP.Text = 0
        End If
    Else
        txtStartingHP.Text = 0
    End If
End Sub

Private Sub txtStrength_Change()
    If IsNumeric(txtStrength.Text) Then
        If txtStrength.Text > MAX_MAPX Then
            txtStrength.Text = MAX_MAPX
        ElseIf txtStrength.Text < 0 Then
            txtStrength.Text = 0
        End If
    Else
        txtStrength.Text = 0
    End If
End Sub
