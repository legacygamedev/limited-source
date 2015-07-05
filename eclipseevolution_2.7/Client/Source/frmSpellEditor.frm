VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmSpellEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spell Editor"
   ClientHeight    =   5655
   ClientLeft      =   195
   ClientTop       =   525
   ClientWidth     =   6375
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
   Icon            =   "frmSpellEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   377
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   300
      Left            =   4080
      TabIndex        =   46
      Top             =   5280
      Width           =   2190
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Save Spell"
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
      Left            =   120
      TabIndex        =   45
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   120
      Top             =   120
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   8916
      _Version        =   393216
      TabHeight       =   397
      TabMaxWidth     =   3545
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Information"
      TabPicture(0)   =   "frmSpellEditor.frx":0FC2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblSound"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblVitalMod"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblRange"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblElement"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label9"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "info"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "scrlSound"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "scrlVitalMod"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "scrlRange"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "scrlElement"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "chkArea"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Animation"
      TabPicture(1)   =   "frmSpellEditor.frx":0FDE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "scrlSpellAnim"
      Tab(1).Control(1)=   "scrlSpellTime"
      Tab(1).Control(2)=   "scrlSpellDone"
      Tab(1).Control(3)=   "Command1"
      Tab(1).Control(4)=   "picSpell"
      Tab(1).Control(5)=   "chkBig"
      Tab(1).Control(6)=   "Frame3"
      Tab(1).Control(7)=   "Picture1"
      Tab(1).Control(8)=   "lblSpellAnim"
      Tab(1).Control(9)=   "lblSpellTime"
      Tab(1).Control(10)=   "lblSpellDone"
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "Requirements"
      TabPicture(2)   =   "frmSpellEditor.frx":0FFA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).ControlCount=   1
      Begin VB.HScrollBar scrlSpellAnim 
         Height          =   270
         Left            =   -74760
         Max             =   2000
         TabIndex        =   41
         Top             =   720
         Value           =   1
         Width           =   3495
      End
      Begin VB.HScrollBar scrlSpellTime 
         Height          =   270
         Left            =   -74760
         Max             =   500
         Min             =   40
         TabIndex        =   40
         Top             =   3120
         Value           =   40
         Width           =   5655
      End
      Begin VB.HScrollBar scrlSpellDone 
         Height          =   270
         Left            =   -74760
         Max             =   10
         Min             =   1
         TabIndex        =   39
         Top             =   3840
         Value           =   1
         Width           =   5655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Refresh"
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
         Left            =   -70680
         TabIndex        =   38
         Top             =   2400
         Width           =   1605
      End
      Begin VB.PictureBox picSpell 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   -70120
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   36
         Top             =   1275
         Width           =   480
      End
      Begin VB.CheckBox chkBig 
         Caption         =   "Big Spell"
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
         Left            =   -74760
         TabIndex        =   35
         Top             =   1080
         Width           =   1215
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
         Height          =   1215
         Left            =   -74760
         TabIndex        =   28
         Top             =   1440
         Width           =   3735
         Begin VB.HScrollBar HScroll1 
            Height          =   255
            Left            =   840
            Max             =   100
            TabIndex        =   32
            Top             =   600
            Width           =   2775
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
            TabIndex        =   29
            Top             =   360
            Width           =   540
            Begin VB.PictureBox picEmoticon 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   480
               Left            =   15
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   30
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
                  TabIndex        =   31
                  Top             =   0
                  Width           =   480
               End
            End
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Spell ID:"
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
            TabIndex        =   34
            Top             =   360
            Width           =   555
         End
         Begin VB.Label Label13 
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
            Left            =   1440
            TabIndex        =   33
            Top             =   360
            Width           =   315
         End
      End
      Begin VB.CheckBox chkArea 
         Caption         =   "Area Effect"
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
         Left            =   1320
         TabIndex        =   27
         Top             =   3550
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         Caption         =   "Spell Requirements"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   -74760
         TabIndex        =   16
         Top             =   600
         Width           =   5175
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
            ItemData        =   "frmSpellEditor.frx":1016
            Left            =   120
            List            =   "frmSpellEditor.frx":1018
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   480
            Width           =   4905
         End
         Begin VB.HScrollBar scrlCost 
            Height          =   270
            Left            =   120
            Max             =   1000
            TabIndex        =   18
            Top             =   1680
            Width           =   4935
         End
         Begin VB.HScrollBar scrlLevelReq 
            Height          =   270
            Left            =   120
            Max             =   500
            TabIndex        =   17
            Top             =   1080
            Value           =   1
            Width           =   4935
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Class Required"
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
            Width           =   915
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
            Left            =   840
            TabIndex        =   22
            Top             =   1440
            Width           =   75
         End
         Begin VB.Label lblLevelReq 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "God's Only Spell"
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
            TabIndex        =   21
            Top             =   840
            Width           =   1050
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MP Cost:"
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
            TabIndex        =   20
            Top             =   1440
            Width           =   585
         End
         Begin VB.Label Label3 
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
            Left            =   120
            TabIndex        =   19
            Top             =   840
            Width           =   960
         End
      End
      Begin VB.HScrollBar scrlElement 
         Height          =   255
         Left            =   240
         Max             =   1000
         TabIndex        =   13
         Top             =   4560
         Value           =   1
         Width           =   5655
      End
      Begin VB.HScrollBar scrlRange 
         Height          =   270
         Left            =   240
         Max             =   30
         Min             =   1
         TabIndex        =   12
         Top             =   3840
         Value           =   1
         Width           =   5655
      End
      Begin VB.HScrollBar scrlVitalMod 
         Height          =   270
         Left            =   240
         Max             =   1000
         TabIndex        =   5
         Top             =   2400
         Width           =   5655
      End
      Begin VB.HScrollBar scrlSound 
         Height          =   270
         Left            =   240
         Max             =   100
         TabIndex        =   4
         Top             =   3120
         Width           =   5655
      End
      Begin VB.Frame info 
         Caption         =   "Information"
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
         Top             =   510
         Width           =   5655
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
            ItemData        =   "frmSpellEditor.frx":101A
            Left            =   120
            List            =   "frmSpellEditor.frx":1033
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   1080
            Width           =   5355
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
            Width           =   5355
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Spell Type"
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
            TabIndex        =   25
            Top             =   840
            Width           =   840
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Spell Name"
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
            Top             =   240
            Width           =   960
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1605
         Left            =   -70680
         ScaleHeight     =   105
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   105
         TabIndex        =   37
         Top             =   720
         Width           =   1605
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
         Left            =   -74760
         TabIndex        =   44
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblSpellTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time: 40"
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
         Left            =   -74760
         TabIndex        =   43
         Top             =   2880
         Width           =   555
      End
      Begin VB.Label lblSpellDone 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cycle Animation 1 Time"
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
         Left            =   -74760
         TabIndex        =   42
         Top             =   3600
         Width           =   1515
      End
      Begin VB.Label Label9 
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
         Left            =   255
         TabIndex        =   15
         Top             =   4320
         Width           =   555
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
         Left            =   840
         TabIndex        =   14
         Top             =   4320
         Width           =   1410
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
         TabIndex        =   11
         Top             =   3600
         Width           =   75
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Range:"
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
         Top             =   3600
         Width           =   780
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
         Left            =   960
         TabIndex        =   9
         Top             =   2160
         Width           =   75
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Vital Mod:"
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
         Top             =   2160
         Width           =   660
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Sound:"
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
         Top             =   2880
         Width           =   540
      End
      Begin VB.Label lblSound 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Sound"
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
         TabIndex        =   6
         Top             =   2880
         Width           =   615
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

Private Sub chkBig_Click()
    frmSpellEditor.ScaleMode = 3
    Done = 0
    SpellVar = 0
    picSpell.Refresh
    If chkBig.Value = Checked Then
        picSpell.Width = 1440
        picSpell.Height = 1440
        picSpell.Top = 800
        picSpell.Left = 4400
    Else
        picSpell.Width = 480
        picSpell.Height = 480
        picSpell.Top = 1275
        picSpell.Left = 4880
    End If
End Sub

Private Sub Command1_Click()
    Done = 0
End Sub

Private Sub Form_Load()
    scrlElement.Max = MAX_ELEMENTS
End Sub

Private Sub HScroll1_Change()
    Label13.Caption = STR(HScroll1.Value)
    frmSpellEditor.iconn.Top = (HScroll1.Value * 32) * -1
End Sub

Private Sub scrlCost_Change()
    lblCost.Caption = STR(scrlCost.Value)
End Sub

Private Sub scrlElement_Change()
    lblElement.Caption = Element(scrlElement.Value).Name
End Sub

Private Sub scrlLevelReq_Change()
    If STR(scrlLevelReq.Value) = 0 Then
        lblLevelReq.Caption = "God's Only Spell"
    Else
        lblLevelReq.Caption = STR(scrlLevelReq.Value)
    End If
End Sub

Private Sub scrlRange_Change()
    lblRange.Caption = STR(scrlRange.Value)
End Sub

Private Sub scrlSound_Change()
    If STR(scrlSound.Value) = 0 Then
        lblSound.Caption = "No Sound"
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
    String2 = "Times"
    If scrlSpellDone.Value = 1 Then
        String2 = "Time"
    End If
    lblSpellDone.Caption = "Cycle Animation " & scrlSpellDone.Value & " " & String2
    Done = 0
End Sub

Private Sub scrlSpellTime_Change()
    lblSpellTime.Caption = "Time: " & scrlSpellTime.Value
    Done = 0
End Sub

Private Sub scrlVitalMod_Change()
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

    If chkBig.Value = Checked Then
        SpellAnim = scrlSpellAnim.Value * 3
    End If

    If SpellAnim <= 0 Then
        Exit Sub
    End If
    If Done = SpellDone Then
        Exit Sub
    End If
    If chkBig = Checked Then
        With dRECT
            .Top = 0
            .Bottom = PIC_Y + 64
            .Left = 0
            .Right = PIC_X + 64
        End With
    Else
        With dRECT
            .Top = 0
            .Bottom = PIC_Y
            .Left = 0
            .Right = PIC_X
        End With
    End If
    If chkBig.Value = Checked Then
        If SpellVar > 32 Then
            Done = Done + 1
            SpellVar = 0
        End If
        If GetTickCount > Time + SpellTime Then
            Time = GetTickCount
            SpellVar = SpellVar + 3
        End If
    Else
        If SpellVar > 10 Then
            Done = Done + 1
            SpellVar = 0
        End If
        If GetTickCount > Time + SpellTime Then
            Time = GetTickCount
            SpellVar = SpellVar + 1
        End If
    End If
    If chkBig = Checked Then
        If DD_BigSpellAnim Is Nothing Then
        Else
            With sRECT
                .Top = SpellAnim * PIC_Y
                .Bottom = .Top + (PIC_Y * 3)
                .Left = SpellVar * PIC_X
                .Right = .Left + (PIC_X * 3)
            End With

            Call DD_BigSpellAnim.BltToDC(picSpell.hDC, sRECT, dRECT)
            picSpell.Refresh
        End If
    Else
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
    End If
End Sub
Private Sub cmbType_Click()
    If (cmbType.ListIndex = SPELL_TYPE_SCRIPTED) Then
        Label4.Caption = "Script"
        Label8.Visible = False
        lblSound.Visible = False
        scrlSound.Visible = False
        Label2.Visible = False
        lblRange.Visible = False
        scrlRange.Visible = False
        lblSpellAnim.Visible = False
        scrlSpellAnim.Visible = False
        lblSpellTime.Visible = False
        scrlSpellTime.Visible = False
        lblSpellDone.Visible = False
        scrlSpellDone.Visible = False
        chkArea.Visible = False
        Command1.Visible = False
        picSpell.Visible = False

    Else
        Label4.Caption = "Vital Mod"
        Label8.Visible = True
        lblSound.Visible = True
        scrlSound.Visible = True
        Label2.Visible = True
        lblRange.Visible = True
        scrlRange.Visible = True
        lblSpellAnim.Visible = True
        scrlSpellAnim.Visible = True
        lblSpellTime.Visible = True
        scrlSpellTime.Visible = True
        lblSpellDone.Visible = True
        scrlSpellDone.Visible = True
        chkArea.Visible = True
        Command1.Visible = True
        picSpell.Visible = True
    End If
End Sub
