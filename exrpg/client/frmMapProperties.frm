VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form frmMapProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Properties"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
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
   Icon            =   "frmMapProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   5790
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   11245
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
      TabCaption(0)   =   "Variables"
      TabPicture(0)   =   "frmMapProperties.frx":2372
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdCancel"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdOk"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtName"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame3"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame4"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.Frame Frame4 
         Caption         =   "None Player Characters (NPCs)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   120
         TabIndex        =   23
         Top             =   3480
         Width           =   5295
         Begin VB.ComboBox cmbNpc 
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
            Index           =   9
            Left            =   2760
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   360
            Width           =   2295
         End
         Begin VB.ComboBox cmbNpc 
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
            Index           =   8
            Left            =   2760
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   720
            Width           =   2295
         End
         Begin VB.ComboBox cmbNpc 
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
            Index           =   7
            Left            =   2760
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   1140
            Width           =   2295
         End
         Begin VB.ComboBox cmbNpc 
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
            Index           =   6
            ItemData        =   "frmMapProperties.frx":238E
            Left            =   2760
            List            =   "frmMapProperties.frx":2390
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   1560
            Width           =   2295
         End
         Begin VB.ComboBox cmbNpc 
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
            Index           =   5
            Left            =   2760
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   1920
            Width           =   2295
         End
         Begin VB.ComboBox cmbNpc 
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
            Index           =   0
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   360
            Width           =   2295
         End
         Begin VB.ComboBox cmbNpc 
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
            Index           =   1
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   720
            Width           =   2295
         End
         Begin VB.ComboBox cmbNpc 
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
            Index           =   2
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   1140
            Width           =   2295
         End
         Begin VB.ComboBox cmbNpc 
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
            Index           =   3
            ItemData        =   "frmMapProperties.frx":2392
            Left            =   240
            List            =   "frmMapProperties.frx":2394
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   1560
            Width           =   2295
         End
         Begin VB.ComboBox cmbNpc 
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
            Index           =   4
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   1920
            Width           =   2295
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Map Theme"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   2160
         TabIndex        =   19
         Top             =   2280
         Width           =   3255
         Begin VB.CommandButton cmdstop 
            Caption         =   "Stop"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   2400
            TabIndex        =   35
            Top             =   720
            Width           =   735
         End
         Begin VB.CommandButton cmdplay 
            Caption         =   "Test"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1560
            TabIndex        =   34
            Top             =   720
            Width           =   735
         End
         Begin VB.HScrollBar scrlMusic 
            Height          =   255
            Left            =   720
            Max             =   255
            TabIndex        =   20
            Top             =   360
            Value           =   1
            Width           =   1935
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "Music:"
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
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   495
         End
         Begin VB.Label lblMusic 
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
            Left            =   2640
            TabIndex        =   21
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Map Generalisations"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   2160
         TabIndex        =   14
         Top             =   720
         Width           =   3255
         Begin VB.ComboBox cmbMoral 
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
            ItemData        =   "frmMapProperties.frx":2396
            Left            =   720
            List            =   "frmMapProperties.frx":23A0
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   480
            Width           =   2415
         End
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
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   960
            Width           =   2415
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Moral:"
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
            Left            =   120
            TabIndex        =   18
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label12 
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
            Left            =   120
            TabIndex        =   17
            Top             =   960
            Width           =   495
         End
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
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   4695
      End
      Begin VB.Frame Frame1 
         Caption         =   "Map Switchovers"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1935
         Begin VB.TextBox txtLeft 
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
            TabIndex        =   9
            Text            =   "0"
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox txtRight 
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
            TabIndex        =   8
            Text            =   "0"
            Top             =   1680
            Width           =   735
         End
         Begin VB.TextBox txtDown 
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
            TabIndex        =   7
            Text            =   "0"
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox txtUp 
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
            TabIndex        =   6
            Text            =   "0"
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Right:"
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
            TabIndex        =   13
            Top             =   1680
            Width           =   495
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Left:"
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
            TabIndex        =   12
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Down:"
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
            TabIndex        =   11
            Top             =   960
            Width           =   495
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Up:"
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
            TabIndex        =   10
            Top             =   600
            Width           =   375
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
         Left            =   120
         TabIndex        =   2
         Top             =   6000
         Width           =   1095
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
         Left            =   4320
         TabIndex        =   1
         Top             =   6000
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
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
         Left            =   4920
         TabIndex        =   4
         Top             =   360
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmMapProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPlay_Click()
Call StopMidi
Call PlayMidi("music" & scrlMusic.Value & ".mid")
End Sub

Private Sub cmdstop_Click()
Call StopMidi
End Sub

Private Sub Form_Load()
Dim X As Long, y As Long, i As Long

    txtName.Text = Trim(Map.Name)
    txtUp.Text = STR(Map.Up)
    txtDown.Text = STR(Map.Down)
    txtLeft.Text = STR(Map.Left)
    txtRight.Text = STR(Map.Right)
    cmbMoral.ListIndex = Map.Moral
    scrlMusic.Value = Map.Music
 '   txtBootMap.Text = STR(Map.BootMap)
 '   txtBootX.Text = STR(Map.BootX)
 '   txtBootY.Text = STR(Map.BootY)
    
    cmbShop.AddItem "No Shop"
    For X = 1 To MAX_SHOPS
        cmbShop.AddItem X & ": " & Trim(Shop(X).Name)
    Next X
    cmbShop.ListIndex = Map.Shop
    
    For X = 1 To MAX_MAP_NPCS
        cmbNpc(X - 1).AddItem "No NPC"
    Next X
    
    For y = 1 To MAX_NPCS
        For X = 1 To MAX_MAP_NPCS
            cmbNpc(X - 1).AddItem y & ": " & Trim(Npc(y).Name)
        Next X
    Next y
    
    For i = 1 To MAX_MAP_NPCS
        cmbNpc(i - 1).ListIndex = Map.Npc(i)
    Next i
End Sub

Private Sub scrlMusic_Change()
    lblMusic.Caption = STR(scrlMusic.Value)
End Sub

Private Sub cmdOk_Click()
Dim X As Long, y As Long, i As Long

    Map.Name = txtName.Text
    Map.Up = Val(txtUp.Text)
    Map.Down = Val(txtDown.Text)
    Map.Left = Val(txtLeft.Text)
    Map.Right = Val(txtRight.Text)
    Map.Moral = cmbMoral.ListIndex
    Map.Music = scrlMusic.Value
    'Map.BootMap = Val(txtBootMap.Text)
    'Map.BootX = Val(txtBootX.Text)
    'Map.BootY = Val(txtBootY.Text)
    Map.Shop = cmbShop.ListIndex
    
    Call PlayMidi("music" & Trim(STR(Map.Music)) & ".mid")
    
    For i = 1 To MAX_MAP_NPCS
        Map.Npc(i) = cmbNpc(i - 1).ListIndex
    Next i
    
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Call StopMidi
    Call PlayMidi("music" & Trim(STR(Map.Music)) & ".mid")
    Unload Me
End Sub

