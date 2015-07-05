VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMapProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Properties"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
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
   ScaleHeight     =   5850
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   9975
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
      TabPicture(0)   =   "frmMapProperties.frx":0E42
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdCancel"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdOk"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtName"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "frmMaxSizes"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      Begin VB.Frame frmMaxSizes 
         Caption         =   "Max Sizes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   32
         Top             =   2520
         Width           =   1935
         Begin VB.TextBox txtMaxY 
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
            TabIndex        =   35
            Text            =   "0"
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtMaxX 
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
            TabIndex        =   33
            Text            =   "0"
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "Max Y"
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
            TabIndex        =   36
            Top             =   630
            Width           =   615
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "Max X"
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
            TabIndex        =   34
            Top             =   270
            Width           =   615
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
         TabIndex        =   17
         Top             =   2400
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
            TabIndex        =   22
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
            TabIndex        =   21
            Top             =   720
            Width           =   735
         End
         Begin VB.HScrollBar scrlMusic 
            Height          =   255
            Left            =   720
            Max             =   255
            TabIndex        =   18
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
            TabIndex        =   20
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
            TabIndex        =   19
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
         Height          =   1575
         Left            =   2160
         TabIndex        =   14
         Top             =   750
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
            ItemData        =   "frmMapProperties.frx":0E5E
            Left            =   720
            List            =   "frmMapProperties.frx":0E68
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   480
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
            TabIndex        =   16
            Top             =   480
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
         Top             =   390
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
         Height          =   1695
         Left            =   120
         TabIndex        =   5
         Top             =   750
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
            Left            =   720
            TabIndex        =   9
            Text            =   "0"
            Top             =   960
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
            Left            =   720
            TabIndex        =   8
            Text            =   "0"
            Top             =   1320
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
            Left            =   720
            TabIndex        =   7
            Text            =   "0"
            Top             =   600
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
            Left            =   720
            TabIndex        =   6
            Text            =   "0"
            Top             =   240
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
            Left            =   120
            TabIndex        =   13
            Top             =   1320
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
            Left            =   240
            TabIndex        =   12
            Top             =   960
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
            Left            =   120
            TabIndex        =   11
            Top             =   600
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
            Left            =   240
            TabIndex        =   10
            Top             =   240
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
         Top             =   5280
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
         Top             =   5280
         Width           =   1095
      End
      Begin VB.Frame Frame5 
         Caption         =   "On Death"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   120
         TabIndex        =   23
         Top             =   3600
         Width           =   5295
         Begin VB.HScrollBar scrlX 
            Height          =   255
            Left            =   960
            Max             =   15
            TabIndex        =   26
            Top             =   720
            Width           =   3615
         End
         Begin VB.HScrollBar scrlY 
            Height          =   255
            Left            =   960
            Max             =   11
            TabIndex        =   25
            Top             =   1080
            Width           =   3615
         End
         Begin VB.TextBox txtMap 
            Alignment       =   1  'Right Justify
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
            TabIndex        =   24
            Text            =   "0"
            Top             =   360
            Width           =   4215
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Map :"
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
            Left            =   300
            TabIndex        =   31
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label8 
            Caption         =   "Map X :"
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
            TabIndex        =   30
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label9 
            Caption         =   "Map Y :"
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
            TabIndex        =   29
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label lblX 
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
            Left            =   4560
            TabIndex        =   28
            Top             =   720
            Width           =   495
         End
         Begin VB.Label lblY 
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
            Left            =   4560
            TabIndex        =   27
            Top             =   1080
            Width           =   495
         End
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
         Top             =   390
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

    txtName.Text = Trim$(Map.Name)
    txtUp.Text = Map.Up
    txtDown.Text = Map.Down
    txtLeft.Text = Map.Left
    txtRight.Text = Map.Right
    cmbMoral.ListIndex = Map.Moral
    scrlMusic.Value = Map.Music
    txtMap.Text = Map.BootMap
    lblX.Caption = Map.BootX
    lblY.Caption = Map.BootY
    scrlX.Value = Map.BootX
    scrlY.Value = Map.BootY
        
    txtMaxX.Text = Map.MaxX
    txtMaxY.Text = Map.MaxY
End Sub

Private Sub scrlMusic_Change()
    lblMusic.Caption = STR$(scrlMusic.Value)
End Sub

Private Sub scrlX_Change()
    lblX.Caption = STR$(scrlX.Value)
End Sub

Private Sub scrlY_Change()
    lblY.Caption = STR$(scrlY.Value)
End Sub

Private Sub txtMap_Change()
    If Val(txtMap.Text) > MAX_MAPS Then
        txtMap.Text = "0"
    End If
End Sub

Private Sub cmdOk_Click()
Dim i As Long
Dim X As Long, X2 As Long
Dim Y As Long, Y2 As Long
Dim tempArr() As TileRec

    If Not IsNumeric(txtMaxX.Text) Then txtMaxX.Text = Map.MaxX
    If Val(txtMaxX.Text) < MAX_MAPX Then txtMaxX.Text = MAX_MAPX
    If Val(txtMaxX.Text) > MAX_BYTE Then txtMaxX.Text = MAX_BYTE
    
    If Not IsNumeric(txtMaxY.Text) Then txtMaxY.Text = Map.MaxY
    If Val(txtMaxY.Text) < MAX_MAPY Then txtMaxY.Text = MAX_MAPY
    If Val(txtMaxY.Text) > MAX_BYTE Then txtMaxY.Text = MAX_BYTE
    
    Map.Name = txtName.Text
    Map.Up = Val(txtUp.Text)
    Map.Down = Val(txtDown.Text)
    Map.Left = Val(txtLeft.Text)
    Map.Right = Val(txtRight.Text)
    Map.Moral = cmbMoral.ListIndex
    Map.Music = scrlMusic.Value
    Map.BootMap = Val(txtMap.Text)
    Map.BootX = Val(lblX.Caption)
    Map.BootY = Val(lblY.Caption)
    
    ' set the data before changing it
    tempArr = Map.Tile
    X2 = Map.MaxX
    Y2 = Map.MaxY
    
    ' change the data
    Map.MaxX = Val(txtMaxX.Text)
    Map.MaxY = Val(txtMaxY.Text)
    ReDim Map.Tile(0 To Map.MaxX, 0 To Map.MaxY)
    
    If X2 > Map.MaxX Then X2 = Map.MaxX
    If Y2 > Map.MaxY Then Y2 = Map.MaxY
    
    For X = 0 To X2
        For Y = 0 To Y2
            Map.Tile(X, Y) = tempArr(X, Y)
        Next
    Next
    
    ClearTempTile

    Call PlayMidi("music" & Trim$(STR$(Map.Music)) & ".mid")
    
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Call StopMidi
    Call PlayMidi("music" & Trim$(STR$(Map.Music)) & ".mid")
    Unload Me
End Sub
