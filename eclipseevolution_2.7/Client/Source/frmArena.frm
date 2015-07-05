VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmArena 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Arena Attribute"
   ClientHeight    =   2175
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   5055
   ControlBox      =   0   'False
   Icon            =   "frmArena.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   3413
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   353
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
      TabCaption(0)   =   "Set Spawn"
      TabPicture(0)   =   "frmArena.frx":0FC2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblY"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblX"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblMap"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "scrlY"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "scrlX"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdOk"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdCancel"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "scrlMap"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.HScrollBar scrlMap 
         Height          =   255
         Left            =   240
         Max             =   30
         TabIndex        =   5
         Top             =   600
         Width           =   4335
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
         Left            =   2640
         TabIndex        =   4
         Top             =   1560
         Width           =   1935
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
         Left            =   240
         TabIndex        =   3
         Top             =   1560
         Width           =   1935
      End
      Begin VB.HScrollBar scrlX 
         Height          =   255
         Left            =   240
         Max             =   30
         TabIndex        =   2
         Top             =   1200
         Width           =   2055
      End
      Begin VB.HScrollBar scrlY 
         Height          =   255
         Left            =   2520
         Max             =   30
         TabIndex        =   1
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label lblMap 
         AutoSize        =   -1  'True
         Caption         =   "Map: 0"
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
         TabIndex        =   8
         Top             =   360
         Width           =   435
      End
      Begin VB.Label lblX 
         AutoSize        =   -1  'True
         Caption         =   "X: 0"
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
         TabIndex        =   7
         Top             =   960
         Width           =   240
      End
      Begin VB.Label lblY 
         AutoSize        =   -1  'True
         Caption         =   "Y: 0"
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
         Left            =   2520
         TabIndex        =   6
         Top             =   960
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmArena"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    scrlMap.Value = 0
    scrlX.Value = 0
    scrlY.Value = 0

    Unload Me
End Sub

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    scrlMap.Max = MAX_MAPS
    scrlX.Max = MAX_MAPX
    scrlY.Max = MAX_MAPY

    If Arena1 < scrlMap.Min Then
        Arena1 = scrlMap.Min
    End If

    scrlMap.Value = Arena1

    If Arena2 < scrlX.Min Then
        Arena2 = scrlX.Min
    End If

    scrlX.Value = Arena2

    If Arena3 < scrlY.Min Then
        Arena3 = scrlY.Min
    End If

    scrlY.Value = Arena3
End Sub

Private Sub scrlMap_Change()
    lblMap.Caption = "Map: " & scrlMap.Value
    Arena1 = scrlMap.Value
End Sub

Private Sub scrlX_Change()
    lblX.Caption = "X: " & scrlX.Value
    Arena2 = scrlX.Value
End Sub

Private Sub scrlY_Change()
    lblY.Caption = "Y: " & scrlY.Value
    Arena3 = scrlY.Value
End Sub
