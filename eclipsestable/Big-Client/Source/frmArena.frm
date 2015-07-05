VERSION 5.00
Begin VB.Form frmArena 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Arena Attribute"
   ClientHeight    =   2175
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   4830
   ControlBox      =   0   'False
   Icon            =   "frmArena.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4830
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraSetSpawn 
      Caption         =   "Set Spawn"
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.HScrollBar scrlY 
         Height          =   255
         Left            =   2520
         Max             =   30
         TabIndex        =   5
         Top             =   1200
         Width           =   2055
      End
      Begin VB.HScrollBar scrlX 
         Height          =   255
         Left            =   240
         Max             =   30
         TabIndex        =   4
         Top             =   1200
         Width           =   2055
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
         Top             =   1680
         Width           =   1935
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
         TabIndex        =   2
         Top             =   1680
         Width           =   1935
      End
      Begin VB.HScrollBar scrlMap 
         Height          =   255
         Left            =   240
         Max             =   30
         TabIndex        =   1
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label lblY 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         TabIndex        =   8
         Top             =   960
         Width           =   255
      End
      Begin VB.Label lblX 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
      Begin VB.Label lblMap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MapNumber: 0"
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
         TabIndex        =   6
         Top             =   240
         Width           =   930
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
    scrlMap.max = MAX_MAPS
    scrlX.max = MAX_MAPX
    scrlY.max = MAX_MAPY

    If Arena1 < scrlMap.min Then
        Arena1 = scrlMap.min
    End If

    scrlMap.Value = Arena1

    If Arena2 < scrlX.min Then
        Arena2 = scrlX.min
    End If

    scrlX.Value = Arena2

    If Arena3 < scrlY.min Then
        Arena3 = scrlY.min
    End If

    scrlY.Value = Arena3
End Sub

Private Sub scrlMap_Change()
    lblMap.Caption = "MapNumber: " & scrlMap.Value
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
