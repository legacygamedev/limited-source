VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMapWarp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Warp"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   ControlBox      =   0   'False
   Icon            =   "frmMapWarp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   3836
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   370
      TabMaxWidth     =   1764
      TabCaption(0)   =   "Prop."
      TabPicture(0)   =   "frmMapWarp.frx":0E42
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdOk"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdCancel"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   255
         Left            =   3960
         TabIndex        =   11
         Top             =   1800
         Width           =   855
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "Ok"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         Width           =   855
      End
      Begin VB.Frame Frame1 
         Height          =   1455
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   4695
         Begin VB.HScrollBar scrlX 
            Height          =   255
            Left            =   720
            Max             =   255
            TabIndex        =   4
            Top             =   720
            Width           =   3255
         End
         Begin VB.HScrollBar scrlY 
            Height          =   255
            Left            =   720
            Max             =   255
            TabIndex        =   3
            Top             =   1080
            Width           =   3255
         End
         Begin VB.TextBox txtMap 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   720
            TabIndex        =   2
            Text            =   "1"
            Top             =   240
            Width           =   3855
         End
         Begin VB.Label Label1 
            Caption         =   "Map"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "Map X:"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Map Y:"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label lblX 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            Height          =   255
            Left            =   4080
            TabIndex        =   6
            Top             =   720
            Width           =   495
         End
         Begin VB.Label lblY 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            Height          =   255
            Left            =   4080
            TabIndex        =   5
            Top             =   1080
            Width           =   495
         End
      End
   End
End
Attribute VB_Name = "frmMapWarp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
    EditorWarpMap = Val(txtMap.Text)
    EditorWarpX = scrlX.Value
    EditorWarpY = scrlY.Value
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub scrlX_Change()
    lblX.Caption = scrlX.Value
End Sub

Private Sub scrlY_Change()
    lblY.Caption = scrlY.Value
End Sub

Private Sub txtMap_Change()
    If Val(txtMap.Text) <= 0 Or Val(txtMap.Text) > MAX_MAPS Then
        txtMap.Text = "1"
    End If
End Sub

