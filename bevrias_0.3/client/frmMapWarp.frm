VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMapWarp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Warp"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   6165
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   370
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
      TabCaption(0)   =   "Warp To.."
      TabPicture(0)   =   "frmMapWarp.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblX"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblY"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "scrlX"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "scrlY"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdOk"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdCancel"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtMap"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "fraWarp"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      Begin VB.Frame fraWarp 
         Caption         =   "Access"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   240
         TabIndex        =   11
         Top             =   1680
         Width           =   4335
         Begin VB.TextBox txtMsg 
            Enabled         =   0   'False
            Height          =   285
            Left            =   600
            TabIndex        =   18
            Text            =   "Your level is too low to use the Warp Attribute!"
            Top             =   960
            Width           =   3375
         End
         Begin VB.HScrollBar scrlLevel 
            Enabled         =   0   'False
            Height          =   255
            Left            =   840
            Max             =   25000
            TabIndex        =   13
            Top             =   240
            Width           =   2775
         End
         Begin VB.HScrollBar scrlLevMax 
            Enabled         =   0   'False
            Height          =   255
            Left            =   840
            Max             =   25000
            TabIndex        =   12
            Top             =   600
            Width           =   2775
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Msg"
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
            TabIndex        =   19
            Top             =   1000
            Width           =   255
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lvl Max."
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
            TabIndex        =   17
            Top             =   680
            Width           =   525
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lvl Min."
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
            TabIndex        =   16
            Top             =   270
            Width           =   495
         End
         Begin VB.Label lblLevel 
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
            Left            =   3720
            TabIndex        =   15
            Top             =   270
            Width           =   315
         End
         Begin VB.Label lblLevelMax 
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
            Left            =   3720
            TabIndex        =   14
            Top             =   675
            Width           =   315
         End
      End
      Begin VB.TextBox txtMap 
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
         Left            =   720
         TabIndex        =   5
         Text            =   "1"
         Top             =   480
         Width           =   3855
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
         Left            =   2520
         TabIndex        =   4
         Top             =   3120
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
         Top             =   3120
         Width           =   2055
      End
      Begin VB.HScrollBar scrlY 
         Height          =   255
         Left            =   720
         Max             =   30
         TabIndex        =   2
         Top             =   1320
         Width           =   3255
      End
      Begin VB.HScrollBar scrlX 
         Height          =   255
         Left            =   720
         Max             =   30
         TabIndex        =   1
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label lblY 
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
         Height          =   255
         Left            =   4080
         TabIndex        =   10
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblX 
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
         Height          =   255
         Left            =   4080
         TabIndex        =   9
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Y"
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
         TabIndex        =   8
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "X"
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
         TabIndex        =   7
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Map :"
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
         TabIndex        =   6
         Top             =   480
         Width           =   495
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
    EditorMinLevel = scrlLevel.Value
    EditorMaxLevel = scrlLevMax.Value
    EditorMsg = txtMsg.Text
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    scrlX.Max = MAX_MAPX
    scrlY.Max = MAX_MAPY
    
    If EditorWarpX < scrlX.Min Then EditorWarpX = scrlX.Min
    scrlX.Value = EditorWarpX
    If EditorWarpY < scrlY.Min Then EditorWarpY = scrlY.Min
    scrlY.Value = EditorWarpY
    txtMap.Text = EditorWarpMap
End Sub

Private Sub scrlLevel_Change()
    lblLevel.Caption = str(scrlLevel.Value)
End Sub

Private Sub scrlLevMax_Change()
    lblLevelMax.Caption = str(scrlLevMax.Value)
End Sub

Private Sub scrlX_Change()
    lblX.Caption = str(scrlX.Value)
End Sub

Private Sub scrlY_Change()
    lblY.Caption = str(scrlY.Value)
End Sub

Private Sub txtMap_Change()
    If Val(txtMap.Text) <= 0 Or Val(txtMap.Text) > MAX_MAPS Then
        txtMap.Text = "1"
    End If
End Sub

