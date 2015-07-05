VERSION 5.00
Begin VB.Form frmMapWarp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Warp"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   4710
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmAccess 
      Caption         =   "Access"
      Height          =   1440
      Left            =   45
      TabIndex        =   10
      Top             =   1365
      Width           =   4650
      Begin VB.HScrollBar scrlLevMax 
         Height          =   255
         Left            =   1230
         Max             =   500
         TabIndex        =   16
         Top             =   555
         Width           =   2730
      End
      Begin VB.TextBox txtMsg 
         Height          =   360
         Left            =   630
         TabIndex        =   12
         Text            =   "You are of too low level"
         Top             =   930
         Width           =   3915
      End
      Begin VB.HScrollBar scrlLevel 
         Height          =   255
         Left            =   1260
         Max             =   500
         TabIndex        =   11
         Top             =   255
         Width           =   2700
      End
      Begin VB.Label lblLevelMax 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4020
         TabIndex        =   18
         Top             =   555
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Lev Max"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   17
         Top             =   555
         Width           =   1035
      End
      Begin VB.Label Label4 
         Caption         =   "Msg"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   60
         TabIndex        =   15
         Top             =   930
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Lev min"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   14
         Top             =   255
         Width           =   1020
      End
      Begin VB.Label lblLevel 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4020
         TabIndex        =   13
         Top             =   255
         Width           =   495
      End
   End
   Begin VB.HScrollBar scrlX 
      Height          =   255
      Left            =   720
      Max             =   15
      TabIndex        =   4
      Top             =   600
      Width           =   3255
   End
   Begin VB.HScrollBar scrlY 
      Height          =   255
      Left            =   720
      Max             =   11
      TabIndex        =   3
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   60
      TabIndex        =   2
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2610
      TabIndex        =   1
      Top             =   2880
      Width           =   2055
   End
   Begin VB.TextBox txtMap 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   720
      TabIndex        =   0
      Text            =   "1"
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Map"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
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
   Begin VB.Label lblX 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   6
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblY 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   5
      Top             =   960
      Width           =   495
   End
End
Attribute VB_Name = "frmMapWarp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
    EditorWarpMap = Val(txtMap.text)
    EditorWarpX = scrlX.value
    EditorWarpY = scrlY.value
    EditorMinLevel = scrlLevel.value
    EditorMaxLevel = scrlLevMax.value
    EditorMsg = txtMsg.text
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub




Private Sub Form_Load()
If frmMirage.cmbAtributes.ListIndex = 2 Then
    frmAccess.Visible = True
Else
    frmAccess.Visible = False
End If
End Sub

Private Sub scrlLevel_Change()
    lblLevel.Caption = str(scrlLevel.value)
End Sub

Private Sub scrlLevMax_Change()
    lblLevelMax.Caption = str(scrlLevMax.value)
End Sub

Private Sub scrlX_Change()
    lblX.Caption = str(scrlX.value)
End Sub

Private Sub scrlY_Change()
    lblY.Caption = str(scrlY.value)
End Sub

Private Sub txtMap_Change()
    If Val(txtMap.text) <= 0 Or Val(txtMap.text) > MAX_MAPS Then
        txtMap.text = "1"
    End If
End Sub

