VERSION 5.00
Begin VB.Form frmMapWarp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Warp"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   130
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   314
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar scrlX 
      Height          =   255
      Left            =   720
      Max             =   15
      TabIndex        =   5
      Top             =   600
      Width           =   3255
   End
   Begin VB.HScrollBar scrlY 
      Height          =   255
      Left            =   720
      Max             =   11
      TabIndex        =   4
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox txtMap 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   720
      TabIndex        =   1
      Text            =   "1"
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Map"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "X"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   600
      UseMnemonic     =   0   'False
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Y"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   495
   End
   Begin VB.Label lblX 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   4080
      TabIndex        =   7
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblY 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   4080
      TabIndex        =   6
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

' ******************************************
' **            Mirage Source 4           **
' ******************************************

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
    lblX.Caption = CStr(scrlX.Value)
End Sub

Private Sub scrlY_Change()
    lblY.Caption = CStr(scrlY.Value)
End Sub

Private Sub txtMap_Change()
    If Val(txtMap.Text) <= 0 Or Val(txtMap.Text) > MAX_MAPS Then
        txtMap.Text = "1"
    End If
End Sub

