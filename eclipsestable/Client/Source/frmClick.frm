VERSION 5.00
Begin VB.Form frmClick 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OnClick Tile"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   5055
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Set Script"
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin VB.HScrollBar scrlClick 
         Height          =   255
         Left            =   240
         Max             =   100
         TabIndex        =   3
         Top             =   600
         Width           =   4455
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
         TabIndex        =   2
         Top             =   1080
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
         Left            =   2760
         TabIndex        =   1
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Script:"
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
         TabIndex        =   5
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblScript 
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
         Height          =   255
         Left            =   960
         TabIndex        =   4
         Top             =   360
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmClick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    ClickScript = scrlClick.Value
    Unload Me
End Sub

Private Sub Form_Load()
    If ClickScript < scrlClick.min Then
        ClickScript = scrlClick.min
    End If
    scrlClick.Value = ClickScript
    lblScript.Caption = scrlClick.Value
    SendScriptTile ClickScript
End Sub

Private Sub scrlClick_Change()
    lblScript.Caption = scrlClick.Value
    Call SendScriptTile(ClickScript)
End Sub
