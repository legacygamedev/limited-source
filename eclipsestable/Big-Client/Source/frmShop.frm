VERSION 5.00
Begin VB.Form frmShop 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shop"
   ClientHeight    =   1695
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   4815
   ControlBox      =   0   'False
   Icon            =   "frmShop.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   4815
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Set Shop"
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
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
         Left            =   120
         TabIndex        =   3
         Top             =   1320
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
         TabIndex        =   2
         Top             =   1320
         Width           =   1935
      End
      Begin VB.HScrollBar scrlNum 
         Height          =   255
         Left            =   240
         Max             =   30
         Min             =   1
         TabIndex        =   1
         Top             =   480
         Value           =   1
         Width           =   4335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shop Num:"
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
         TabIndex        =   5
         Top             =   240
         Width           =   705
      End
      Begin VB.Label lblNum 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
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
         Left            =   1080
         TabIndex        =   4
         Top             =   240
         Width           =   75
      End
   End
End
Attribute VB_Name = "frmShop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    scrlNum.Value = 1
    Unload Me
End Sub

Private Sub cmdOk_Click()
    EditorShopNum = scrlNum.Value
    scrlNum.Value = 1
    Unload Me
End Sub

Private Sub Form_Load()
    lblNum.Caption = scrlNum.Value & " - " & Trim$(Shop(scrlNum.Value).name)
    If EditorShopNum < scrlNum.min Then
        EditorShopNum = scrlNum.min
    End If
    scrlNum.Value = EditorShopNum
End Sub

Private Sub scrlNum_Change()
    lblNum.Caption = scrlNum.Value & " - " & Trim$(Shop(scrlNum.Value).name)
End Sub
