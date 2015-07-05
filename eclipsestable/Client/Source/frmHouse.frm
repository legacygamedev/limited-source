VERSION 5.00
Begin VB.Form frmHouse 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "House Attribute"
   ClientHeight    =   2070
   ClientLeft      =   90
   ClientTop       =   495
   ClientWidth     =   4590
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "House"
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.HScrollBar scrlItem 
         Height          =   255
         Left            =   120
         Max             =   30
         TabIndex        =   4
         Top             =   480
         Width           =   4335
      End
      Begin VB.HScrollBar scrlCost 
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         Max             =   30000
         TabIndex        =   3
         Top             =   1080
         Width           =   4335
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
         Top             =   1560
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
         Left            =   2400
         TabIndex        =   1
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Cost"
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
         Left            =   840
         TabIndex        =   8
         Top             =   240
         Width           =   510
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item:"
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
         Left            =   405
         TabIndex        =   7
         Top             =   240
         Width           =   345
      End
      Begin VB.Label lblCost 
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
         Left            =   840
         TabIndex        =   6
         Top             =   840
         Width           =   1755
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Value:"
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
         Left            =   345
         TabIndex        =   5
         Top             =   840
         Width           =   405
      End
   End
End
Attribute VB_Name = "frmHouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    frmHouse.Visible = False
End Sub

Private Sub cmdOk_Click()
    HouseItem = scrlItem.Value
    HousePrice = scrlCost.Value
    scrlCost.Value = 0
    scrlItem.Value = 0
    frmHouse.Visible = False
End Sub

Private Sub scrlCost_Change()
    lblCost.Caption = scrlCost.Value
End Sub

Private Sub scrlItem_Change()
    If scrlItem.Value = 0 Then
        lblItem.Caption = "No Cost"
        Exit Sub
    Else
        lblItem.Caption = scrlItem.Value & " - " & Trim$(Item(scrlItem.Value).name)
    End If

    If Item(scrlItem.Value).Type = ITEM_TYPE_CURRENCY Then
        scrlCost.Enabled = True
    Else
        scrlCost.Enabled = False
    End If
End Sub

