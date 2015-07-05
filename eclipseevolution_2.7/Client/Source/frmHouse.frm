VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmHouse 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "House Attribute"
   ClientHeight    =   2175
   ClientLeft      =   90
   ClientTop       =   495
   ClientWidth     =   5055
   ControlBox      =   0   'False
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
      TabCaption(0)   =   "House"
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblItem"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblCost"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "scrlItem"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "scrlCost"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdOk"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdCancel"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
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
         Left            =   360
         TabIndex        =   3
         Top             =   1560
         Width           =   1935
      End
      Begin VB.HScrollBar scrlCost 
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         Max             =   30000
         TabIndex        =   2
         Top             =   1200
         Width           =   4335
      End
      Begin VB.HScrollBar scrlItem 
         Height          =   255
         Left            =   240
         Max             =   30
         TabIndex        =   1
         Top             =   600
         Width           =   4335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Left            =   465
         TabIndex        =   8
         Top             =   960
         Width           =   405
      End
      Begin VB.Label lblCost 
         AutoSize        =   -1  'True
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
         Left            =   960
         TabIndex        =   7
         Top             =   960
         Width           =   1755
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Left            =   525
         TabIndex        =   6
         Top             =   360
         Width           =   345
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
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
         Left            =   960
         TabIndex        =   5
         Top             =   360
         Width           =   510
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
        lblItem.Caption = scrlItem.Value & " - " & Trim$(Item(scrlItem.Value).Name)
    End If

    If Item(scrlItem.Value).Type = ITEM_TYPE_CURRENCY Then
        scrlCost.Enabled = True
    Else
        scrlCost.Enabled = False
    End If
End Sub

