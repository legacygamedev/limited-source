VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmShopEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shop Editor"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5505
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmShopEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   5505
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   11033
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   370
      TabMaxWidth     =   1764
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Info."
      TabPicture(0)   =   "frmShopEditor.frx":0E42
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdOk"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdCancel"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraShop"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraInn"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.Frame fraInn 
         Caption         =   "Inn Bind Point"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         TabIndex        =   20
         Top             =   1680
         Width           =   5055
         Begin VB.TextBox txtMap 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   720
            TabIndex        =   23
            Text            =   "0"
            Top             =   240
            Width           =   4215
         End
         Begin VB.HScrollBar scrlY 
            Height          =   255
            Left            =   720
            Max             =   11
            TabIndex        =   22
            Top             =   960
            Width           =   3735
         End
         Begin VB.HScrollBar scrlX 
            Height          =   255
            Left            =   720
            Max             =   15
            TabIndex        =   21
            Top             =   600
            Width           =   3735
         End
         Begin VB.Label lblY 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4440
            TabIndex        =   28
            Top             =   960
            Width           =   495
         End
         Begin VB.Label lblX 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4440
            TabIndex        =   27
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label9 
            Caption         =   "Map Y :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label8 
            Caption         =   "Map X :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Map :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame fraShop 
         Caption         =   "Shop Info"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Visible         =   0   'False
         Width           =   5055
         Begin VB.ListBox lstTradeItem 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1620
            ItemData        =   "frmShopEditor.frx":0E5E
            Left            =   120
            List            =   "frmShopEditor.frx":0E7A
            TabIndex        =   19
            Top             =   2280
            Width           =   4815
         End
         Begin VB.ComboBox cmbItemGive 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   240
            Width           =   3975
         End
         Begin VB.TextBox txtItemGiveValue 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   960
            TabIndex        =   12
            Text            =   "1"
            Top             =   720
            Width           =   1335
         End
         Begin VB.ComboBox cmbItemGet 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1320
            Width           =   3975
         End
         Begin VB.TextBox txtItemGetValue 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   960
            TabIndex        =   10
            Text            =   "1"
            Top             =   1800
            Width           =   1335
         End
         Begin VB.CommandButton cmdUpdate 
            Caption         =   "Update "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2520
            TabIndex        =   9
            Top             =   1800
            Width           =   2415
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Item Give:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   17
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Value:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   16
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Item Get:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   15
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Value:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   14
            Top             =   1800
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1335
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   5055
         Begin VB.ComboBox cmbType 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmShopEditor.frx":0E9E
            Left            =   1080
            List            =   "frmShopEditor.frx":0EA8
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   960
            Width           =   3855
         End
         Begin VB.TextBox txtJoinSay 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1080
            MaxLength       =   500
            TabIndex        =   5
            Top             =   600
            Width           =   3855
         End
         Begin VB.TextBox txtName 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1080
            TabIndex        =   4
            Top             =   240
            Width           =   3855
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Type:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -480
            TabIndex        =   18
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Message:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -480
            TabIndex        =   7
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Name:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   6
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   2
         Top             =   5760
         Width           =   1095
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   5760
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmShopEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbType_Click()
    fraInn.Visible = False
    fraShop.Visible = False
    
    If (cmbType.ListIndex = SHOP_TYPE_SHOP) Then
        fraShop.Visible = True
    ElseIf (cmbType.ListIndex) = SHOP_TYPE_INN Then
        fraInn.Visible = True
    End If
End Sub

Private Sub cmdOk_Click()
    Call ShopEditorOk
End Sub

Private Sub cmdCancel_Click()
    Call ShopEditorCancel
End Sub

Private Sub cmdUpdate_Click()
Dim Index As Long

    Index = lstTradeItem.ListIndex + 1
    Shop(EditorIndex).TradeItem(Index).GiveItem = cmbItemGive.ListIndex
    Shop(EditorIndex).TradeItem(Index).GiveValue = Val(txtItemGiveValue.Text)
    Shop(EditorIndex).TradeItem(Index).GetItem = cmbItemGet.ListIndex
    Shop(EditorIndex).TradeItem(Index).GetValue = Val(txtItemGetValue.Text)
    
    Call UpdateShopTrade
End Sub

Private Sub scrlX_Change()
    lblX.Caption = scrlX.Value
End Sub

Private Sub scrlY_Change()
    lblY.Caption = scrlY.Value
End Sub

Private Sub txtMap_Change()
    If Not IsNumeric(txtMap.Text) Then txtMap.Text = "0"
    
    If Val(txtMap.Text) > MAX_MAPS Then txtMap.Text = "0"
End Sub
