VERSION 5.00
Begin VB.Form frmShopEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shop Editor"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmdRestock 
      Height          =   390
      ItemData        =   "frmShopEditor.frx":0000
      Left            =   3840
      List            =   "frmShopEditor.frx":000D
      TabIndex        =   22
      Text            =   "cmbRestock"
      Top             =   4320
      Width           =   1575
   End
   Begin VB.TextBox txtStock 
      Alignment       =   1  'Right Justify
      Height          =   390
      Left            =   1440
      TabIndex        =   20
      Text            =   "1"
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CheckBox chkFixesItems 
      Caption         =   "Fixes Items"
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   1680
      Width           =   5295
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update "
      Height          =   375
      Left            =   3000
      TabIndex        =   17
      Top             =   3840
      Width           =   2415
   End
   Begin VB.TextBox txtItemGetValue 
      Alignment       =   1  'Right Justify
      Height          =   390
      Left            =   1440
      TabIndex        =   6
      Text            =   "1"
      Top             =   3840
      Width           =   1335
   End
   Begin VB.ComboBox cmbItemGet 
      Height          =   390
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3360
      Width           =   3975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   6840
      Width           =   2535
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Top             =   6840
      Width           =   2535
   End
   Begin VB.TextBox txtItemGiveValue 
      Alignment       =   1  'Right Justify
      Height          =   390
      Left            =   1440
      TabIndex        =   4
      Text            =   "1"
      Top             =   2760
      Width           =   1335
   End
   Begin VB.ComboBox cmbItemGive 
      Height          =   390
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2280
      Width           =   3975
   End
   Begin VB.ListBox lstTradeItem 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      ItemData        =   "frmShopEditor.frx":0024
      Left            =   120
      List            =   "frmShopEditor.frx":0040
      TabIndex        =   7
      Top             =   5040
      Width           =   5295
   End
   Begin VB.TextBox txtLeaveSay 
      Height          =   390
      Left            =   1440
      TabIndex        =   2
      Top             =   1080
      Width           =   3975
   End
   Begin VB.TextBox txtName 
      Height          =   390
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
   Begin VB.TextBox txtJoinSay 
      Height          =   390
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   3975
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Restock Time:"
      Height          =   615
      Left            =   2880
      TabIndex        =   21
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "Item Stock"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Value"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Item Get"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Value"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Item Give"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Leave Say"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label14 
      Caption         =   "Join Say"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmShopEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
    Shop(EditorIndex).TradeItem(Index).Stock = CInt(txtStock.Text)
    Shop(EditorIndex).TradeItem(Index).MaxStock = CInt(txtStock.Text)
    Call UpdateShopTrade
End Sub

