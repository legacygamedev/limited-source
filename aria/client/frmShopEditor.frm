VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL3N.OCX"
Begin VB.Form frmShopEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shop Editor"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   5670
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
   ScaleHeight     =   7335
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   2655
      Left            =   240
      TabIndex        =   20
      Top             =   4080
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   4683
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   353
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Weapon"
      TabPicture(0)   =   "frmShopEditor.frx":0FC2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lstTradeItem(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Shield"
      TabPicture(1)   =   "frmShopEditor.frx":0FDE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstTradeItem(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Armor"
      TabPicture(2)   =   "frmShopEditor.frx":0FFA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lstTradeItem(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Helmet"
      TabPicture(3)   =   "frmShopEditor.frx":1016
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lstTradeItem(3)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Spell"
      TabPicture(4)   =   "frmShopEditor.frx":1032
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lstTradeItem(4)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Others"
      TabPicture(5)   =   "frmShopEditor.frx":104E
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "lstTradeItem(5)"
      Tab(5).ControlCount=   1
      Begin VB.ListBox lstTradeItem 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2205
         Index           =   5
         ItemData        =   "frmShopEditor.frx":106A
         Left            =   -74880
         List            =   "frmShopEditor.frx":106C
         TabIndex        =   26
         Top             =   360
         Width           =   4935
      End
      Begin VB.ListBox lstTradeItem 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2205
         Index           =   4
         ItemData        =   "frmShopEditor.frx":106E
         Left            =   -74880
         List            =   "frmShopEditor.frx":1070
         TabIndex        =   25
         Top             =   360
         Width           =   4935
      End
      Begin VB.ListBox lstTradeItem 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2205
         Index           =   3
         ItemData        =   "frmShopEditor.frx":1072
         Left            =   -74880
         List            =   "frmShopEditor.frx":1074
         TabIndex        =   24
         Top             =   360
         Width           =   4935
      End
      Begin VB.ListBox lstTradeItem 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2205
         Index           =   2
         ItemData        =   "frmShopEditor.frx":1076
         Left            =   -74880
         List            =   "frmShopEditor.frx":1078
         TabIndex        =   23
         Top             =   360
         Width           =   4935
      End
      Begin VB.ListBox lstTradeItem 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2205
         Index           =   1
         ItemData        =   "frmShopEditor.frx":107A
         Left            =   -74880
         List            =   "frmShopEditor.frx":107C
         TabIndex        =   22
         Top             =   360
         Width           =   4935
      End
      Begin VB.ListBox lstTradeItem 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2205
         Index           =   0
         ItemData        =   "frmShopEditor.frx":107E
         Left            =   120
         List            =   "frmShopEditor.frx":1080
         TabIndex        =   21
         Top             =   360
         Width           =   4935
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Shop Properties"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   5415
      Begin VB.ComboBox cmbItemGive 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   360
         Width           =   3975
      End
      Begin VB.TextBox txtItemGiveValue 
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
         Left            =   1200
         TabIndex        =   14
         Text            =   "1"
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox cmbItemGet 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1080
         Width           =   3975
      End
      Begin VB.TextBox txtItemGetValue 
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
         Left            =   1200
         TabIndex        =   12
         Text            =   "1"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update "
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
         TabIndex        =   11
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Item Give :"
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
         TabIndex        =   19
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Value :"
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
         TabIndex        =   18
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Item Get :"
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
         TabIndex        =   17
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Value :"
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
         TabIndex        =   16
         Top             =   1440
         Width           =   735
      End
   End
   Begin VB.CheckBox chkFixesItems 
      Caption         =   "Shop Fixes Items"
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
      Top             =   1800
      Width           =   1575
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
      Left            =   120
      TabIndex        =   3
      Top             =   6960
      Width           =   2535
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
      Left            =   3000
      TabIndex        =   4
      Top             =   6960
      Width           =   2535
   End
   Begin VB.TextBox txtLeaveSay 
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
      Left            =   1320
      TabIndex        =   2
      Top             =   1200
      Width           =   3975
   End
   Begin VB.TextBox txtName 
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
      Left            =   1320
      TabIndex        =   0
      Top             =   480
      Width           =   3975
   End
   Begin VB.TextBox txtJoinSay 
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
      Left            =   1320
      TabIndex        =   1
      Top             =   840
      Width           =   3975
   End
   Begin VB.Frame Frame1 
      Caption         =   "General Properties"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   5415
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Join Say :"
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
         TabIndex        =   9
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Name :"
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
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Leave Say :"
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
         Top             =   1080
         Width           =   855
      End
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
Dim Index As Long, i As Long, ItemNum As Long
    
    Index = lstTradeItem(SSTab1.Tab).ListIndex + 1
    i = SSTab1.Tab + 1
    ItemNum = cmbItemGet.ListIndex
    
    If ItemNum > 0 Then
        If i = 1 Then
            If Item(ItemNum).Type = ITEM_TYPE_WEAPON Then
            ElseIf Item(ItemNum).Type = ITEM_TYPE_SHIELD Then
                MsgBox "Please click the tab that says shield to add this!"
                Exit Sub
            ElseIf Item(ItemNum).Type = ITEM_TYPE_ARMOR Then
                MsgBox "Please click the tab that says armor to add this!"
                Exit Sub
            ElseIf Item(ItemNum).Type = ITEM_TYPE_HELMET Then
                MsgBox "Please click the tab that says helmet to add this!"
                Exit Sub
            ElseIf Item(ItemNum).Type = ITEM_TYPE_SPELL Then
                MsgBox "Please click the tab that says spell to add this!"
                Exit Sub
            Else
                MsgBox "Please click the tab that says other to add this!"
                Exit Sub
            End If
        ElseIf i = 2 Then
            If Item(ItemNum).Type = ITEM_TYPE_WEAPON Then
                MsgBox "Please click the tab that says weapon to add this!"
                Exit Sub
            ElseIf Item(ItemNum).Type = ITEM_TYPE_SHIELD Then
            ElseIf Item(ItemNum).Type = ITEM_TYPE_ARMOR Then
                MsgBox "Please click the tab that says armor to add this!"
                Exit Sub
            ElseIf Item(ItemNum).Type = ITEM_TYPE_HELMET Then
                MsgBox "Please click the tab that says helmet to add this!"
                Exit Sub
            ElseIf Item(ItemNum).Type = ITEM_TYPE_SPELL Then
                MsgBox "Please click the tab that says spell to add this!"
                Exit Sub
            Else
                MsgBox "Please click the tab that says other to add this!"
                Exit Sub
            End If
        ElseIf i = 3 Then
            If Item(ItemNum).Type = ITEM_TYPE_WEAPON Then
                MsgBox "Please click the tab that says weapon to add this!"
                Exit Sub
            ElseIf Item(ItemNum).Type = ITEM_TYPE_SHIELD Then
                MsgBox "Please click the tab that says shield to add this!"
                Exit Sub
            ElseIf Item(ItemNum).Type = ITEM_TYPE_ARMOR Then
            ElseIf Item(ItemNum).Type = ITEM_TYPE_HELMET Then
                MsgBox "Please click the tab that says helmet to add this!"
                Exit Sub
            ElseIf Item(ItemNum).Type = ITEM_TYPE_SPELL Then
                MsgBox "Please click the tab that says spell to add this!"
                Exit Sub
            Else
                MsgBox "Please click the tab that says other to add this!"
                Exit Sub
            End If
        ElseIf i = 4 Then
            If Item(ItemNum).Type = ITEM_TYPE_WEAPON Then
                MsgBox "Please click the tab that says weapon to add this!"
                Exit Sub
            ElseIf Item(ItemNum).Type = ITEM_TYPE_SHIELD Then
                MsgBox "Please click the tab that says shield to add this!"
                Exit Sub
            ElseIf Item(ItemNum).Type = ITEM_TYPE_ARMOR Then
                MsgBox "Please click the tab that says armor to add this!"
                Exit Sub
            ElseIf Item(ItemNum).Type = ITEM_TYPE_HELMET Then
            ElseIf Item(ItemNum).Type = ITEM_TYPE_SPELL Then
                MsgBox "Please click the tab that says spell to add this!"
                Exit Sub
            Else
                MsgBox "Please click the tab that says other to add this!"
                Exit Sub
            End If
        ElseIf i = 5 Then
            If Item(ItemNum).Type = ITEM_TYPE_WEAPON Then
                MsgBox "Please click the tab that says weapon to add this!"
                Exit Sub
            ElseIf Item(ItemNum).Type = ITEM_TYPE_SHIELD Then
                MsgBox "Please click the tab that says shield to add this!"
                Exit Sub
            ElseIf Item(ItemNum).Type = ITEM_TYPE_ARMOR Then
                MsgBox "Please click the tab that says armor to add this!"
                Exit Sub
            ElseIf Item(ItemNum).Type = ITEM_TYPE_HELMET Then
                MsgBox "Please click the tab that says helmet to add this!"
                Exit Sub
            ElseIf Item(ItemNum).Type = ITEM_TYPE_SPELL Then
            Else
                MsgBox "Please click the tab that says other to add this!"
                Exit Sub
            End If
        ElseIf i = 6 Then
            If Item(ItemNum).Type = ITEM_TYPE_WEAPON Then
                MsgBox "Please click the tab that says weapon to add this!"
                Exit Sub
            ElseIf Item(ItemNum).Type = ITEM_TYPE_SHIELD Then
                MsgBox "Please click the tab that says shield to add this!"
                Exit Sub
            ElseIf Item(ItemNum).Type = ITEM_TYPE_ARMOR Then
                MsgBox "Please click the tab that says armor to add this!"
                Exit Sub
            ElseIf Item(ItemNum).Type = ITEM_TYPE_HELMET Then
                MsgBox "Please click the tab that says helmet to add this!"
                Exit Sub
            ElseIf Item(ItemNum).Type = ITEM_TYPE_SPELL Then
                MsgBox "Please click the tab that says spell to add this!"
                Exit Sub
            Else
            End If
        End If
    End If
    
    Shop(EditorIndex).TradeItem(SSTab1.Tab + 1).Value(Index).GiveItem = cmbItemGive.ListIndex
    Shop(EditorIndex).TradeItem(SSTab1.Tab + 1).Value(Index).GiveValue = Val(txtItemGiveValue.Text)
    Shop(EditorIndex).TradeItem(SSTab1.Tab + 1).Value(Index).GetItem = cmbItemGet.ListIndex
    Shop(EditorIndex).TradeItem(SSTab1.Tab + 1).Value(Index).GetValue = Val(txtItemGetValue.Text)

    Call UpdateShopTrade
End Sub

