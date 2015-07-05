VERSION 5.00
Begin VB.Form frmEditShop 
   BorderStyle     =   0  'None
   Caption         =   "Shop Editor"
   ClientHeight    =   5775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   ScaleHeight     =   385
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   353
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox txtShopName 
      Height          =   285
      Left            =   1080
      TabIndex        =   11
      Top             =   240
      Width           =   3855
   End
   Begin VB.CheckBox chkShopFixesItems 
      Caption         =   "Shop Fixes Items"
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   600
      Width           =   4335
   End
   Begin VB.ComboBox cmbShopItemGive 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1440
      Width           =   3855
   End
   Begin VB.ComboBox cmbShopItemGet 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2520
      Width           =   3855
   End
   Begin VB.TextBox txtShopGiveValue 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Text            =   "1"
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox txtShopGetValue 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Text            =   "1"
      Top             =   2880
      Width           =   1575
   End
   Begin VB.ListBox lstShopTradeItem 
      Height          =   1620
      ItemData        =   "frmEditShop.frx":0000
      Left            =   240
      List            =   "frmEditShop.frx":001C
      TabIndex        =   5
      Top             =   3360
      Width           =   4815
   End
   Begin VB.CommandButton cmdShopOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   5160
      Width           =   1935
   End
   Begin VB.CommandButton cmdShopCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   5160
      Width           =   1935
   End
   Begin VB.CommandButton cmdShopUpdate 
      Caption         =   "Update"
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   3000
      Width           =   2055
   End
   Begin VB.HScrollBar scrlGetItem 
      Height          =   255
      Left            =   1080
      Max             =   255
      Min             =   1
      TabIndex        =   1
      Top             =   2160
      Value           =   1
      Width           =   3615
   End
   Begin VB.HScrollBar scrlGiveItem 
      Height          =   255
      Left            =   1080
      Max             =   255
      Min             =   1
      TabIndex        =   0
      Top             =   1080
      Value           =   1
      Width           =   3615
   End
   Begin VB.Label Label7 
      Caption         =   "Value"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Item Get"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Get Item"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Value"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Item Give"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Give Item"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   240
      Width           =   735
   End
   Begin VB.Label lblGetItem 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Left            =   4800
      TabIndex        =   13
      Top             =   2160
      Width           =   90
   End
   Begin VB.Label lblGiveItem 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Left            =   4800
      TabIndex        =   12
      Top             =   1080
      Width           =   90
   End
End
Attribute VB_Name = "frmEditShop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   This file is part of the Cerberus Engine 2nd Edition.
'
'    The Cerberus Engine 2nd Edition is free software; you can redistribute it
'    and/or modify it under the terms of the GNU General Public License as
'    published by the Free Software Foundation; either version 2 of the License,
'    or (at your option) any later version.
'
'    Cerberus 2nd Edition is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Cerberus 2nd Edition; if not, write to the Free Software
'    Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301  USA

Option Explicit

Private Sub cmdShopOk_Click()
    Call ShopEditorOk
End Sub

Private Sub cmdShopCancel_Click()
    Call ShopEditorCancel
End Sub

Private Sub lstShopTradeItem_Click()
    cmbShopItemGive.ListIndex = Shop(EditorIndex).TradeItem(lstShopTradeItem.ListIndex + 1).GiveItem(scrlGiveItem.Value)
    txtShopGiveValue.Text = Shop(EditorIndex).TradeItem(lstShopTradeItem.ListIndex + 1).GiveValue(scrlGiveItem.Value)
    lblGiveItem.Caption = scrlGiveItem.Value
    cmbShopItemGet.ListIndex = Shop(EditorIndex).TradeItem(lstShopTradeItem.ListIndex + 1).GetItem(scrlGetItem.Value)
    txtShopGetValue.Text = Shop(EditorIndex).TradeItem(lstShopTradeItem.ListIndex + 1).GetValue(scrlGetItem.Value)
    lblGetItem.Caption = scrlGetItem.Value
End Sub

Private Sub scrlGiveItem_Change()
    cmbShopItemGive.ListIndex = Shop(EditorIndex).TradeItem(lstShopTradeItem.ListIndex + 1).GiveItem(scrlGiveItem.Value)
    txtShopGiveValue.Text = Shop(EditorIndex).TradeItem(lstShopTradeItem.ListIndex + 1).GiveValue(scrlGiveItem.Value)
    lblGiveItem.Caption = scrlGiveItem.Value
End Sub

Private Sub cmbShopItemGive_Click()
    If lstShopTradeItem.ListIndex >= 0 Then
        Shop(EditorIndex).TradeItem(lstShopTradeItem.ListIndex + 1).GiveItem(scrlGiveItem.Value) = cmbShopItemGive.ListIndex
    End If
End Sub

Private Sub txtShopGiveValue_Change()
    Shop(EditorIndex).TradeItem(lstShopTradeItem.ListIndex + 1).GiveValue(scrlGiveItem.Value) = Val(txtShopGiveValue.Text)
End Sub

Private Sub scrlGetItem_Change()
    cmbShopItemGet.ListIndex = Shop(EditorIndex).TradeItem(lstShopTradeItem.ListIndex + 1).GetItem(scrlGetItem.Value)
    txtShopGetValue.Text = Shop(EditorIndex).TradeItem(lstShopTradeItem.ListIndex + 1).GetValue(scrlGetItem.Value)
    lblGetItem.Caption = scrlGetItem.Value
End Sub

Private Sub cmbShopItemGet_Click()
    If lstShopTradeItem.ListIndex >= 0 Then
        Shop(EditorIndex).TradeItem(lstShopTradeItem.ListIndex + 1).GetItem(scrlGetItem.Value) = cmbShopItemGet.ListIndex
    End If
End Sub

Private Sub txtShopGetValue_Change()
    Shop(EditorIndex).TradeItem(lstShopTradeItem.ListIndex + 1).GetValue(scrlGetItem.Value) = Val(txtShopGetValue.Text)
End Sub

Private Sub cmdShopUpdate_Click()
'Dim Index As Long
'Dim i As Long

    'Index = lstShopTradeItem.ListIndex + 1
    'For i = 1 To MAX_GIVE_ITEMS
        'Shop(EditorIndex).TradeItem(Index).GiveItem(scrlGiveItem.Value) = cmbShopItemGive.ListIndex
        'Shop(EditorIndex).TradeItem(Index).GiveValue(scrlGiveItem.Value) = Val(txtShopGiveValue.Text)
    'Next i
    'For i = 1 To MAX_GET_ITEMS
        'Shop(EditorIndex).TradeItem(Index).GetItem(scrlGetItem.Value) = cmbShopItemGet.ListIndex
        'Shop(EditorIndex).TradeItem(Index).GetValue(scrlGetItem.Value) = Val(txtShopGetValue.Text)
    'Next i
    
    Call UpdateShopTrade
End Sub
