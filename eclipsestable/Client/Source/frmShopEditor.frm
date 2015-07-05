VERSION 5.00
Begin VB.Form frmShopEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shop Editor"
   ClientHeight    =   4620
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   5910
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
   ScaleHeight     =   4620
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkSellsItems 
      Caption         =   "Shop Buys Items"
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
      Left            =   2040
      TabIndex        =   23
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CheckBox chkShow 
      Caption         =   "Show item info"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3840
      TabIndex        =   7
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Shop Items"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   5655
      Begin VB.Frame frmAddEditItem 
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   5295
         Begin VB.CommandButton cmdAECancel 
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
            Height          =   270
            Left            =   1920
            TabIndex        =   22
            Top             =   1200
            Width           =   1455
         End
         Begin VB.CommandButton cmdAddEdit 
            Caption         =   "Item"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   240
            TabIndex        =   21
            Top             =   1200
            Width           =   1455
         End
         Begin VB.TextBox txtPrice 
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
            TabIndex        =   20
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox txtNumber 
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
            TabIndex        =   19
            Top             =   600
            Width           =   975
         End
         Begin VB.ComboBox cmbItemList 
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
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label Label4 
            Caption         =   "Price:"
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
            Left            =   2040
            TabIndex        =   17
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Number:"
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
            TabIndex        =   16
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label2 
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
            Height          =   135
            Left            =   360
            TabIndex        =   15
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.CommandButton cmdDelItem 
         Caption         =   "Delete Item"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   360
         TabIndex        =   13
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CommandButton cmdEditItem 
         Caption         =   "Edit Item"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   360
         TabIndex        =   12
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add Item"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   360
         TabIndex        =   11
         Top             =   480
         Width           =   1815
      End
      Begin VB.ListBox lstItems 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1545
         Left            =   2880
         TabIndex        =   10
         Top             =   360
         Width           =   2415
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
      TabIndex        =   3
      Top             =   1440
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
      TabIndex        =   1
      Top             =   4200
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
      Left            =   3120
      TabIndex        =   2
      Top             =   4200
      Width           =   2535
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
      Width           =   4095
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
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5655
      Begin VB.ComboBox cmbCurrency 
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
         ItemData        =   "frmShopEditor.frx":0FC2
         Left            =   1200
         List            =   "frmShopEditor.frx":0FC4
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   720
         Width           =   4095
      End
      Begin VB.Label lblCurrency 
         Alignment       =   1  'Right Justify
         Caption         =   "Currency used:"
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
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Name:"
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
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmShopEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private addItem As Boolean

' Temporary array so that we don't modify the shop while editing
Private ShopItemList(1 To MAX_SHOP_ITEMS) As ShopItemRec

' Loads shop item data into our temp array
Public Sub LoadShopItemData(shopNum As Integer)
    Dim i As Integer

    For i = 1 To 25
        ShopItemList(i).Amount = Shop(shopNum).ShopItem(i).Amount
        ShopItemList(i).ItemNum = Shop(shopNum).ShopItem(i).ItemNum
        ShopItemList(i).Price = Shop(shopNum).ShopItem(i).Price
    Next i
End Sub

' Adds the specified item to the shop list and temporary array
Public Sub AddShopItem(ByVal itemN As Integer, ByVal prc As Integer, ByVal cItem As Integer, Optional ByVal AMT As Integer = 0)
    Dim itemStr As String
    If itemN > 0 And itemN <= MAX_ITEMS Then

        If Item(itemN).Stackable = 1 Then
            ' It's stackable so add the amount
            itemStr = AMT & " "
        End If

        ' Add the rest
        itemStr = itemStr & Trim$(Item(itemN).Name) & " for " & prc & " " & Trim$(Item(cItem).Name)

        lstItems.addItem itemStr

        ' Add to the temp array
        ShopItemList(lstItems.ListCount).Amount = AMT
        ShopItemList(lstItems.ListCount).ItemNum = itemN
        ShopItemList(lstItems.ListCount).Price = prc
    End If
End Sub

' Edits the shop item in the list and array
Public Sub EditShopItem(ByVal Index As Integer, ByVal itemN As Integer, ByVal prc As Integer, ByVal cItem As Integer, Optional ByVal AMT As Integer = 0)
    Dim itemStr As String

    If itemN > 0 And itemN <= MAX_ITEMS Then
        If Index >= 0 And Index <= MAX_SHOP_ITEMS Then

            ' Delete the existing item
            Call lstItems.RemoveItem(Index)

            If Item(itemN).Stackable = 1 Then
                ' It's stackable so add the amount
                itemStr = AMT & " "
            End If

            ' Add the rest
            itemStr = itemStr & Trim$(Item(itemN).Name) & " for " & prc & " " & Trim$(Item(cItem).Name)

            lstItems.addItem itemStr, Index

            ' Add to temp array
            ShopItemList(Index + 1).Amount = AMT
            ShopItemList(Index + 1).ItemNum = itemN
            ShopItemList(Index + 1).Price = prc
        End If
    End If
End Sub

Private Sub cmdAdd_Click()
    frmAddEditItem.Visible = True
    frmAddEditItem.Caption = "Add Item"
    cmdAddEdit.Caption = "Add Item"
    addItem = True

    ' Make all the values blank
    cmbItemList.ListIndex = 0
    txtNumber.Text = vbNullString
    txtPrice.Text = vbNullString
End Sub

Private Sub cmdAddEdit_Click()
    Dim currencyItem As Integer

    currencyItem = cmbCurrency.ItemData(cmbCurrency.ListIndex)
    ' Check for invalid input
    If Not IsNumeric(txtNumber.Text) Or Not IsNumeric(txtPrice.Text) Then
        Call MsgBox("Invalid input - please enter a number!", vbExclamation)
    Else
        ' Input was okay
        If addItem Then
            Call AddShopItem(cmbItemList.ListIndex + 1, Val(txtPrice.Text), currencyItem, Val(txtNumber.Text))
        Else
            ' Edit the item - make sure something was selected
            If lstItems.ListIndex >= 0 Then
                Call EditShopItem(lstItems.ListIndex, cmbItemList.ListIndex + 1, Val(txtPrice.Text), cmbCurrency.ListIndex + 1, Val(txtNumber.Text))
            End If
        End If
        frmAddEditItem.Visible = False
    End If

End Sub

Private Sub cmdAECancel_Click()
    frmAddEditItem.Visible = False
End Sub

Private Sub cmdDelItem_Click()
    If lstItems.ListIndex > 0 Then
        ' Remove the item
        Call lstItems.RemoveItem(lstItems.ListIndex)
    End If
End Sub

Private Sub cmdEditItem_Click()
    If lstItems.ListIndex > -1 Then
        frmAddEditItem.Visible = True
        addItem = False
        cmdAddEdit.Caption = "Ok"

        ' Set all the values
        cmbItemList.ListIndex = ShopItemList(lstItems.ListIndex + 1).ItemNum - 1
        txtNumber.Text = ShopItemList(lstItems.ListIndex + 1).Amount
        txtPrice.Text = ShopItemList(lstItems.ListIndex + 1).Price
    Else
        MsgBox "Select an item first!"
    End If
End Sub

Private Sub cmdOk_Click()
    Call ShopEditorOk
End Sub

Private Sub cmdCancel_Click()
    Call ShopEditorCancel
End Sub

' Returns amount of shopitem in the temp array
Public Function GetShopItemAmt(ByVal Item As Integer) As Integer
    If Item > 0 And Item < lstItems.ListCount + 1 Then
        GetShopItemAmt = ShopItemList(Item).Amount
    ElseIf Item < 26 Then
        GetShopItemAmt = 0
    End If
End Function

' Returns item num of shopitem in temp array
Public Function GetShopItemNum(ByVal Item As Integer) As Integer
    If Item > 0 And Item < lstItems.ListCount + 1 Then
        GetShopItemNum = ShopItemList(Item).ItemNum
    ElseIf Item < 26 Then
        GetShopItemNum = 0
    End If
End Function

' Returns item price of shopitem in temp array
Public Function GetShopItemPrice(ByVal Item As Integer) As Integer
    If Item > 0 And Item < lstItems.ListCount + 1 Then
        GetShopItemPrice = ShopItemList(Item).Price
    ElseIf Item < 26 Then
        GetShopItemPrice = 0
    End If
End Function
