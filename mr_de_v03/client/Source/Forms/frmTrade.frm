VERSION 5.00
Begin VB.Form frmTrade 
   Appearance      =   0  'Flat
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trade"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5460
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTrade.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   351
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   364
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picItemDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5250
      Left            =   3600
      Picture         =   "frmTrade.frx":000C
      ScaleHeight     =   348
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   120
      TabIndex        =   6
      Top             =   0
      Width           =   1830
      Begin VB.Label lblItemName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001D2B34&
         Height          =   255
         Left            =   75
         TabIndex        =   11
         Top             =   600
         Width           =   1680
      End
      Begin VB.Label lblItemDescReq 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "DescReq"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001D2B34&
         Height          =   180
         Left            =   600
         TabIndex        =   10
         Top             =   1425
         Width           =   630
      End
      Begin VB.Label lblRequirement 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Requirement"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001D2B34&
         Height          =   180
         Left            =   435
         TabIndex        =   9
         Top             =   1155
         Width           =   960
      End
      Begin VB.Label lblItemDescName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ItemDescName"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H003E8CA6&
         Height          =   240
         Left            =   60
         TabIndex        =   8
         Top             =   75
         Width           =   1695
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblItemNeeded 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ItemNeeded"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001D2B34&
         Height          =   180
         Left            =   450
         TabIndex        =   7
         Top             =   1680
         Width           =   930
      End
   End
   Begin VB.PictureBox picTrade 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5250
      Left            =   0
      Picture         =   "frmTrade.frx":1EC7E
      ScaleHeight     =   350
      ScaleMode       =   0  'User
      ScaleWidth      =   242
      TabIndex        =   0
      Top             =   0
      Width           =   3630
      Begin VB.PictureBox picSetHome 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2100
         Left            =   240
         Picture         =   "frmTrade.frx":5D010
         ScaleHeight     =   2100
         ScaleWidth      =   3105
         TabIndex        =   1
         Top             =   2160
         Visible         =   0   'False
         Width           =   3105
         Begin VB.Label lblSetHome 
            BackStyle       =   0  'Transparent
            Height          =   495
            Left            =   840
            TabIndex        =   2
            Top             =   480
            Visible         =   0   'False
            Width           =   1440
         End
      End
      Begin VB.Label lblShopName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "lblShopName"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H003E8CA6&
         Height          =   240
         Left            =   180
         TabIndex        =   5
         Top             =   195
         Width           =   3300
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblShopDescription 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "lblShopDescription"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001D2B34&
         Height          =   930
         Left            =   375
         TabIndex        =   4
         Top             =   780
         Width           =   2925
      End
      Begin VB.Label lblBye 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1170
         TabIndex        =   3
         Top             =   4680
         Width           =   1320
      End
   End
End
Attribute VB_Name = "frmTrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    ClearInfo
End Sub

Private Sub Form_Unload(Cancel As Integer)
    InShop = 0
    ClearInfo
End Sub

Private Sub lblBye_Click()
    InShop = 0
    ShopNpcNum = 0
    Unload Me
End Sub

Private Sub lblSetHome_Click()
    SendSetBound
    InShop = 0
    ShopNpcNum = 0
    Unload Me
End Sub

Private Sub picTrade_DblClick()
Dim i As Long
Dim rec_pos As RECT

    If Shop(InShop).Type <> SHOP_TYPE_SHOP Then Exit Sub
    
    For i = 1 To MAX_TRADES
        If Shop(InShop).TradeItem(i).GetItem > 0 Then
            With rec_pos
                .Top = 150 + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                .Bottom = .Top + PIC_Y
                .Left = 31 + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                .Right = .Left + PIC_X
            
                If TradeX >= .Left And TradeX <= .Right Then
                    If TradeY >= .Top And TradeY <= .Bottom Then
                        SendTradeRequest i
                        Exit Sub
                    End If
                End If
            End With
        End If
    Next
End Sub

Private Sub picTrade_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
Dim ItemNum As Long
Dim ItemType As Long
Dim X2 As Long
Dim Y2 As Long
Dim rec_pos As RECT

    If Shop(InShop).Type <> SHOP_TYPE_SHOP Then Exit Sub

    TradeX = X
    TradeY = Y

    For i = 1 To MAX_TRADES
        If Shop(InShop).TradeItem(i).GetItem > 0 Then
            With rec_pos
                .Top = 150 + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                .Bottom = .Top + PIC_Y
                .Left = 31 + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                .Right = .Left + PIC_X
            End With
            
            If X >= rec_pos.Left And X <= rec_pos.Right Then
                If Y >= rec_pos.Top And Y <= rec_pos.Bottom Then
                    
                    ItemNum = Shop(InShop).TradeItem(i).GetItem
                    ItemType = Item(ItemNum).Type
            
                    lblItemDescName.Caption = Trim$(Item(ItemNum).Name)
                    
                    Select Case ItemType
                        Case ITEM_TYPE_NONE
                            lblItemName.Caption = "Item"
                            lblRequirement.Caption = ItemReq(ItemNum)
                            lblItemDescReq.Caption = "Value: " & Shop(InShop).TradeItem(i).GetValue & vbNewLine
            
                        Case ITEM_TYPE_EQUIPMENT
                            lblItemName.Caption = EquipmentName(Item(ItemNum).Data1)
                            lblRequirement.Caption = ItemReq(ItemNum)
                            lblItemDescReq.Caption = ItemDesc(ItemNum)
                            
                        Case ITEM_TYPE_POTION
                            lblItemName.Caption = "Potion"
                            lblRequirement.Caption = ItemReq(ItemNum)
                            lblItemDescReq.Caption = ItemDesc(ItemNum)
                            
                        Case ITEM_TYPE_KEY
                            lblItemName.Caption = "Key"
                            lblRequirement.Caption = ItemReq(ItemNum)
                            lblItemDescReq.Caption = "Value: " & Shop(InShop).TradeItem(i).GetValue & vbNewLine
                        
                        Case ITEM_TYPE_SPELL
                            lblItemName.Caption = "Spell"
                            lblRequirement.Caption = ItemReq(ItemNum)
                            lblItemDescReq.Caption = "Spell" & vbNewLine & Trim$(Spell(Item(ItemNum).Data1).Name) & vbNewLine
                            
                    End Select
                    
                    lblItemNeeded.Caption = "Item Needed For Trade" & vbNewLine & Shop(InShop).TradeItem(i).GiveValue & " " & Trim$(Item(Shop(InShop).TradeItem(i).GiveItem).Name) & vbNewLine
                    
                    lblItemDescReq.Top = lblRequirement.Top + lblRequirement.Height
                    lblItemNeeded.Top = lblItemDescReq.Top + lblItemDescReq.Height
                    Exit Sub
                End If
            End If
        End If
    Next
    
    'ClearInfo
End Sub

Private Sub ClearInfo()
    lblItemDescName.Caption = "No Data"
    lblItemName.Caption = "No Data"
    lblRequirement.Caption = "Requirements" & vbNewLine
    lblItemDescReq.Caption = "Description" & vbNewLine
    lblItemNeeded.Caption = "Item Needed For Trade" & vbNewLine
    lblItemDescReq.Top = lblRequirement.Top + lblRequirement.Height
    lblItemNeeded.Top = lblItemDescReq.Top + lblItemDescReq.Height
End Sub
