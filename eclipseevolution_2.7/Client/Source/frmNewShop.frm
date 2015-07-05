VERSION 5.00
Begin VB.Form frmNewShop 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Shop"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4290
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   315
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   286
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picItemInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   3975
      Left            =   1680
      ScaleHeight     =   3945
      ScaleWidth      =   2385
      TabIndex        =   22
      Top             =   120
      Width           =   2415
      Begin VB.Label lblMagicReq 
         BackStyle       =   0  'Transparent
         Caption         =   "Str Req:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label lblDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Desc"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   240
         TabIndex        =   32
         Top             =   3120
         Width           =   1935
      End
      Begin VB.Label lblSpdBonus 
         BackStyle       =   0  'Transparent
         Caption         =   "SpdBonus:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label lblMagiBonus 
         BackStyle       =   0  'Transparent
         Caption         =   "MagiBonus:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label lblDefBonus 
         BackStyle       =   0  'Transparent
         Caption         =   "DefBonus:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label lblAddStr 
         BackStyle       =   0  'Transparent
         Caption         =   "StrBonus:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   2160
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label lblVital 
         BackStyle       =   0  'Transparent
         Caption         =   "Vital Mod:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label lblSpdReq 
         BackStyle       =   0  'Transparent
         Caption         =   "Str Req:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblDefReq 
         BackStyle       =   0  'Transparent
         Caption         =   "Str Req:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblStrReq 
         BackStyle       =   0  'Transparent
         Caption         =   "Str Req:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "-Item Info-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      Height          =   255
      Left            =   3840
      TabIndex        =   14
      Top             =   4320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   4320
      Width           =   255
   End
   Begin VB.PictureBox imgBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   4
      Left            =   120
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   10
      Top             =   3480
      Width           =   540
      Begin VB.PictureBox picEmoticon 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   4
         Left            =   15
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   11
         Top             =   15
         Width           =   480
         Begin VB.PictureBox iconn 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   4
            Left            =   0
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   12
            Top             =   0
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox imgBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   3
      Left            =   120
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   7
      Top             =   2640
      Width           =   540
      Begin VB.PictureBox picEmoticon 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   3
         Left            =   15
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   8
         Top             =   15
         Width           =   480
         Begin VB.PictureBox iconn 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   3
            Left            =   0
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   9
            Top             =   0
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox imgBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   2
      Left            =   120
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   4
      Top             =   1800
      Width           =   540
      Begin VB.PictureBox picEmoticon 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   2
         Left            =   15
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   5
         Top             =   15
         Width           =   480
         Begin VB.PictureBox iconn 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   2
            Left            =   0
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   6
            Top             =   0
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox imgBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   0
      Left            =   120
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   2
      Top             =   120
      Width           =   540
      Begin VB.PictureBox picEmoticon 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   1
         Left            =   15
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   3
         Top             =   15
         Width           =   480
         Begin VB.PictureBox iconn 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   0
            Left            =   0
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   21
            Top             =   0
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox imgBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   1
      Left            =   120
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   0
      Top             =   960
      Width           =   540
      Begin VB.PictureBox picEmoticon 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   0
         Left            =   15
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   1
         Top             =   15
         Width           =   480
         Begin VB.PictureBox iconn 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   1
            Left            =   0
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   20
            Top             =   0
            Width           =   480
         End
      End
   End
   Begin VB.Label lblSell 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Sell Items"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   35
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label lblFix 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fix Items"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2760
      TabIndex        =   34
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label lblPage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Page: X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1560
      TabIndex        =   33
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label lblItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   4
      Left            =   840
      TabIndex        =   19
      Top             =   3600
      Width           =   3255
   End
   Begin VB.Label lblItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   3
      Left            =   840
      TabIndex        =   18
      Top             =   2760
      Width           =   3255
   End
   Begin VB.Label lblItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   2
      Left            =   840
      TabIndex        =   17
      Top             =   1920
      Width           =   3255
   End
   Begin VB.Label lblItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   840
      TabIndex        =   16
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Label lblItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   840
      TabIndex        =   15
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "frmNewShop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private numItems As Integer
Private pageIndex As Integer
Public shopNum As Integer
Public fixItems As Boolean 'Is the shop fixing items?
Public SellItems As Boolean 'Is the shop selling items?
Public maxpages As Long

' Loads shop data into the form for the first time.
Public Sub loadShop(ByVal sNum As Integer)
    Dim i As Integer
    numItems = 0
    pageIndex = 0
    shopNum = sNum
    cmdBack.Visible = False

    Me.Caption = Shop(sNum).Name

    ' Check to see if there are more pages
    For i = 1 To MAX_SHOP_ITEMS
        If Shop(shopNum).ShopItem(i).ItemNum > 0 Then
            numItems = numItems + 1
        End If
    Next i

    maxpages = numItems / 5

    If numItems > 5 Then
        cmdNext.Visible = True
    Else
        cmdNext.Visible = False
    End If

    ' Check if this shop fixes items
    If Shop(sNum).FixesItems = YES Then
        lblFix.Visible = True
    Else
        lblFix.Visible = False
    End If

    ' Check if this shop buys back items
    If Shop(sNum).BuysItems = YES Then
        lblSell.Visible = True
    Else
        lblSell.Visible = False
    End If

    ' Set it not to fix item mode by default
    fixItems = False

End Sub

' Shows the specified page
Public Sub showPage(ByVal page As Integer)
    Dim i As Integer
    Dim itemName As String
    Dim shopCurrency As String

    On Error GoTo showPage_Error

    lblPage.Caption = "Page: " & (page + 1)

    For i = 1 To 5
        If Shop(shopNum).ShopItem(page * 5 + i).ItemNum = 0 Then
            imgBox(i - 1).Visible = False
            lblItem(i - 1).Visible = False
        Else
            imgBox(i - 1).Visible = True
            lblItem(i - 1).Visible = True

            itemName = Trim$(Item(Shop(shopNum).ShopItem(pageIndex * 5 + i).ItemNum).Name)
            shopCurrency = Trim$(Item(Shop(shopNum).currencyItem).Name)
            lblItem(i - 1).Caption = itemName & vbNewLine & "Price: " & STR(Shop(shopNum).ShopItem(pageIndex * 5 + i).Price) & " " & shopCurrency

            Me.iconn(i - 1).Cls

            Call BltIcon(i - 1, Shop(shopNum).ShopItem(pageIndex * 5 + i).ItemNum)
        End If
    Next i

    ' If numItems / 5 - (pageIndex * 5) > 1 Then
    If page < maxpages Then
        cmdNext.Visible = True
    Else
        cmdNext.Visible = False
    End If

    If pageIndex > 0 Then
        cmdBack.Visible = True
    Else
        cmdBack.Visible = False
    End If

    On Error GoTo 0
    Exit Sub

showPage_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure showPage of Form frmNewShop"
    If MsgBox("Could not show page.", vbRetryCancel) = vbRetry Then
        Call showPage(page)
    Else
        frmNewShop.Visible = False
    End If
    Exit Sub
End Sub

Public Sub Buy(shopItemIndex As Integer)
    ' Send buy request to server
    Call SendData("buy" & SEP_CHAR & shopNum & SEP_CHAR & shopItemIndex & END_CHAR)
End Sub

Public Sub FixItem(ByVal Item As Integer)
    Call SendData("FIXITEM" & SEP_CHAR & snumber & SEP_CHAR & Item & END_CHAR)
End Sub

Public Sub Buyback(ByVal Item As Integer, ByVal slot As Integer, Optional ByVal AMT As Integer = 1)
    Call SendData("SELLITEM" & SEP_CHAR & shopNum & SEP_CHAR & Item & SEP_CHAR & slot & SEP_CHAR & AMT & END_CHAR)
End Sub

' Draws icons to teh boxx0r
Private Sub BltIcon(ByVal iconNum As Integer, ByVal ItemNum As Integer)
    On Error Resume Next
    Dim itemX As Integer, itemY As Integer

    ItemNum = Shop(shopNum).ShopItem(pageIndex * 5 + iconNum + 1).ItemNum


    itemX = (Item(ItemNum).Pic - Int(Item(ItemNum).Pic / 6) * 6) * PIC_X
    itemY = Int(Item(ItemNum).Pic / 6) * PIC_Y

    Call BitBlt(iconn(iconNum).hDC, 0, 0, 32, 32, frmMirage.picItems.hDC, itemX, itemY, SRCCOPY)

    ' Clear any errors
    Err.Clear

End Sub

Private Sub ShowItemInfo(ByVal itemN As Integer)

    picItemInfo.Visible = True

    If Item(itemN).StrReq > 0 Then
        lblStrReq.Caption = STAT1 & " Req: " & Item(itemN).StrReq
    Else
        lblStrReq.Caption = STAT1 & " Req: None"
    End If
    If Item(itemN).DefReq > 0 Then
        lblDefReq.Caption = STAT2 & " Req: " & Item(itemN).DefReq
    Else
        lblDefReq.Caption = STAT2 & " Req: None"
    End If
    If Item(itemN).MagicReq > 0 Then
        lblMagicReq.Caption = STAT3 & " Req: " & Item(itemN).MagicReq
    Else
        lblMagicReq.Caption = STAT3 & " Req: None"
    End If
    If Item(itemN).SpeedReq > 0 Then
        lblSpdReq.Caption = STAT4 & " Req: " & Item(itemN).SpeedReq
    Else
        lblSpdReq.Caption = STAT4 & " Req: None"
    End If

    If Item(itemN).Type = ITEM_TYPE_WEAPON Then
        lblVital.Caption = "Attack: " & Item(itemN).Data2
    ElseIf Item(itemN).Type >= ITEM_TYPE_ARMOR And Item(itemN).Type <= ITEM_TYPE_LEGS Then
        lblVital.Caption = "Defense: " & Item(itemN).Data2
    ElseIf Item(itemN).Type >= ITEM_TYPE_POTIONADDHP And Item(itemN).Type <= ITEM_TYPE_POTIONSUBSP Then
        lblVital.Caption = "Value: " & Item(itemN).Data2
    Else
        lblVital.Caption = vbNullString
    End If

    If Item(itemN).AddSTR > 0 Then
        lblAddStr.Caption = STAT1 & " Bonus: " & Item(itemN).AddSTR
    Else
        lblAddStr.Caption = STAT1 & " Bonus: None"
    End If
    If Item(itemN).AddDEF > 0 Then
        lblDefBonus.Caption = STAT2 & " Bonus: " & Item(itemN).AddDEF
    Else
        lblDefBonus.Caption = STAT2 & " Bonus: None"
    End If
    If Item(itemN).AddMAGI > 0 Then
        lblMagiBonus.Caption = STAT3 & " Bonus: " & Item(itemN).AddMAGI
    Else
        lblMagiBonus.Caption = STAT3 & " Bonus: None"
    End If
    If Item(itemN).AddSpeed > 0 Then
        lblSpdBonus.Caption = STAT4 & " Bonus: " & Item(itemN).AddSpeed
    Else
        lblSpdBonus.Caption = STAT4 & " Bonus: None"
    End If

    lblDesc.Caption = Item(itemN).desc

End Sub

Private Sub HideItemInfo()
    picItemInfo.Visible = False
End Sub

Private Sub lblFix_Click()
    frmFixItem.Show vbModal
End Sub

Private Sub lblSell_Click()
    frmSellItem.Visible = True
End Sub

Private Sub cmdBack_Click()
    pageIndex = pageIndex - 1
    Call showPage(pageIndex)
End Sub

Private Sub cmdNext_Click()
    pageIndex = pageIndex + 1
    Call showPage(pageIndex)
End Sub

Private Sub Form_Load()
    Dim i As Long
    Dim Ending As String
    
    For i = 1 To 3
        If i = 1 Then
            Ending = ".gif"
        End If
        If i = 2 Then
            Ending = ".jpg"
        End If
        If i = 3 Then
            Ending = ".png"
        End If

        If FileExists("GUI\Shop" & Ending) Then
            Me.Picture = LoadPicture(App.Path & "\GUI\Shop" & Ending)
        End If
    Next i
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Shop(shopNum).ShowInfo = 1 Then
        Call HideItemInfo
    End If
End Sub

' Buy item
Private Sub imgBox_Click(Index As Integer)
    Buy pageIndex * 5 + Index + 1
End Sub

' Buy item
Private Sub iconn_Click(Index As Integer)
    Buy pageIndex * 5 + Index + 1
End Sub

Private Sub iconn_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
    If Shop(shopNum).ShowInfo = 1 Then
        Call ShowItemInfo(Shop(shopNum).ShopItem(pageIndex * 5 + Index + 1).ItemNum)
    End If
End Sub
