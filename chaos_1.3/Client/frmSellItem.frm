VERSION 5.00
Begin VB.Form frmSellItem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sell Item's"
   ClientHeight    =   5700
   ClientLeft      =   465
   ClientTop       =   660
   ClientWidth     =   3105
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSellItem.frx":0000
   ScaleHeight     =   5700
   ScaleWidth      =   3105
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmUpdate 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox picVisInv 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   0
      Picture         =   "frmSellItem.frx":741A
      ScaleHeight     =   5505
      ScaleWidth      =   3105
      TabIndex        =   2
      Top             =   360
      Width           =   3135
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   23
         Left            =   2160
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   30
         Top             =   3120
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   22
         Left            =   1560
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   29
         Top             =   3120
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   21
         Left            =   960
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   28
         Top             =   3120
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   20
         Left            =   360
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   27
         Top             =   3120
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   19
         Left            =   2160
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   26
         Top             =   2520
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   18
         Left            =   1560
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   25
         Top             =   2520
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   17
         Left            =   960
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   24
         Top             =   2520
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   16
         Left            =   360
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   23
         Top             =   2520
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   15
         Left            =   2160
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   22
         Top             =   1920
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   14
         Left            =   1560
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   21
         Top             =   1920
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   13
         Left            =   960
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   20
         Top             =   1920
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   12
         Left            =   360
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   19
         Top             =   1920
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   11
         Left            =   2160
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   18
         Top             =   1320
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   10
         Left            =   1560
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   17
         Top             =   1320
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   9
         Left            =   960
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   16
         Top             =   1320
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   8
         Left            =   360
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   15
         Top             =   1320
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   7
         Left            =   2160
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   14
         Top             =   720
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   6
         Left            =   1560
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   13
         Top             =   720
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   5
         Left            =   960
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   12
         Top             =   720
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   4
         Left            =   360
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   11
         Top             =   720
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   3
         Left            =   2160
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   10
         Top             =   120
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   2
         Left            =   1560
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   9
         Top             =   120
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   1
         Left            =   960
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   8
         Top             =   120
         Width           =   480
      End
      Begin VB.PictureBox picInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   0
         Left            =   360
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   7
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblGold 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Gold:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   120
         TabIndex        =   31
         Top             =   4680
         Width           =   2865
      End
      Begin VB.Label lblSold 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   0
         TabIndex        =   6
         Top             =   4200
         Width           =   3080
      End
      Begin VB.Label lblPrice 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   0
         TabIndex        =   5
         Top             =   3840
         Width           =   3080
      End
      Begin VB.Label CloseSell 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   2160
         TabIndex        =   4
         Top             =   5040
         Width           =   615
      End
      Begin VB.Label lblSellItem 
         BackStyle       =   0  'Transparent
         Caption         =   "Sell Item"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   240
         TabIndex        =   3
         Top             =   5040
         Width           =   855
      End
      Begin VB.Shape SelectedItemInv 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         Height          =   525
         Left            =   360
         Top             =   105
         Width           =   525
      End
   End
   Begin VB.ListBox lstSellItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   6600
      TabIndex        =   0
      Top             =   5160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer tmrClear 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3360
      Top             =   240
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sell Items"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   -120
      TabIndex        =   1
      Top             =   120
      Width           =   3240
   End
End
Attribute VB_Name = "frmSellItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()
    
        frmSellIItem.lblTitle.Caption = Trim(Map(GetPlayerMap(MyIndex)).name)
        frmSellIItem.lstSellItem.Clear
        frmSellIItem.lstBank.Clear
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(MyIndex, i) > 0 Then
                If Item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Then
                    frmSellIItem.lstInventory.AddItem i & "> " & Trim(Item(GetPlayerInvItemNum(MyIndex, i)).name) & " (" & GetPlayerInvItemValue(MyIndex, i) & ")"
                Else
                    If GetPlayerWeaponSlot(MyIndex) = i Or GetPlayerArmorSlot(MyIndex) = i Or GetPlayerHelmetSlot(MyIndex) = i Or GetPlayerShieldSlot(MyIndex) = i Or GetPlayerLegsSlot(MyIndex) = i Or GetPlayerBootsSlot(MyIndex) = i Or GetPlayerGlovesSlot(MyIndex) = i Or GetPlayerRing1Slot(MyIndex) = i Or GetPlayerRing2Slot(MyIndex) = i Or GetPlayerAmuletSlot(MyIndex) = i Then
                        frmSellIItem.lstInventory.AddItem i & "> " & Trim(Item(GetPlayerInvItemNum(MyIndex, i)).name) & " (worn)"
                    Else
                        frmSellIItem.lstInventory.AddItem i & "> " & Trim(Item(GetPlayerInvItemNum(MyIndex, i)).name)
                    End If
                End If
            Else
                frmSellIItem.lstInventory.AddItem i & "> Empty"
            End If
            DoEvents
        Next i
End Sub

Private Sub lblSellItem_Click()
Dim Packet As String
Dim ItemNum As Long
Dim ItemSlot As Integer

ItemNum = GetPlayerInvItemNum(MyIndex, (lstSellItem.ListIndex + 1))
ItemSlot = lstSellItem.ListIndex
          If GetPlayerInvItemNum(MyIndex, (lstSellItem.ListIndex + 1)) > 0 Then
                   If Item(GetPlayerInvItemNum(MyIndex, (lstSellItem.ListIndex + 1))).Type = ITEM_TYPE_CURRENCY Then
Exit Sub
                Else
                    If GetPlayerWeaponSlot(MyIndex) = (lstSellItem.ListIndex + 1) Or GetPlayerArmorSlot(MyIndex) = (lstSellItem.ListIndex + 1) Or GetPlayerHelmetSlot(MyIndex) = (lstSellItem.ListIndex + 1) Or GetPlayerShieldSlot(MyIndex) = (lstSellItem.ListIndex + 1) Or GetPlayerLegsSlot(MyIndex) = (lstSellItem.ListIndex + 1) Or GetPlayerBootsSlot(MyIndex) = (lstSellItem.ListIndex + 1) Or GetPlayerGlovesSlot(MyIndex) = (lstSellItem.ListIndex + 1) Or GetPlayerRing1Slot(MyIndex) = (lstSellItem.ListIndex + 1) Or GetPlayerRing2Slot(MyIndex) = (lstSellItem.ListIndex + 1) Or GetPlayerAmuletSlot(MyIndex) = (lstSellItem.ListIndex + 1) Then
Exit Sub
                    Else
If Item(GetPlayerInvItemNum(MyIndex, (lstSellItem.ListIndex + 1))).Price > 0 Then
Packet = "sellitem" & SEP_CHAR & ItemNum & SEP_CHAR & ItemSlot & SEP_CHAR & END_CHAR
Call SendData(Packet)
lblSold.Caption = "You sold one " & Trim$(Item(ItemNum).name) & "."

tmrClear.Enabled = True
Call UpdatePlayerSellVisInv

Else
Exit Sub
End If
                    End If
                End If
                       Else
Exit Sub
       End If
       
   frmSellItem.lstSellItem.Clear
   For i = 1 To MAX_INV
          If GetPlayerInvItemNum(MyIndex, i) > 0 Then
                   If Item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Then
                    frmSellItem.lstSellItem.AddItem i & "> " & Trim(Item(GetPlayerInvItemNum(MyIndex, i)).name) & " (" & GetPlayerInvItemValue(MyIndex, i) & ")"
                Else
                    If GetPlayerWeaponSlot(MyIndex) = i Or GetPlayerArmorSlot(MyIndex) = i Or GetPlayerHelmetSlot(MyIndex) = i Or GetPlayerShieldSlot(MyIndex) = i Or GetPlayerLegsSlot(MyIndex) = i Or GetPlayerBootsSlot(MyIndex) = i Or GetPlayerGlovesSlot(MyIndex) = i Or GetPlayerRing1Slot(MyIndex) = i Or GetPlayerRing2Slot(MyIndex) = i Or GetPlayerAmuletSlot(MyIndex) = i Then
                        frmSellItem.lstSellItem.AddItem i & "> " & Trim(Item(GetPlayerInvItemNum(MyIndex, i)).name) & " (worn)"
                    Else
                        frmSellItem.lstSellItem.AddItem i & "> " & Trim(Item(GetPlayerInvItemNum(MyIndex, i)).name)
                    End If
                End If
                       Else
           frmSellItem.lstSellItem.AddItem i & "> Empty"
       End If
   Next i
   frmSellItem.lstSellItem.ListIndex = 0
        
        'Call UpdateSelectedSellInvItem(index)
End Sub

Private Sub lstSellItem_Click()
          If GetPlayerInvItemNum(MyIndex, (lstSellItem.ListIndex + 1)) > 0 Then
                   If Item(GetPlayerInvItemNum(MyIndex, (lstSellItem.ListIndex + 1))).Type = ITEM_TYPE_CURRENCY Then
lblPrice.Caption = "Not a valid selection"
                Else
                    If GetPlayerWeaponSlot(MyIndex) = (lstSellItem.ListIndex + 1) Or GetPlayerArmorSlot(MyIndex) = (lstSellItem.ListIndex + 1) Or GetPlayerHelmetSlot(MyIndex) = (lstSellItem.ListIndex + 1) Or GetPlayerShieldSlot(MyIndex) = (lstSellItem.ListIndex + 1) Or GetPlayerLegsSlot(MyIndex) = (lstSellItem.ListIndex + 1) Or GetPlayerBootsSlot(MyIndex) = (lstSellItem.ListIndex + 1) Or GetPlayerGlovesSlot(MyIndex) = (lstSellItem.ListIndex + 1) Or GetPlayerRing1Slot(MyIndex) = (lstSellItem.ListIndex + 1) Or GetPlayerRing2Slot(MyIndex) = (lstSellItem.ListIndex + 1) Or GetPlayerAmuletSlot(MyIndex) = (lstSellItem.ListIndex + 1) Then
lblPrice.Caption = "Please unequip this item first"
                    Else
If Item(GetPlayerInvItemNum(MyIndex, (lstSellItem.ListIndex + 1))).Price > 0 Then
lblPrice.Caption = "Price: " & Item(GetPlayerInvItemNum(MyIndex, (lstSellItem.ListIndex + 1))).Price & " Gold"
Else
lblPrice.Caption = "Not for sale"
End If
                    End If
                End If
                       Else
lblPrice.Caption = "Not a valid selection"
       End If
End Sub
Private Sub Form_Load()
Dim i As Long
Dim Ending As String

    For i = 1 To 3
        If i = 1 Then Ending = ".GIF"
        If i = 2 Then Ending = ".JPG"
        If i = 3 Then Ending = ".PNG"
 
        If FileExist("GUI\SellItem" & Ending) Then frmChars.Picture = LoadPicture(App.Path & "\GUI\SellItem" & Ending)
    Next i
    lblSold.Caption = ""
    lblPrice.Caption = ""
End Sub

Private Sub tmrClear_Timer()
lblSold.Caption = ""

End Sub

Private Sub CloseSell_Click()
Unload Me
End Sub

Private Sub picInv_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    SelectedItemInv.Top = picInv(Index).Top - 15
    SelectedItemInv.Left = picInv(Index).Left - 15
    
    If Button = 1 Then
        Call UpdateSelectedSellInvItem(Index)
        lstSellItem.ListIndex = Index
    End If
End Sub

Private Sub tmUpdate_Timer()
    Call UpdatePlayerSellVisInv
End Sub
