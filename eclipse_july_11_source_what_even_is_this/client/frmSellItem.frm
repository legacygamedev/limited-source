VERSION 5.00
Begin VB.Form frmSellItem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sell Item's"
   ClientHeight    =   4800
   ClientLeft      =   465
   ClientTop       =   660
   ClientWidth     =   3285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   3285
   StartUpPosition =   3  'Windows Default
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
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
   Begin VB.Timer tmrClear 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3360
      Top             =   240
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Refresh"
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
      Height          =   180
      Left            =   1200
      TabIndex        =   6
      Top             =   3960
      Width           =   855
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
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   2640
      TabIndex        =   5
      Top             =   3960
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
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   0
      TabIndex        =   4
      Top             =   3960
      Width           =   855
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   3240
   End
   Begin VB.Label lblPrice 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   4200
      Width           =   3255
   End
   Begin VB.Label lblSold 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   4440
      Width           =   3255
   End
End
Attribute VB_Name = "frmSellItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()
    
        frmBank.lblBank.Caption = Trim$(Map(GetPlayerMap(MyIndex)).Name)
        frmBank.lstInventory.Clear
        frmBank.lstBank.Clear
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(MyIndex, i) > 0 Then
                If Item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Then
                    frmBank.lstInventory.AddItem i & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name) & " (" & GetPlayerInvItemValue(MyIndex, i) & ")"
                Else
                    If GetPlayerWeaponSlot(MyIndex) = i Or GetPlayerArmorSlot(MyIndex) = i Or GetPlayerHelmetSlot(MyIndex) = i Or GetPlayerShieldSlot(MyIndex) = i Then
                        frmBank.lstInventory.AddItem i & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name) & " (worn)"
                    Else
                        frmBank.lstInventory.AddItem i & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name)
                    End If
                End If
            Else
                frmBank.lstInventory.AddItem i & "> Empty"
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
                    If GetPlayerWeaponSlot(MyIndex) = (lstSellItem.ListIndex + 1) Or GetPlayerArmorSlot(MyIndex) = (lstSellItem.ListIndex + 1) Or GetPlayerHelmetSlot(MyIndex) = (lstSellItem.ListIndex + 1) Or GetPlayerShieldSlot(MyIndex) = (lstSellItem.ListIndex + 1) Then
Exit Sub
                    Else
If Item(GetPlayerInvItemNum(MyIndex, (lstSellItem.ListIndex + 1))).Price > 0 Then
Packet = "sellitem" & SEP_CHAR & ItemNum & SEP_CHAR & ItemSlot & SEP_CHAR & END_CHAR
Call SendData(Packet)
lblSold.Caption = "You sold one " & Trim$(Item(ItemNum).Name) & "."

tmrClear.Enabled = True

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
                    frmSellItem.lstSellItem.AddItem i & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name) & " (" & GetPlayerInvItemValue(MyIndex, i) & ")"
                Else
                    If GetPlayerWeaponSlot(MyIndex) = i Or GetPlayerArmorSlot(MyIndex) = i Or GetPlayerHelmetSlot(MyIndex) = i Or GetPlayerShieldSlot(MyIndex) = i Then
                        frmSellItem.lstSellItem.AddItem i & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name) & " (worn)"
                    Else
                        frmSellItem.lstSellItem.AddItem i & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name)
                    End If
                End If
                       Else
           frmSellItem.lstSellItem.AddItem i & "> Empty"
       End If
   Next i
   frmSellItem.lstSellItem.ListIndex = 0
        
End Sub

Private Sub lstSellItem_Click()
          If GetPlayerInvItemNum(MyIndex, (lstSellItem.ListIndex + 1)) > 0 Then
                   If Item(GetPlayerInvItemNum(MyIndex, (lstSellItem.ListIndex + 1))).Type = ITEM_TYPE_CURRENCY Then
lblPrice.Caption = "Not a valid selection"
                Else
                    If GetPlayerWeaponSlot(MyIndex) = (lstSellItem.ListIndex + 1) Or GetPlayerArmorSlot(MyIndex) = (lstSellItem.ListIndex + 1) Or GetPlayerHelmetSlot(MyIndex) = (lstSellItem.ListIndex + 1) Or GetPlayerShieldSlot(MyIndex) = (lstSellItem.ListIndex + 1) Then
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
