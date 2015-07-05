VERSION 5.00
Begin VB.Form frmSellItem 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sell Item's"
   ClientHeight    =   6120
   ClientLeft      =   465
   ClientTop       =   660
   ClientWidth     =   3285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   3285
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   0
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
      Height          =   4905
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
      Left            =   1320
      TabIndex        =   6
      Top             =   5160
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
      Top             =   5160
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
      Top             =   5160
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
      Top             =   5400
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
      Top             =   5760
      Width           =   3255
   End
End
Attribute VB_Name = "frmSellItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Label1_Click()
    Dim i As Long


    ' frmBank.lblBank.Caption = Trim$(Map(GetPlayerMap(MyIndex)).Name)
    frmBank.lstInventory.Clear
    frmBank.lstBank.Clear
    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(MyIndex, i) > 0 Then
            If Item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Then
                frmBank.lstInventory.addItem i & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name) & " (" & GetPlayerInvItemValue(MyIndex, i) & ")"
            Else
                If GetPlayerWeaponSlot(MyIndex) = i Or GetPlayerArmorSlot(MyIndex) = i Or GetPlayerHelmetSlot(MyIndex) = i Or GetPlayerShieldSlot(MyIndex) = i Then
                    frmBank.lstInventory.addItem i & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name) & " (worn)"
                Else
                    frmBank.lstInventory.addItem i & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name)
                End If
            End If
        Else
            frmBank.lstInventory.addItem i & "> Empty"
        End If

    Next i
    frmSellItem.lstSellItem.Clear
    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(MyIndex, i) > 0 Then
            If Item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, i)).Stackable = 1 Then
                frmSellItem.lstSellItem.addItem i & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name) & " (" & GetPlayerInvItemValue(MyIndex, i) & ")"
            Else
                If GetPlayerWeaponSlot(MyIndex) = i Or GetPlayerArmorSlot(MyIndex) = i Or GetPlayerHelmetSlot(MyIndex) = i Or GetPlayerShieldSlot(MyIndex) = i Then
                    frmSellItem.lstSellItem.addItem i & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name) & " (worn)"
                Else
                    frmSellItem.lstSellItem.addItem i & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name)
                End If
            End If
        Else
            frmSellItem.lstSellItem.addItem i & "> Empty"
        End If
    Next i
    frmSellItem.lstSellItem.ListIndex = 0
End Sub

Private Sub lblSellItem_Click()
    Dim packet As String
    Dim ItemNum As Long
    Dim ItemSlot As Integer
    Dim AMT As Long

    ItemNum = GetPlayerInvItemNum(MyIndex, (lstSellItem.ListIndex + 1))
    ItemSlot = lstSellItem.ListIndex + 1
    If GetPlayerInvItemNum(MyIndex, (lstSellItem.ListIndex + 1)) > 0 Then
        If Item(GetPlayerInvItemNum(MyIndex, (lstSellItem.ListIndex + 1))).Type = ITEM_TYPE_CURRENCY Then
            Exit Sub
        Else
            If GetPlayerWeaponSlot(MyIndex) = (lstSellItem.ListIndex + 1) Or GetPlayerArmorSlot(MyIndex) = (lstSellItem.ListIndex + 1) Or GetPlayerHelmetSlot(MyIndex) = (lstSellItem.ListIndex + 1) Or GetPlayerShieldSlot(MyIndex) = (lstSellItem.ListIndex + 1) Then
                Exit Sub
            Else
                If Item(GetPlayerInvItemNum(MyIndex, (lstSellItem.ListIndex + 1))).Price > 0 Then
                    If Item(GetPlayerInvItemNum(MyIndex, (lstSellItem.ListIndex + 1))).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, (lstSellItem.ListIndex + 1))).Stackable = 1 Then
                        AMT = InputBox("How many " & Item(GetPlayerInvItemNum(MyIndex, (lstSellItem.ListIndex + 1))).Name & " would you like to sell?", "Sell " & Trim$(Item(GetPlayerInvItemNum(MyIndex, (lstSellItem.ListIndex + 1))).Name), 0)
                        If IsNumeric(AMT) Then
                            packet = "sellitem" & SEP_CHAR & snumber & SEP_CHAR & ItemNum & SEP_CHAR & ItemSlot & SEP_CHAR & AMT & END_CHAR
                            Call SendData(packet)
                            lblSold.Caption = "You sold " & AMT & " " & Trim$(Item(ItemNum).Name) & "s ."
                        End If
                    Else
                        packet = "sellitem" & SEP_CHAR & snumber & SEP_CHAR & ItemNum & SEP_CHAR & ItemSlot & SEP_CHAR & 1 & END_CHAR
                        Call SendData(packet)
                        lblSold.Caption = "You sold 1 " & Trim$(Item(ItemNum).Name) & "."
                    End If
                    tmrClear.Enabled = True

                Else
                    Exit Sub
                End If
            End If
        End If
    Else
        Exit Sub
    End If
    Timer1.Enabled = True

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
        If i = 1 Then
            Ending = ".GIF"
        End If
        If i = 2 Then
            Ending = ".JPG"
        End If
        If i = 3 Then
            Ending = ".PNG"
        End If

        If FileExists("GUI\SellItem" & Ending) Then
            frmChars.Picture = LoadPicture(App.Path & "\GUI\SellItem" & Ending)
        End If
    Next i
    lblSold.Caption = vbNullString
    lblPrice.Caption = vbNullString
    frmSellItem.lstSellItem.Clear
    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(MyIndex, i) > 0 Then
            If Item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, i)).Stackable = 1 Then
                frmSellItem.lstSellItem.addItem i & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name) & " (" & GetPlayerInvItemValue(MyIndex, i) & ")"
            Else
                If GetPlayerWeaponSlot(MyIndex) = i Or GetPlayerArmorSlot(MyIndex) = i Or GetPlayerHelmetSlot(MyIndex) = i Or GetPlayerShieldSlot(MyIndex) = i Then
                    frmSellItem.lstSellItem.addItem i & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name) & " (worn)"
                Else
                    frmSellItem.lstSellItem.addItem i & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name)
                End If
            End If
        Else
            frmSellItem.lstSellItem.addItem i & "> Empty"
        End If
    Next i
    frmSellItem.lstSellItem.ListIndex = 0
End Sub

Private Sub Timer1_Timer()
    Call Label1_Click
    Timer1.Enabled = False
End Sub

Private Sub tmrClear_Timer()
    lblSold.Caption = vbNullString

End Sub


Private Sub CloseSell_Click()
    Unload Me
End Sub
