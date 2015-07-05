VERSION 5.00
Begin VB.Form frmSellItem 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Sell Items"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   -45
   ClientWidth     =   6000
   ControlBox      =   0   'False
   Icon            =   "frmSellItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSellItem.frx":014A
   ScaleHeight     =   6000
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrClear 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   4200
      Top             =   2280
   End
   Begin VB.ListBox lstSellItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      ItemData        =   "frmSellItem.frx":7E4B
      Left            =   720
      List            =   "frmSellItem.frx":7E52
      TabIndex        =   3
      Top             =   1080
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
      Left            =   720
      TabIndex        =   5
      Top             =   5280
      Width           =   3255
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
      Left            =   720
      TabIndex        =   4
      Top             =   5040
      Width           =   3255
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
      Left            =   720
      TabIndex        =   2
      Top             =   840
      Width           =   3240
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
      Left            =   720
      TabIndex        =   1
      Top             =   4800
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
      Left            =   3360
      TabIndex        =   0
      Top             =   4800
      Width           =   615
   End
End
Attribute VB_Name = "frmSellItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CloseSell_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim i As Long
Dim Ending As String
    For i = 1 To 3
        If i = 1 Then Ending = ".GIF"
        If i = 2 Then Ending = ".JPG"
        If i = 3 Then Ending = ".PNG"
 
        If FileExist("GUI\CharacterSelect" & Ending) Then frmChars.Picture = LoadPicture(App.Path & "\GUI\CharacterSelect" & Ending)
    Next i
    lblSold.Caption = ""
    lblPrice.Caption = ""
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
lblSold.Caption = "You sold one " & Trim$(Item(ItemNum).name) & "."
tmrClear.Enabled = True
Else
Exit Sub
End If
                    End If
                End If
                       Else
Exit Sub
       End If
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

Private Sub Picture1_Click()

End Sub

Private Sub tmrClear_Timer()
lblSold.Caption = ""
End Sub
