VERSION 5.00
Begin VB.Form frmInventory 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventory"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3750
   Icon            =   "frmInventory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   3750
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstInv 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   2175
      ItemData        =   "frmInventory.frx":2372
      Left            =   150
      List            =   "frmInventory.frx":2374
      TabIndex        =   0
      Top             =   1140
      Width           =   3450
   End
   Begin VB.Label lblUseItem 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Use Item"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1515
      TabIndex        =   2
      Top             =   3465
      Width           =   750
   End
   Begin VB.Label lblDropItem 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Drop Item"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1530
      TabIndex        =   1
      Top             =   3795
      Width           =   750
   End
End
Attribute VB_Name = "frmInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim i As Long
Dim Ending As String
    For i = 1 To 3
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"

        If FileExist("\core files\interface\inventory" & Ending) Then frmInventory.Picture = LoadPicture(App.Path & "\core files\interface\inventory" & Ending)
    Next i
End Sub

Private Sub lblUseItem_Click()
    Call SendUseItem(frmInventory.lstInv.ListIndex + 1)
End Sub
Private Sub lblDropItem_Click()
Dim Value As Long
Dim InvNum As Long

    InvNum = frmInventory.lstInv.ListIndex + 1

    If GetPlayerInvItemNum(MyIndex, InvNum) > 0 And GetPlayerInvItemNum(MyIndex, InvNum) <= MAX_ITEMS Then
        If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_CURRENCY Then
            ' Show them the drop dialog
            frmDrop.Show vbModal
        Else
            Call SendDropItem(frmInventory.lstInv.ListIndex + 1, 0)
        End If
    End If
End Sub


