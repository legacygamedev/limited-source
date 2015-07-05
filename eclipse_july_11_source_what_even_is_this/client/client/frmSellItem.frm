VERSION 5.00
Begin VB.Form frmSellItem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4620
   ClientLeft      =   135
   ClientTop       =   495
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   3975
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstSellItem 
      Height          =   4155
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label lblSellItem 
      BackStyle       =   0  'Transparent
      Caption         =   "Sell Item"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmSellItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lblSellItem_Click()
Dim Packet As String
Dim ItemNum As Long
Dim ItemSlot As Integer

ItemNum = GetPlayerInvItemNum(MyIndex, (lstSellItem.ListIndex + 1))
ItemSlot = lstSellItem.ListIndex

Packet = "sellitem" & SEP_CHAR & ItemNum & SEP_CHAR & ItemSlot & SEP_CHAR & END_CHAR
Call SendData(Packet)







    frmSellItem.lstSellItem.Clear
   
   For i = 1 To 24
       If GetPlayerInvItemNum(MyIndex, i) > 0 Then
           frmSellItem.lstSellItem.AddItem i & " " & Item(GetPlayerInvItemNum(MyIndex, i)).Name & " - " & Item(GetPlayerInvItemNum(MyIndex, i)).Price
       Else
           frmSellItem.lstSellItem.AddItem "None"
       End If
   Next i
   frmSellItem.lstSellItem.ListIndex = 0
End Sub

