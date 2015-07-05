VERSION 5.00
Begin VB.Form frmTrade 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "M:RPGe (Trade)"
   ClientHeight    =   6540
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   7680
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
   Icon            =   "frmTrade.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   7680
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstTrade 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   2760
      ItemData        =   "frmTrade.frx":0442
      Left            =   2520
      List            =   "frmTrade.frx":0444
      TabIndex        =   0
      Top             =   2880
      Width           =   4575
   End
   Begin VB.Image Image4 
      Height          =   195
      Left            =   7080
      Picture         =   "frmTrade.frx":0446
      Top             =   120
      Width           =   195
   End
   Begin VB.Image Image3 
      Height          =   195
      Left            =   7320
      Picture         =   "frmTrade.frx":25B9
      Top             =   120
      Width           =   195
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   2835
      Left            =   2490
      Top             =   2850
      Width           =   4635
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   0
      Picture         =   "frmTrade.frx":475F
      Top             =   0
      Width           =   7680
   End
   Begin VB.Image picCancel 
      Height          =   480
      Left            =   480
      Picture         =   "frmTrade.frx":732E
      Top             =   3960
      Width           =   1320
   End
   Begin VB.Image picDeal 
      Height          =   480
      Left            =   480
      Picture         =   "frmTrade.frx":780E
      Top             =   3360
      Width           =   1320
   End
   Begin VB.Image picFixItems 
      Height          =   480
      Left            =   480
      Picture         =   "frmTrade.frx":7CA9
      Top             =   2760
      Width           =   1320
   End
   Begin VB.Image Image2 
      Height          =   6195
      Left            =   0
      Picture         =   "frmTrade.frx":810D
      Top             =   360
      Width           =   7680
   End
End
Attribute VB_Name = "frmTrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image3_Click()
Unload Me
End Sub

Private Sub Image4_Click()
frmTrade.WindowState = vbMinimized
End Sub

Private Sub picDeal_Click()
    If lstTrade.ListCount > 0 Then
        Call SendData("traderequest" & SEP_CHAR & lstTrade.ListIndex + 1 & SEP_CHAR & END_CHAR)
    End If
End Sub

Private Sub picFixItems_Click()
Dim i As Long

    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(MyIndex, i) > 0 Then
            frmFixItem.cmbItem.AddItem i & ": " & Trim(Item(GetPlayerInvItemNum(MyIndex, i)).Name)
        Else
            frmFixItem.cmbItem.AddItem "Unused Slot"
        End If
    Next i
    frmFixItem.cmbItem.ListIndex = 0
    frmFixItem.Show vbModal
End Sub

Private Sub picCancel_Click()
    Unload Me
End Sub

