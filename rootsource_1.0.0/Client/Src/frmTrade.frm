VERSION 5.00
Begin VB.Form frmTrade 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5550
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3900
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
   Picture         =   "frmTrade.frx":08CA
   ScaleHeight     =   370
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstTrade 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1920
      ItemData        =   "frmTrade.frx":47066
      Left            =   360
      List            =   "frmTrade.frx":47068
      TabIndex        =   0
      Top             =   2160
      Width           =   3255
   End
   Begin VB.Label lblFixItem 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   4320
      UseMnemonic     =   0   'False
      Width           =   315
   End
   Begin VB.Label lblDeal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   4680
      Width           =   795
   End
   Begin VB.Label lblCancel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   5040
      Width           =   1035
   End
End
Attribute VB_Name = "frmTrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ******************************************
' **               rootSource               **
' ******************************************

Private Sub Form_Load()
    Me.Caption = GAME_NAME
    Me.Picture = LoadPicture(App.Path & "/gfx/interface/Menu.bmp")
End Sub

Private Sub lblCancel_Click()
    Unload Me
End Sub

Private Sub lblDeal_Click()
Dim Buffer As clsBuffer

    If lstTrade.ListCount > 0 Then
        Set Buffer = New clsBuffer
        Buffer.PreAllocate 6
        Buffer.WriteInteger CTradeRequest
        Buffer.WriteLong lstTrade.ListIndex + 1
        Call SendData(Buffer.ToArray())
    End If
End Sub

Private Sub lblFixItem_Click()
Dim i As Long

    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(MyIndex, i) > 0 Then
            frmFixItem.cmbItem.AddItem i & ": " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name)
        Else
            frmFixItem.cmbItem.AddItem "<empty slot>"
        End If
    Next
    frmFixItem.cmbItem.ListIndex = 0
    frmFixItem.Show vbModal
End Sub
