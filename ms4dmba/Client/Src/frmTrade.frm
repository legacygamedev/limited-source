VERSION 5.00
Begin VB.Form frmTrade 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4530
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7530
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
   ScaleHeight     =   302
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   502
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstTrade 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   1920
      ItemData        =   "frmTrade.frx":08CA
      Left            =   3240
      List            =   "frmTrade.frx":08CC
      TabIndex        =   1
      Top             =   840
      Width           =   3975
   End
   Begin VB.Label lblFixItem 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Fix Item"
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
      Left            =   6120
      TabIndex        =   4
      Top             =   2880
      UseMnemonic     =   0   'False
      Width           =   915
   End
   Begin VB.Label lblDeal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Deal"
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
      Left            =   3360
      TabIndex        =   3
      Top             =   2880
      Width           =   915
   End
   Begin VB.Label lblCancel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
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
      Left            =   6360
      TabIndex        =   2
      Top             =   4080
      Width           =   915
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Trade"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   240
      UseMnemonic     =   0   'False
      Width           =   2535
   End
End
Attribute VB_Name = "frmTrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ******************************************
' **            Mirage Source 4           **
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
        
        Buffer.WriteLong CTradeRequest
        Buffer.WriteLong lstTrade.ListIndex + 1
        
        SendData Buffer.ToArray()
        
        Set Buffer = Nothing
    
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
