VERSION 5.00
Begin VB.Form frmDrop 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asphodel Source (Drop Item)"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDrop 
      Height          =   360
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   3855
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   2535
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Amount"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   600
      UseMnemonic     =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Item"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmDrop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ------------------------------------------
' --               Asphodel               --
' ------------------------------------------

Private Sub Form_Load()
Dim InvSlot As Long

    InvSlot = IsItem(InvPosX, InvPosY)
    
    If InvSlot > 0 Then frmDrop.lblName.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, InvSlot)).Name)
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    
        Case vbKeyReturn
            cmdOk_Click
        
        Case vbKeyEscape
            cmdCancel_Click
    
    End Select
End Sub

Private Sub cmdOk_Click()
Dim InvSlot As Long
Dim Amount As Long

    InvSlot = IsItem(InvPosX, InvPosY)
    
    If InvSlot < 1 Then Exit Sub
    
    ' checks if value is numeric and if negative
    If Val(txtDrop.Text) < 1 Then Amount = 1
    
    Amount = txtDrop.Text
    
    If GetPlayerInvItemValue(MyIndex, InvSlot) < Amount Then Amount = GetPlayerInvItemValue(MyIndex, InvSlot)
    
    SendDropItem InvSlot, Amount
    Unload Me
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

