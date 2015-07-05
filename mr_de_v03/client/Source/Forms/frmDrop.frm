VERSION 5.00
Begin VB.Form frmDrop 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Drop Amount"
   ClientHeight    =   1245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDrop.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   3750
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar scrlAmount 
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox txtamount 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      TabIndex        =   3
      Text            =   "1"
      Top             =   480
      Width           =   915
   End
   Begin VB.Label cmdOk 
      BackStyle       =   0  'Transparent
      Caption         =   "Drop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1688
      TabIndex        =   4
      Top             =   945
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Amount:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   510
      Width           =   885
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Item:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   360
      TabIndex        =   1
      Top             =   165
      Width           =   540
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   960
      TabIndex        =   0
      Top             =   180
      Width           =   1890
   End
End
Attribute VB_Name = "frmDrop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Amount As Long

Private Sub Form_Load()
Dim ItemNum As Long
    
    ItemNum = Current_InvItemNum(MyIndex, DropNum)
    
    frmDrop.lblName = Trim$(Item(ItemNum).Name)
    scrlAmount.Min = 1
    scrlAmount.Max = Item(ItemNum).StackMax
    scrlAmount.Value = Current_InvItemValue(MyIndex, DropNum)
End Sub

Private Sub cmdOk_Click()
Dim InvNum As Long

    If IsNumeric(txtamount.Text) Then Amount = txtamount.Text
    InvNum = DropNum

    Call ProcessAmount
    
    Call SendDropItem(InvNum, Amount)
    Unload Me
End Sub

Private Sub ProcessAmount()

    ' Check if more then Max and set back to Max if so
    If Amount > Current_InvItemValue(MyIndex, DropNum) Then Amount = Current_InvItemValue(MyIndex, DropNum)

    ' Make sure its not 0
    If Amount <= 0 Then Amount = 1
End Sub

Private Sub scrlAmount_Change()
    txtamount.Text = scrlAmount.Value
End Sub

Private Sub txtamount_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdOk_Click
    End If
End Sub
