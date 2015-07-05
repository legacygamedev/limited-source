VERSION 5.00
Begin VB.Form frmDrop 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mirage Source Engine (Drop Item)"
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDrop 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1440
      TabIndex        =   5
      Top             =   600
      Width           =   3855
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   2535
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Ammount"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Item"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   2
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

Private Ammount As Long

Private Sub Form_Load()
Dim InvNum As Long
    Ammount = 1
    InvNum = frmMirage.lstInv.ListIndex + 1
    frmDrop.lblName.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, InvNum)).Name)
    frmDrop.txtDrop.Text = "1"
End Sub

Private Sub cmdOk_Click()
Dim InvNum As Long

    InvNum = frmMirage.lstInv.ListIndex + 1
    If txtDrop.Text = vbNullString Then
        Exit Sub
    End If
    If Val(txtDrop.Text) < 1 Then
        Exit Sub
    End If
    Call SendDropItem(InvNum, txtDrop.Text)
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
