VERSION 5.00
Begin VB.Form frmBank 
   BorderStyle     =   0  'None
   Caption         =   "Bank"
   ClientHeight    =   6000
   ClientLeft      =   4470
   ClientTop       =   2880
   ClientWidth     =   6000
   Icon            =   "frmBank.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmBank.frx":014A
   ScaleHeight     =   6000
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstInventory 
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
      Height          =   4320
      ItemData        =   "frmBank.frx":7838
      Left            =   240
      List            =   "frmBank.frx":783A
      TabIndex        =   1
      Top             =   720
      Width           =   2775
   End
   Begin VB.ListBox lstBank 
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
      Height          =   4320
      Left            =   3000
      TabIndex        =   0
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label lblBank 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BankName"
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
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Width           =   5520
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Deposit"
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
      Left            =   240
      TabIndex        =   5
      Top             =   5040
      Width           =   540
   End
   Begin VB.Label label2 
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
      Left            =   4560
      TabIndex        =   4
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Withdraw"
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
      Left            =   3120
      TabIndex        =   3
      Top             =   5040
      Width           =   690
   End
   Begin VB.Label lblMsg 
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
      Height          =   195
      Left            =   159
      TabIndex        =   2
      Top             =   5270
      Width           =   5640
   End
End
Attribute VB_Name = "frmBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Label1_Click()
    Call BankItems
End Sub

Private Sub Label2_Click()
Call SendData("save" & SEP_CHAR & END_CHAR)
    Unload Me
End Sub

Private Sub Label3_Click()
    Call InvItems
End Sub

Sub BankItems()
Dim InvNum As Long
Dim GoldAmount As String
On Error GoTo Done

    InvNum = lstInventory.ListIndex + 1
    If GetPlayerInvItemNum(MyIndex, InvNum) > 0 And GetPlayerInvItemNum(MyIndex, InvNum) <= MAX_ITEMS Then
        If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_CURRENCY Then
            GoldAmount = InputBox("How much " & Trim(Item(GetPlayerInvItemNum(MyIndex, InvNum)).name) & "(" & GetPlayerInvItemValue(MyIndex, InvNum) & ") would you like to deposit?", "Deposit " & Trim(Item(GetPlayerInvItemNum(MyIndex, InvNum)).name), 0, frmBank.Left, frmBank.Top)
            If IsNumeric(GoldAmount) Then
                Call SendData("bankdeposit" & SEP_CHAR & lstInventory.ListIndex + 1 & SEP_CHAR & GoldAmount & SEP_CHAR & END_CHAR)
            End If
        Else
            Call SendData("bankdeposit" & SEP_CHAR & lstInventory.ListIndex + 1 & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
        End If
    End If
    Exit Sub
Done:
    If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_CURRENCY Then
        lblMsg.Caption = "The variable can't handle that amount!"
    End If
End Sub

Sub InvItems()
Dim BankNum As Long
Dim GoldAmount As String
On Error GoTo Done

    BankNum = lstBank.ListIndex + 1
    If GetPlayerBankItemNum(MyIndex, BankNum) > 0 And GetPlayerBankItemNum(MyIndex, BankNum) <= MAX_ITEMS Then
        If Item(GetPlayerBankItemNum(MyIndex, BankNum)).Type = ITEM_TYPE_CURRENCY Then
            GoldAmount = InputBox("How much " & Trim(Item(GetPlayerBankItemNum(MyIndex, BankNum)).name) & "(" & GetPlayerBankItemValue(MyIndex, BankNum) & ") would you like to withdraw?", "Withdraw " & Trim(Item(GetPlayerBankItemNum(MyIndex, BankNum)).name), 0, frmBank.Left, frmBank.Top)
            If IsNumeric(GoldAmount) Then
                Call SendData("bankwithdraw" & SEP_CHAR & lstBank.ListIndex + 1 & SEP_CHAR & GoldAmount & SEP_CHAR & END_CHAR)
            End If
        Else
            Call SendData("bankwithdraw" & SEP_CHAR & lstBank.ListIndex + 1 & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
        End If
    End If
    Exit Sub
Done:
    If Item(GetPlayerBankItemNum(MyIndex, BankNum)).Type = ITEM_TYPE_CURRENCY Then
lblMsg.Caption = "The variable can't handle that amount!"
    End If
End Sub

