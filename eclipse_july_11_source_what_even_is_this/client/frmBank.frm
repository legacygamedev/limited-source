VERSION 5.00
Begin VB.Form frmBank 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "Bank"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   ScaleHeight     =   400
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   457
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstBank 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFC0&
      Height          =   4710
      Left            =   3480
      TabIndex        =   1
      Top             =   480
      Width           =   3015
   End
   Begin VB.ListBox lstInventory 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFC0&
      Height          =   4710
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   3015
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
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   600
      TabIndex        =   6
      Top             =   5400
      Width           =   5640
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
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Left            =   0
      TabIndex        =   5
      Top             =   120
      Width           =   6840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Withdraw"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   210
      Left            =   3480
      TabIndex        =   4
      Top             =   5160
      Width           =   780
   End
   Begin VB.Label label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   180
      Left            =   5670
      TabIndex        =   3
      Top             =   5610
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Deposit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   210
      Left            =   360
      TabIndex        =   2
      Top             =   5160
      Width           =   630
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
            GoldAmount = InputBox("How much " & Trim$(Item(GetPlayerInvItemNum(MyIndex, InvNum)).Name) & "(" & GetPlayerInvItemValue(MyIndex, InvNum) & ") would you like to deposit?", "Deposit " & Trim$(Item(GetPlayerInvItemNum(MyIndex, InvNum)).Name), 0, frmBank.Left, frmBank.Top)
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
        'MsgBox "The variable cant handle that amount!"
    End If
End Sub

Sub InvItems()
Dim BankNum As Long
Dim GoldAmount As String
On Error GoTo Done

    BankNum = lstBank.ListIndex + 1
    If GetPlayerBankItemNum(MyIndex, BankNum) > 0 And GetPlayerBankItemNum(MyIndex, BankNum) <= MAX_ITEMS Then
        If Item(GetPlayerBankItemNum(MyIndex, BankNum)).Type = ITEM_TYPE_CURRENCY Then
            GoldAmount = InputBox("How much " & Trim$(Item(GetPlayerBankItemNum(MyIndex, BankNum)).Name) & "(" & GetPlayerBankItemValue(MyIndex, BankNum) & ") would you like to deposit?", "Deposit " & Trim$(Item(GetPlayerBankItemNum(MyIndex, BankNum)).Name), 0, frmBank.Left, frmBank.Top)
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
        MsgBox "The variable cant handle that amount!"
    End If
End Sub

Private Sub Form_Load()
Dim i As Long
Dim Ending As String
    For i = 1 To 3
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"

        If FileExist("GUI\Bank" & Ending) Then frmBank.Picture = LoadPicture(App.Path & "\GUI\Bank" & Ending)
    Next i
End Sub
    
