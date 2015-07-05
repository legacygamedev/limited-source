VERSION 5.00
Begin VB.Form frmBank 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bank"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   399
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstBank 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFC0&
      Height          =   4320
      Left            =   3120
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
   Begin VB.ListBox lstInventory 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFC0&
      Height          =   4320
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   240
      TabIndex        =   5
      Top             =   5160
      Width           =   5520
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   4410
      TabIndex        =   4
      Top             =   4800
      Width           =   60
   End
   Begin VB.Label label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   240
      TabIndex        =   3
      Top             =   5520
      Width           =   5535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   1515
      TabIndex        =   2
      Top             =   4800
      Width           =   60
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

