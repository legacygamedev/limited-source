VERSION 5.00
Begin VB.Form frmBank 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bank"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5985
   Icon            =   "frmBank.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   399
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   1
      Left            =   7320
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   9
      Top             =   2400
      Width           =   540
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   1
         Left            =   15
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   10
         Top             =   15
         Width           =   480
         Begin VB.PictureBox PicBank 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   1
            Left            =   15
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   128
            TabIndex        =   11
            Top             =   15
            Width           =   1920
         End
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   0
      Left            =   7320
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   6
      Top             =   1680
      Width           =   540
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   0
         Left            =   15
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   7
         Top             =   15
         Width           =   480
         Begin VB.PictureBox PicBank 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   0
            Left            =   15
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   128
            TabIndex        =   8
            Top             =   15
            Width           =   1920
         End
      End
   End
   Begin VB.ListBox lstBank 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFC0&
      Height          =   3540
      Left            =   3120
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin VB.ListBox lstInventory 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFC0&
      Height          =   3540
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2535
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
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   240
      TabIndex        =   5
      Top             =   4800
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
      Height          =   330
      Left            =   3840
      TabIndex        =   4
      Top             =   4200
      Width           =   915
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
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   5280
      Width           =   615
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
      Height          =   330
      Left            =   1200
      TabIndex        =   2
      Top             =   4200
      Width           =   780
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
            GoldAmount = InputBox("How much " & Trim$(Item(GetPlayerInvItemNum(MyIndex, InvNum)).name) & "(" & GetPlayerInvItemValue(MyIndex, InvNum) & ") would you like to deposit?", "Deposit " & Trim$(Item(GetPlayerInvItemNum(MyIndex, InvNum)).name), 0, frmBank.Left, frmBank.Top)
            If IsNumeric(GoldAmount) Then
                Call SendData("bankdeposit" & SEP_CHAR & lstInventory.ListIndex + 1 & SEP_CHAR & GoldAmount & END_CHAR)
            End If
        Else
            Call SendData("bankdeposit" & SEP_CHAR & lstInventory.ListIndex + 1 & SEP_CHAR & 0 & END_CHAR)
        End If
    End If
    Exit Sub
Done:
    If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_CURRENCY Then
    ' MsgBox "The variable cant handle that amount!"
    End If
End Sub

Sub InvItems()
    Dim BankNum As Long
    Dim GoldAmount As String
    On Error GoTo Done

    BankNum = lstBank.ListIndex + 1
    If GetPlayerBankItemNum(MyIndex, BankNum) > 0 And GetPlayerBankItemNum(MyIndex, BankNum) <= MAX_ITEMS Then
        If Item(GetPlayerBankItemNum(MyIndex, BankNum)).Type = ITEM_TYPE_CURRENCY Then
            GoldAmount = InputBox("How much " & Trim$(Item(GetPlayerBankItemNum(MyIndex, BankNum)).name) & "(" & GetPlayerBankItemValue(MyIndex, BankNum) & ") would you like to deposit?", "Deposit " & Trim$(Item(GetPlayerBankItemNum(MyIndex, BankNum)).name), 0, frmBank.Left, frmBank.Top)
            If IsNumeric(GoldAmount) Then
                Call SendData("bankwithdraw" & SEP_CHAR & lstBank.ListIndex + 1 & SEP_CHAR & GoldAmount & END_CHAR)
            End If
        Else
            Call SendData("bankwithdraw" & SEP_CHAR & lstBank.ListIndex + 1 & SEP_CHAR & 0 & END_CHAR)
        End If
    End If
    Exit Sub
Done:
    If Item(GetPlayerBankItemNum(MyIndex, BankNum)).Type = ITEM_TYPE_CURRENCY Then
        MsgBox "The variable cant handle that amount!"
    End If
End Sub

Private Sub Form_Load()
    Dim I As Long
    Dim Ending As String

    For I = 1 To 3
        If I = 1 Then
            Ending = ".gif"
        End If
        If I = 2 Then
            Ending = ".jpg"
        End If
        If I = 3 Then
            Ending = ".png"
        End If

        If FileExists("GUI\Bank" & Ending) Then
            frmBank.Picture = LoadPicture(App.Path & "\GUI\Bank" & Ending)
        End If
    Next I
End Sub

