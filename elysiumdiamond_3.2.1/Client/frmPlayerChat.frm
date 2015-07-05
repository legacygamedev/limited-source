VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmPlayerChat 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Player Chat"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3435
   LinkTopic       =   "Form1"
   Picture         =   "frmPlayerChat.frx":0000
   ScaleHeight     =   3780
   ScaleWidth      =   3435
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   2055
      Left            =   390
      TabIndex        =   3
      Top             =   960
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   3625
      _Version        =   393217
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmPlayerChat.frx":57C6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtSay 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   390
      TabIndex        =   0
      Top             =   3000
      Width           =   2655
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00808080&
      Height          =   315
      Left            =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chatting With: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   2730
   End
End
Attribute VB_Name = "frmPlayerChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Copyright (c) 2007-2008 Elysium Source. All rights reserved.
' This code is licensed under the Elysium General License.

Private Sub Form_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Then
        KeyAscii = 0
        If Trim$(txtSay.Text) = vbNullString Then Exit Sub
        s = vbNewLine & GetPlayerName(MyIndex) & "> " & Trim$(txtSay.Text)
        txtChat.SelStart = Len(txtChat.Text)
        txtChat.SelColor = QBColor(Black)
        txtChat.SelText = s
        txtChat.SelStart = Len(txtChat.Text) - 1
        
        Call SendData(SENDCHAT_CHAR & SEP_CHAR & txtSay.Text & END_CHAR)
        txtSay.Text = vbNullString
    End If
End Sub

Private Sub Form_Load()
Dim I As Long
Dim Ending As String
    For I = 1 To 3
        If I = 1 Then Ending = ".gif"
        If I = 2 Then Ending = ".jpg"
        If I = 3 Then Ending = ".png"
 
        If FileExist("GUI\contentbox" & Ending) Then frmPlayerChat.Picture = LoadPicture(App.Path & "\GUI\contentbox" & Ending)
    Next I
End Sub

Private Sub Label2_Click()
    Call SendData(QCHAT_CHAR & END_CHAR)
End Sub

Private Sub txtChat_GotFocus()
    txtSay.SetFocus
End Sub

Private Sub txtSay_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Then
        KeyAscii = 0
        If Trim$(txtSay.Text) = vbNullString Then Exit Sub
        s = vbNewLine & GetPlayerName(MyIndex) & "> " & Trim$(txtSay.Text)
        txtChat.SelStart = Len(txtChat.Text)
        txtChat.SelColor = QBColor(Black)
        txtChat.SelText = s
        txtChat.SelStart = Len(txtChat.Text) - 1
        
        Call SendData(SENDCHAT_CHAR & SEP_CHAR & txtSay.Text & END_CHAR)
        txtSay.Text = vbNullString
    End If
End Sub
