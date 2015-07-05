VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmPlayerChat 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "Player Chat"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6855
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "frmPlayerChat.frx":0000
   ScaleHeight     =   400
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   457
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   4695
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   8281
      _Version        =   393217
      BackColor       =   0
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmPlayerChat.frx":86002
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtSay 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   5160
      Width           =   6375
   End
   Begin VB.Label label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "End Chat"
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
      Left            =   5640
      TabIndex        =   3
      Top             =   5610
      Width           =   1095
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chatting With: "
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
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmPlayerChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Then
        KeyAscii = 0
        If Trim(txtSay.Text) = "" Then Exit Sub
        s = vbNewLine & GetPlayerName(MyIndex) & "> " & Trim(txtSay.Text)
        txtChat.SelStart = Len(txtChat.Text)
        txtChat.SelColor = QBColor(Grey)
        txtChat.SelText = s
        txtChat.SelStart = Len(txtChat.Text) - 1
        
        Call SendData("sendchat" & SEP_CHAR & txtSay.Text & SEP_CHAR & END_CHAR)
        txtSay.Text = ""
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    SOffsetX = x
    SOffsetY = y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MoveForm(frmPlayerChat, Button, Shift, x, y)
End Sub

Private Sub Label2_Click()
    Call SendData("qchat" & SEP_CHAR & END_CHAR)
End Sub

Private Sub txtChat_GotFocus()
    txtSay.SetFocus
End Sub

Private Sub txtSay_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Then
        KeyAscii = 0
        If Trim(txtSay.Text) = "" Then Exit Sub
        s = vbNewLine & GetPlayerName(MyIndex) & "> " & Trim(txtSay.Text)
        txtChat.SelStart = Len(txtChat.Text)
        txtChat.SelColor = QBColor(Grey)
        txtChat.SelText = s
        txtChat.SelStart = Len(txtChat.Text) - 1
        
        Call SendData("sendchat" & SEP_CHAR & txtSay.Text & SEP_CHAR & END_CHAR)
        txtSay.Text = ""
    End If
End Sub
