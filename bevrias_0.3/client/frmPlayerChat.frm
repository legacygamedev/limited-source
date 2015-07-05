VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmPlayerChat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Player Chat"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4260
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPlayerChat.frx":0000
   ScaleHeight     =   5235
   ScaleWidth      =   4260
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   3255
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   5741
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmPlayerChat.frx":1550
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
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   4560
      Width           =   3615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stop Chatting"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3000
      TabIndex        =   2
      Top             =   4920
      Width           =   990
   End
   Begin VB.Label Label1 
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
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   960
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
        txtChat.SelColor = QBColor(Black)
        txtChat.SelText = s
        txtChat.SelStart = Len(txtChat.Text) - 1
        
        Call SendData("sendchat" & SEP_CHAR & txtSay.Text & SEP_CHAR & END_CHAR)
        txtSay.Text = ""
    End If
End Sub

Private Sub Form_Load()
Dim i As Long
Dim Ending As String
    For i = 1 To 3
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"
 
        If FileExist("GUI\Player Chat" & Ending) Then frmPlayerChat.Picture = LoadPicture(App.Path & "\GUI\Player Chat" & Ending)
    Next i
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
        txtChat.SelColor = QBColor(Black)
        txtChat.SelText = s
        txtChat.SelStart = Len(txtChat.Text) - 1
        
        Call SendData("sendchat" & SEP_CHAR & txtSay.Text & SEP_CHAR & END_CHAR)
        txtSay.Text = ""
    End If
End Sub
