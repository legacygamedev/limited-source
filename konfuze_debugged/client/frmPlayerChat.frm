VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmPlayerChat 
   BorderStyle     =   0  'None
   Caption         =   "Player Chat"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   ControlBox      =   0   'False
   Icon            =   "frmPlayerChat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPlayerChat.frx":0FC2
   ScaleHeight     =   6000
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   3495
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   6165
      _Version        =   393217
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmPlayerChat.frx":4991
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
      Top             =   5040
      Width           =   5175
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
      Left            =   4560
      TabIndex        =   2
      Top             =   5400
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
      Top             =   1200
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
