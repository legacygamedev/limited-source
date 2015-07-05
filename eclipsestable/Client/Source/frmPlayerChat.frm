VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmPlayerChat 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Player Chat"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5985
   ControlBox      =   0   'False
   Icon            =   "frmPlayerChat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   399
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   3165
      Left            =   270
      TabIndex        =   3
      Top             =   1695
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   5583
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmPlayerChat.frx":0FC2
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
      Left            =   240
      TabIndex        =   0
      Top             =   4920
      Width           =   5475
   End
   Begin VB.Label Label2 
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
      Height          =   390
      Left            =   1800
      TabIndex        =   2
      Top             =   5400
      Width           =   2370
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Chatting With: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   270
      TabIndex        =   1
      Top             =   1200
      Width           =   5415
   End
End
Attribute VB_Name = "frmPlayerChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim s As String

    If (KeyAscii = vbKeyReturn) Then
        KeyAscii = 0
        If Trim$(txtSay.Text) = vbNullString Then
            Exit Sub
        End If
        s = vbNewLine & GetPlayerName(MyIndex) & "> " & Trim$(txtSay.Text)
        txtChat.SelStart = Len(txtChat.Text)
        txtChat.SelColor = QBColor(BLACK)
        txtChat.SelText = s
        txtChat.SelStart = Len(txtChat.Text) - 1

        Call SendData("sendchat" & SEP_CHAR & txtSay.Text & END_CHAR)
        txtSay.Text = vbNullString
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

        If FileExists("GUI\Player Chat" & Ending) Then
            frmPlayerChat.Picture = LoadPicture(App.Path & "\GUI\Player Chat" & Ending)
        End If
    Next I
End Sub

Private Sub Label2_Click()
    Call SendData("qchat" & END_CHAR)
End Sub

Private Sub txtChat_GotFocus()
    On Error Resume Next
    txtSay.SetFocus
End Sub

Private Sub txtSay_KeyPress(KeyAscii As Integer)
    Dim s As String

    If (KeyAscii = vbKeyReturn) Then
        KeyAscii = 0
        If Trim$(txtSay.Text) = vbNullString Then
            Exit Sub
        End If
        s = vbNewLine & GetPlayerName(MyIndex) & "> " & Trim$(txtSay.Text)
        txtChat.SelStart = Len(txtChat.Text)
        txtChat.SelColor = QBColor(BLACK)
        txtChat.SelText = s
        txtChat.SelStart = Len(txtChat.Text) - 1

        Call SendData("sendchat" & SEP_CHAR & txtSay.Text & END_CHAR)
        txtSay.Text = vbNullString
    End If
End Sub
