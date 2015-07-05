VERSION 5.00
Begin VB.Form frmCharacterSize 
   Caption         =   "Character Size"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3135
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   3135
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Help"
      Height          =   1575
      Left            =   0
      TabIndex        =   6
      Top             =   1560
      Width           =   2895
      Visible         =   0   'False
      Begin VB.CommandButton Command3 
         Caption         =   "Close"
         Height          =   255
         Left            =   2040
         TabIndex        =   13
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "This is measured in pixels, the default is 32*32 pixels on a character. This is only for players, not NPC's."
         Height          =   1215
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sizes"
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.OptionButton Option7 
         Caption         =   "96*96*"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1680
         Width           =   855
      End
      Begin VB.OptionButton Option6 
         Caption         =   "64*96*"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1440
         Width           =   855
      End
      Begin VB.OptionButton Option5 
         Caption         =   "64*64*"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Width           =   855
      End
      Begin VB.OptionButton Option4 
         Caption         =   "32*96*"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton Option3 
         Caption         =   "32*32 Paperdoll"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Help"
         Height          =   255
         Left            =   2040
         TabIndex        =   5
         Top             =   2640
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         Caption         =   "32*64"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "32*32"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Close"
         Height          =   255
         Left            =   1320
         TabIndex        =   1
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Current Size: None"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   2040
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmCharacterSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Frame2.Visible = True
End Sub

Private Sub Command2_Click()
frmCharacterSize.Visible = False
End Sub

Private Sub Command3_Click()
Frame2.Visible = False
End Sub

Private Sub Form_Load()
Dim Packet As String

If GetVar(App.Path & "\Data.ini", "ADDED", "CharSize") = 0 Then
Option1.Value = True
Label1.Caption = "Current Size: 32*32"
Packet = ""
Packet = "sizeofsprites" & SEP_CHAR & "0" & SEP_CHAR & END_CHAR
Call SendDataToAll(Packet)
End If

If GetVar(App.Path & "\Data.ini", "ADDED", "CharSize") = 1 Then
Option2.Value = True
Label1.Caption = "Current Size: 32*64"
Packet = ""
Packet = "sizeofsprites" & SEP_CHAR & "1" & SEP_CHAR & END_CHAR
Call SendDataToAll(Packet)
End If

If GetVar(App.Path & "\Data.ini", "ADDED", "CharSize") = 2 Then
Option3.Value = True
Label1.Caption = "Current Size: 32*32 Paperdoll"
Packet = ""
Packet = "sizeofsprites" & SEP_CHAR & "2" & SEP_CHAR & END_CHAR
Call SendDataToAll(Packet)
End If

End Sub

Private Sub Option1_Click()
Size1 = "0"
Dim Packet As String
Packet = ""
Packet = "sizeofsprites" & SEP_CHAR & "0" & SEP_CHAR & END_CHAR
Call SendDataToAll(Packet)
Label1.Caption = "Current Size: 32*32"
Call PutVar(App.Path & "\Data.ini", "ADDED", "CharSize", "0")
End Sub

Private Sub Option2_Click()
Size1 = "1"
Dim Packet As String
Packet = ""
Packet = "sizeofsprites" & SEP_CHAR & "1" & SEP_CHAR & END_CHAR
Call SendDataToAll(Packet)
Label1.Caption = "Current Size: 32*64"
Call PutVar(App.Path & "\Data.ini", "ADDED", "CharSize", "1")
End Sub

Private Sub Option3_Click()
Size1 = "2"
Dim Packet As String
Packet = ""
Packet = "sizeofsprites" & SEP_CHAR & "2" & SEP_CHAR & END_CHAR
Call SendDataToAll(Packet)
Label1.Caption = "Current Size: 32*32 Paperdoll"
Call PutVar(App.Path & "\Data.ini", "ADDED", "CharSize", "2")
End Sub

Private Sub Option4_Click()
Size1 = "0"
Dim Packet As String
Packet = ""
Packet = "sizeofsprites" & SEP_CHAR & "0" & SEP_CHAR & END_CHAR
Call SendDataToAll(Packet)
Label1.Caption = "Current Size: 32*32"
Call PutVar(App.Path & "\Data.ini", "ADDED", "CharSize", "0")
End Sub

Private Sub Option5_Click()
Size1 = "0"
Dim Packet As String
Packet = ""
Packet = "sizeofsprites" & SEP_CHAR & "0" & SEP_CHAR & END_CHAR
Call SendDataToAll(Packet)
Label1.Caption = "Current Size: 32*32"
Call PutVar(App.Path & "\Data.ini", "ADDED", "CharSize", "0")
End Sub

Private Sub Option6_Click()
Size1 = "0"
Dim Packet As String
Packet = ""
Packet = "sizeofsprites" & SEP_CHAR & "0" & SEP_CHAR & END_CHAR
Call SendDataToAll(Packet)
Label1.Caption = "Current Size: 32*32"
Call PutVar(App.Path & "\Data.ini", "ADDED", "CharSize", "0")
End Sub

Private Sub Option7_Click()
Size1 = "0"
Dim Packet As String
Packet = ""
Packet = "sizeofsprites" & SEP_CHAR & "0" & SEP_CHAR & END_CHAR
Call SendDataToAll(Packet)
Label1.Caption = "Current Size: 32*32"
Call PutVar(App.Path & "\Data.ini", "ADDED", "CharSize", "0")
End Sub
