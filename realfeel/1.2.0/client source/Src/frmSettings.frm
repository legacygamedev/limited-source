VERSION 5.00
Begin VB.Form frmSettings 
   BackColor       =   &H00004080&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3120
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   2970
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSettings.frx":0000
   ScaleHeight     =   3120
   ScaleWidth      =   2970
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picCancel 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1490
      Picture         =   "frmSettings.frx":2A90E
      ScaleHeight     =   375
      ScaleWidth      =   1500
      TabIndex        =   7
      Top             =   2760
      Width           =   1500
   End
   Begin VB.PictureBox picSubmit 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -10
      Picture         =   "frmSettings.frx":2C69C
      ScaleHeight     =   375
      ScaleWidth      =   1500
      TabIndex        =   6
      Top             =   2760
      Width           =   1500
   End
   Begin VB.PictureBox picHeading 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   -10
      Picture         =   "frmSettings.frx":2E42A
      ScaleHeight     =   780
      ScaleWidth      =   3000
      TabIndex        =   5
      Top             =   -10
      Width           =   3000
   End
   Begin VB.Frame fraAccount 
      BackColor       =   &H00F5763F&
      Caption         =   "Load Account"
      Height          =   975
      Left            =   960
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
      Begin VB.OptionButton optOn 
         BackColor       =   &H0009E7F2&
         Caption         =   "On"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   615
      End
      Begin VB.OptionButton optOff 
         BackColor       =   &H0009E7F2&
         Caption         =   "Off"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Text            =   "Port"
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Text            =   "IP"
      Top             =   960
      Width           =   1455
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PicArray() As VB.PictureBox

Public Sub MakePic(ByVal n As Long)
    Set PicArray(n) = Controls.Add("VB.PictureBox", "PicArray" & CStr(n), Me)
    PicArray(n).Appearance = 0
    PicArray(n).BorderStyle = 0
    PicArray(n).AutoRedraw = True
    PicArray(n).AutoSize = True
    PicArray(n).Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Menu Settings", "PicFrame" & n & "Target"))
    PicArray(n).Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Menu Settings", "PicFrame" & n & "X"))
    PicArray(n).top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Menu Settings", "PicFrame" & n & "Y"))
    PicArray(n).Visible = True
End Sub

Public Sub SetArray(ByVal num As Long)
    ReDim PicArray(1 To num) As VB.PictureBox
End Sub

Private Sub picCancel_Click()
    Call LoadMenu
    frmSettings.Visible = False
End Sub

Private Sub picSubmit_Click()
'Check the text boxes
If txtIP.Text = "" Then
    Call MsgBox("Please set your IP!")
    Exit Sub
ElseIf txtPort.Text = "" Then
    Call MsgBox("Please set your Port!")
    Exit Sub
End If

'Set the IP and Port
Call PutVar(App.Path & "\data.dat", "Address", "IP", txtIP.Text)
Call PutVar(App.Path & "\data.dat", "Address", "Port", txtPort.Text)

GAME_IP = txtIP.Text
GAME_PORT = CInt(Trim$(txtPort.Text))
frmDualSolace.Socket.Close
frmDualSolace.Socket.RemoteHost = GAME_IP
frmDualSolace.Socket.RemotePort = GAME_PORT

'Check the account data and follow accordingly
If optOn.Value = True Then
    Call PutVar(App.Path & "\data.dat", "Account", "Enable", "1")
ElseIf optOff.Value = True Then
    Call PutVar(App.Path & "\data.dat", "Account", "Enable", "0")
Else
    Call MsgBox("Problem recording the account settings!")
    Exit Sub
End If

Call LoadMenu
frmSettings.Visible = False

End Sub

Private Sub Form_Load()
'Check data to prevent error
If GetVar(App.Path & "\data.dat", "Address", "IP") = "" Then
    Call PutVar(App.Path & "\data.dat", "Address", "IP", CStr(frmDualSolace.Socket.LocalIP))
    txtIP.Text = frmDualSolace.Socket.LocalIP
Else
    txtIP.Text = GetVar(App.Path & "\data.dat", "Address", "IP")
End If
If GetVar(App.Path & "\data.dat", "Address", "Port") = "" Then
    Call PutVar(App.Path & "\data.dat", "Address", "Port", frmDualSolace.Socket.LocalPort)
    txtPort.Text = frmDualSolace.Socket.LocalPort
Else
    txtPort.Text = GetVar(App.Path & "\data.dat", "Address", "Port")
End If

If GetVar(App.Path & "\data.dat", "Account", "Enable") = "1" Then
    'Set option to on
    optOn.Value = True
Else
    'Set option to off
    optOn.Value = False
End If
End Sub
