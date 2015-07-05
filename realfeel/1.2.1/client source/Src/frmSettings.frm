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
      Height          =   315
      Left            =   1680
      Picture         =   "frmSettings.frx":1E484
      ScaleHeight     =   315
      ScaleWidth      =   1290
      TabIndex        =   7
      Top             =   2805
      Width           =   1290
   End
   Begin VB.PictureBox picSubmit 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   0
      Picture         =   "frmSettings.frx":1FA1A
      ScaleHeight     =   315
      ScaleWidth      =   1290
      TabIndex        =   6
      Top             =   2805
      Width           =   1290
   End
   Begin VB.PictureBox picHeading 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1230
      Left            =   0
      Picture         =   "frmSettings.frx":20FB0
      ScaleHeight     =   1230
      ScaleWidth      =   2970
      TabIndex        =   5
      Top             =   0
      Width           =   2970
   End
   Begin VB.Frame fraAccount 
      BackColor       =   &H002F3336&
      Caption         =   "Load Account"
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   1545
      TabIndex        =   2
      Top             =   1335
      Width           =   1230
      Begin VB.OptionButton optOn 
         BackColor       =   &H002F3336&
         Caption         =   "On"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   285
         MaskColor       =   &H002F3336&
         TabIndex        =   4
         Top             =   600
         Width           =   615
      End
      Begin VB.OptionButton optOff 
         BackColor       =   &H002F3336&
         Caption         =   "Off"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   285
         MaskColor       =   &H002F3336&
         TabIndex        =   3
         Top             =   270
         Width           =   615
      End
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Port"
      Top             =   1875
      Width           =   1260
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "IP"
      Top             =   1470
      Width           =   1260
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
