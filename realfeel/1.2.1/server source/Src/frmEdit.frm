VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dual Solace"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10920
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   10920
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstLibrary 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2790
      ItemData        =   "frmEdit.frx":0000
      Left            =   7590
      List            =   "frmEdit.frx":0002
      TabIndex        =   8
      Top             =   3435
      Width           =   3225
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   8415
      TabIndex        =   7
      Top             =   6525
      Width           =   1575
   End
   Begin RichTextLib.RichTextBox txtCommands 
      Height          =   2670
      Left            =   7590
      TabIndex        =   6
      Top             =   570
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   4710
      _Version        =   393217
      Enabled         =   0   'False
      TextRTF         =   $"frmEdit.frx":0004
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox cmbCommands 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmEdit.frx":007F
      Left            =   7590
      List            =   "frmEdit.frx":0197
      TabIndex        =   5
      Text            =   "          Select a Command"
      Top             =   150
      Width           =   3225
   End
   Begin Server.rtbSyntax rtfEdit 
      Height          =   6060
      Left            =   165
      TabIndex        =   4
      Top             =   150
      Width           =   7230
      _ExtentX        =   10636
      _ExtentY        =   6615
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmEdit.frx":06F6
      RightMargin     =   1.00000e5
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1890
      TabIndex        =   1
      Top             =   6510
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   180
      TabIndex        =   0
      Top             =   6510
      Width           =   1575
   End
   Begin VB.Label lblFile 
      Caption         =   "Blank"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5100
      TabIndex        =   3
      Top             =   6615
      Width           =   1845
   End
   Begin VB.Label Label1 
      Caption         =   "Now Editing: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Top             =   6615
      Width           =   1080
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEdit_Click()
    EditType = EDIT_OTHER
    frmEdit.rtfEdit.LoadFile App.Path & "\Library\" & frmEdit.lstLibrary.List(frmEdit.lstLibrary.ListIndex), rtfText
    lblFile.Caption = frmEdit.lstLibrary.List(frmEdit.lstLibrary.ListIndex)
    rtfEdit.HighlightRefresh
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("frmEdit.frm", "cmdEdit_Click", Err.Number, Err.Description)
End Sub

Private Sub cmdQuit_Click()
'On Error GoTo errorhandler:
    Unload Me
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("frmEdit.frm", "cmdQuit_Click", Err.Number, Err.Description)
End Sub

Private Sub cmdSave_Click()
'On Error GoTo errorhandler:
If EditType = EDIT_SERVERMESSAGE Then
    Call PutVar(App.Path & "\Data\data.ini", "Strings", "Msg", rtfEdit.Text)
    MsgBox ("Save Successful!")
    Exit Sub
ElseIf EditType = EDIT_OTHER Then
    rtfEdit.SaveFile App.Path & "\Library\" & lblFile.Caption, rtfText
    MsgBox ("Save Successful!")
    Call LoadScripts
    Exit Sub
Else
    MsgBox ("Error with saving the text document!")
    Exit Sub
End If
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("frmEdit.frm", "cmdSave_Click", Err.Number, Err.Description)
End Sub

' This sub is huge because of the library. I thought about loading it from a text file, but I chose to hardcode it.
Private Sub cmbCommands_Click()
    ' The 4 line break-up is just for ease of me to enter data, feel free to compile it to one line if you want.
    If cmbCommands.Text = "GetPlayerLogin" Then
        txtCommands.Text = "Function: " & vbCrLf
        txtCommands.Text = txtCommands.Text + "Returns Index's login name." & vbCrLf & vbCrLf
        txtCommands.Text = txtCommands.Text + "Usage: " & vbCrLf
        txtCommands.Text = txtCommands.Text + "GetPlayerLogin(Index)"
    ElseIf cmbCommands.Text = "SetPlayerLogin" Then
        txtCommands.Text = "Sub: " & vbCrLf
        txtCommands.Text = txtCommands.Text + "Changes Index's login name to a string." & vbCrLf & vbCrLf
        txtCommands.Text = txtCommands.Text + "Usage: " & vbCrLf
        txtCommands.Text = txtCommands.Text + "Call SetPlayerLogin(Index,""NewLogin"")"
    ElseIf cmbCommands.Text = "GetPlayerPassword" Then
        txtCommands.Text = "Function: " & vbCrLf
        txtCommands.Text = txtCommands.Text + "Returns Index's password as a string." & vbCrLf & vbCrLf
        txtCommands.Text = txtCommands.Text + "Usage: " & vbCrLf
        txtCommands.Text = txtCommands.Text + "GetPlayerPassword(Index)"
    ElseIf cmbCommands.Text = "SetPlayerPassword" Then
        txtCommands.Text = "Sub: " & vbCrLf
        txtCommands.Text = txtCommands.Text + "Changes Index's password to a string." & vbCrLf & vbCrLf
        txtCommands.Text = txtCommands.Text + "Usage: " & vbCrLf
        txtCommands.Text = txtCommands.Text + "Call SetPlayerPassword(Index,""NewPass"")"
    ElseIf cmbCommands.Text = "GetPlayerName" Then
        txtCommands.Text = "Function: " & vbCrLf
        txtCommands.Text = txtCommands.Text + "Returns Index's name as a string." & vbCrLf & vbCrLf
        txtCommands.Text = txtCommands.Text + "Usage: " & vbCrLf
        txtCommands.Text = txtCommands.Text + "GetPlayerName(Index)"
    ElseIf cmbCommands.Text = "SetPlayerName" Then
        txtCommands.Text = "Sub: " & vbCrLf
        txtCommands.Text = txtCommands.Text + "Changes Index's name to a string." & vbCrLf & vbCrLf
        txtCommands.Text = txtCommands.Text + "Usage: " & vbCrLf
        txtCommands.Text = txtCommands.Text + "Call SetPlayerName(Index,""NewName"")"
    
    Else
        txtCommands.Text = "No Data"
    End If
    
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("frmEdit.frm", "cmdSave_Click", Err.Number, Err.Description)
End Sub
