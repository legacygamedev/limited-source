VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dual Solace"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   6270
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox rtfEdit 
      Height          =   3855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   6800
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmEdit.frx":0000
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   4080
      Width           =   1575
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
    Call PutVar(App.Path & "\Data\data.ini", "Info", "Msg", rtfEdit.Text)
    MsgBox ("Save Successful!")
    Exit Sub
ElseIf EditType = EDIT_SCRIPT Then
    Kill (App.Path & "\scripts\main.txt")
    rtfEdit.SaveFile App.Path & "\scripts\main.txt", rtfText
    MsgBox ("Save Successful!")
    Exit Sub
ElseIf EditType = EDIT_OTHER Then
    rtfEdit.SaveFile lblFile.Caption, rtfText
    MsgBox ("Save Successful!")
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
