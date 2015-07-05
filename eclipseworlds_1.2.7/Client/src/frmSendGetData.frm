VERSION 5.00
Begin VB.Form frmLoad 
   BackColor       =   &H00181C21&
   BorderStyle     =   0  'None
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   4200
   Icon            =   "frmSendGetData.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSendGetData.frx":038A
   ScaleHeight     =   32
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   4200
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If KeyAscii = vbKeyEscape Then
        If GameLoaded Then
            frmLoad.Visible = False
            frmMenu.Visible = True
        End If
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Form_KeyPress", "frmLoad", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

' When the form Close button is pressed
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        frmLoad.Hide
        frmMenu.Show
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Form_QueryUnload", "frmLoad", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub
