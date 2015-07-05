VERSION 5.00
Begin VB.Form frmSendGetData 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loading"
   ClientHeight    =   675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4500
   ControlBox      =   0   'False
   Icon            =   "frmSendGetData.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   675
   ScaleWidth      =   4500
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading.."
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   225
      Width           =   4260
   End
End
Attribute VB_Name = "frmSendGetData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyEscape) Then
        Call GameDestroy
        End
    End If
End Sub

Private Sub form_load()

Dim i As Long
Dim Ending As String
    For i = 1 To 3
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"

        If FileExist("\core files\interface\sendgetdata" & Ending) Then frmSendGetData.Picture = LoadPicture(App.Path & "\core files\interface\sendgetdata" & Ending)
    Next i

End Sub
