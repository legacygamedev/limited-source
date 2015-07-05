VERSION 5.00
Begin VB.Form frmSendGetData 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   5955
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSendGetData.frx":0000
   ScaleHeight     =   400
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   397
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtStatus 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2175
      Left            =   840
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmSendGetData.frx":10134
      Top             =   1920
      Width           =   4455
   End
   Begin VB.Timer fancy 
      Interval        =   100
      Left            =   5520
      Top             =   120
   End
End
Attribute VB_Name = "frmSendGetData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Copyright (c) 2006 Chaos Engine Source. All rights reserved.
' This code is licensed under the Chaos Engine General License.

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyEscape) Then
        Call DestroyDirectX
        Call StopMidi
        InGame = False
        frmMirage.Socket.Close
        frmMainMenu.Visible = True
        Connucted = False
        Unload Me
    End If
End Sub

Private Sub Form_Load()
Dim i As Long
Dim Ending As String
    For i = 1 To 3
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"
 
        
    Next i
End Sub

Private Sub Label1_Click()
    Call DestroyDirectX
    Call StopMidi
    InGame = False
    frmMirage.Socket.Close
    frmMainMenu.Visible = True
    Connucted = False
    Unload Me
End Sub

Private Sub lblStatus_Click()

End Sub
