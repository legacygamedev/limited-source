VERSION 5.00
Begin VB.Form frmKeepNotes 
   BorderStyle     =   0  'None
   Caption         =   "Player Notes"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   0
   ClientWidth     =   3435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmKeepNotes.frx":0000
   ScaleHeight     =   3780
   ScaleWidth      =   3435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Save 
      Appearance      =   0  'Flat
      BackColor       =   &H00789298&
      Caption         =   "Save Notes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   960
      MaskColor       =   &H00789298&
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox Notetext 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   400
      TabIndex        =   0
      Top             =   980
      Width           =   2625
   End
   Begin VB.Label Close 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmKeepNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Copyright (c) 2006 Elysium Source. All rights reserved.
' This code is licensed under the Elysium General License.

'Private Sub Close_Click()
'frmKeepNotes.Visible = False
'End Sub

Private Sub Form_Load()
Dim I As Long
Dim Ending As String
    For I = 1 To 3
        If I = 1 Then Ending = ".gif"
        If I = 2 Then Ending = ".jpg"
        If I = 3 Then Ending = ".png"
 
        If FileExist("GUI\contentlist" & Ending) Then frmKeepNotes.Picture = LoadPicture(App.Path & "\GUI\contentlist" & Ending)
    Next I
End Sub

Private Sub Save_Click()
Dim iFileNum As Integer

'Get a free file handle
iFileNum = FreeFile

'If the file is not there, one will be created
'If the file does exist, this one will
'overwrite it.
Open App.Path & "\notes.txt" For Output As iFileNum

Print #iFileNum, Notetext.Text

Close iFileNum

End Sub

