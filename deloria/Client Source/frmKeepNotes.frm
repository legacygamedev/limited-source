VERSION 5.00
Begin VB.Form frmKeepNotes 
   Caption         =   "Player Notes"
   ClientHeight    =   5970
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmKeepNotes.frx":0000
   ScaleHeight     =   5970
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Close 
      Appearance      =   0  'Flat
      Caption         =   "Close Notes"
      Height          =   300
      Left            =   3600
      TabIndex        =   2
      Top             =   5130
      Width           =   2070
   End
   Begin VB.CommandButton Save 
      Appearance      =   0  'Flat
      BackColor       =   &H00789298&
      Caption         =   "Save Notes"
      Height          =   300
      Left            =   285
      MaskColor       =   &H00789298&
      TabIndex        =   1
      Top             =   5130
      Width           =   2070
   End
   Begin VB.TextBox Notetext 
      Appearance      =   0  'Flat
      Height          =   3870
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   5400
   End
End
Attribute VB_Name = "frmKeepNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Close_Click()
frmKeepNotes.Visible = False
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

