VERSION 5.00
Begin VB.Form frmKeepNotes 
   BorderStyle     =   0  'None
   Caption         =   "Player Notes"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   0
   ClientWidth     =   6000
   Icon            =   "frmKeepNotes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmKeepNotes.frx":0FC2
   ScaleHeight     =   6000
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Close 
      Appearance      =   0  'Flat
      Caption         =   "Close Notes"
      Height          =   300
      Left            =   3600
      TabIndex        =   2
      Top             =   5400
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
      Top             =   5400
      Width           =   2070
   End
   Begin VB.TextBox Notetext 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1320
      Width           =   5400
   End
End
Attribute VB_Name = "frmKeepNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Close_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
Dim i As Long
Dim Ending As String

    For i = 1 To 3
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"
 
        If FileExist("GUI\Notes" & Ending) Then frmKeepNotes.Picture = LoadPicture(App.Path & "\GUI\Notes" & Ending)
    Next i
End Sub

Private Sub Save_Click()
    Call SendData("NOTES" & SEP_CHAR & Notetext.Text & SEP_CHAR & END_CHAR)
    Me.Hide
End Sub

