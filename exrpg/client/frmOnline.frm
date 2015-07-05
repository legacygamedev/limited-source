VERSION 5.00
Begin VB.Form frmOnline 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Who's Online"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3750
   Icon            =   "frmOnline.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   3750
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstOnline 
      Appearance      =   0  'Flat
      Height          =   3150
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   3495
   End
End
Attribute VB_Name = "frmOnline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Terminate()
Call SendOnlineList
End Sub
Private Sub Form_Load()

Dim i As Long
Dim Ending As String
    For i = 1 To 3
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"

        If FileExist("\core files\interface\online" & Ending) Then frmOnline.Picture = LoadPicture(App.Path & "\core files\interface\online" & Ending)
    Next i
End Sub
