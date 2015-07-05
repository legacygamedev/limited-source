VERSION 5.00
Begin VB.Form frmCredits 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credits"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   ControlBox      =   0   'False
   Icon            =   "frmCredits.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   3750
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Free game engine at www.ramzaengine.com.ar"
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   3405
      Width           =   3525
   End
   Begin VB.Label lblCredits 
      BackStyle       =   0  'Transparent
      Caption         =   "User Credits Here"
      Height          =   1785
      Left            =   240
      TabIndex        =   1
      Top             =   1350
      Width           =   3285
   End
   Begin VB.Label picCancel 
      BackStyle       =   0  'Transparent
      Caption         =   "back to menu"
      Height          =   225
      Left            =   1335
      TabIndex        =   0
      Top             =   3690
      Width           =   1035
   End
End
Attribute VB_Name = "frmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Terminate()
    Call GameDestroy
End Sub

Private Sub picCancel_Click()
    frmMainMenu.Visible = True
    frmCredits.Visible = False
End Sub

Private Sub Form_Load()
lblCredits.Caption = ReadIniValue(App.Path & "\Core Files\Configuration.ini", "general", "credits")

Dim i As Long
Dim Ending As String
    For i = 1 To 3
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"

        If FileExist("\core files\interface\credits" & Ending) Then frmCredits.Picture = LoadPicture(App.Path & "\core files\interface\credits" & Ending)
    Next i
'read/write ini stuff
'WriteIniValue App.Path & "\MyTest.ini", "Default", "Text1", Text1.Text
End Sub

