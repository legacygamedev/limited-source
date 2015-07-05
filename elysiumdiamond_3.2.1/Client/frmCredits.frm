VERSION 5.00
Begin VB.Form frmCredits 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Engine Credits"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   -45
   ClientWidth     =   6000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCredits.frx":0000
   ScaleHeight     =   6000
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCredits 
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
      Left            =   1690
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2060
      Width           =   2625
   End
   Begin VB.Label picCancel 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      Top             =   1200
      Width           =   360
   End
End
Attribute VB_Name = "frmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Copyright (c) 2007-2008 Elysium Source. All rights reserved.
' This code is licensed under the Elysium General License.

Option Explicit

Private Sub Form_Load()
Dim I As Long
Dim Ending As String
Dim f As Long
    For I = 1 To 3
        If I = 1 Then Ending = ".gif"
        If I = 2 Then Ending = ".jpg"
        If I = 3 Then Ending = ".png"

        If FileExist("GUI\mediumlist" & Ending) Then frmChars.Picture = LoadPicture(App.Path & "\GUI\mediumlist" & Ending)
    Next I
    
    txtCredits.Text = "--- Elysium Diamond v3 ---" & vbNewLine & "This version edited by Giaken" & vbNewLine & vbNewLine & "Previous version by pingu (pingu@splamm.com)" & vbNewLine & vbNewLine & "Previous versions made by GodSentDeath and Coke" & vbNewLine & vbNewLine & " Special thanks to l4lucas (for the cool map editor), to Johnman (for the SS converter and hosting), and to nex666 for the awesome GUI!" & vbNewLine & vbNewLine & " Elysium Diamond is covered under the EGL license" & vbNewLine & " If you downloaded this software and the owner claims to have made it and is not an owner of Elysium Source, please visit the forum at www.elysiumsource.net"
    
    f = FreeFile
    If FileExist("credits.txt") Then
        Open App.Path & "\credits.txt" For Input As #f
            txtCredits.Text = txtCredits.Text & vbNewLine & Input$(LOF(f), f)
        Close #f
    End If
End Sub

Private Sub picCancel_Click()
    frmMainMenu.Visible = True
    frmCredits.Visible = False
End Sub
