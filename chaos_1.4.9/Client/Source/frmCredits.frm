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
      Text            =   "frmCredits.frx":10134
      Top             =   2060
      Width           =   2625
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Chaos Engine"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   1080
      TabIndex        =   2
      Top             =   1200
      Width           =   3705
   End
   Begin VB.Label picCancel 
      BackStyle       =   0  'Transparent
      Caption         =   " X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5640
      TabIndex        =   0
      Top             =   120
      Width           =   360
   End
End
Attribute VB_Name = "frmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Copyright (c) 2006 Chaos Engine Source. All rights reserved.
' This code is licensed under the Chaos Engine General License.

Option Explicit

Private Sub Form_Load()
Dim i As Long
Dim Ending As String
Dim f As Long
    For i = 1 To 3
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"

        
    Next i
    
    txtCredits.Text = "--- Chaos Engine Diamond ---" & vbNewLine & "Made by pingu (pingu@splamm.com)" & vbNewLine & vbNewLine & "Previous versions made by GodSentDeath and Coke" & vbNewLine & vbNewLine & " Special thanks to l4lucas (for the cool map editor), and to Johnman (for the SS converter and hosting)" & vbNewLine & vbNewLine & " Chaos Engine Diamond is covered under the EGL license" & vbNewLine & " If you downloaded this software and the owner claims to have made it and he is not pingu or an owner of Chaos Engine Source, please contact pingu (pingu@splamm.com)-Fixes Added by Frozengod-"
    
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

