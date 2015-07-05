VERSION 5.00
Begin VB.Form frmSpells 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spells"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3750
   Icon            =   "frmSpells.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   3750
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstSpells 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   2565
      ItemData        =   "frmSpells.frx":2372
      Left            =   150
      List            =   "frmSpells.frx":2374
      TabIndex        =   0
      Top             =   1065
      Width           =   3450
   End
   Begin VB.Label lblCast 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cast"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   915
      TabIndex        =   1
      Top             =   3795
      Width           =   1935
   End
End
Attribute VB_Name = "frmSpells"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim i As Long
Dim Ending As String
    For i = 1 To 3
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"

        If FileExist("\core files\interface\spells" & Ending) Then frmSpells.Picture = LoadPicture(App.Path & "\core files\interface\spells" & Ending)
    Next i
End Sub

Private Sub lblCast_Click()
    If Player(MyIndex).Spell(lstSpells.ListIndex + 1) > 0 Then
        If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
            If Player(MyIndex).Moving = 0 Then
                Call SendData("cast" & SEP_CHAR & lstSpells.ListIndex + 1 & SEP_CHAR & END_CHAR)
                Player(MyIndex).Attacking = 1
                Player(MyIndex).AttackTimer = GetTickCount
                Player(MyIndex).CastedSpell = YES
            Else
                Call AddText("Cannot cast while walking!", BrightRed)
            End If
        End If
    Else
        Call AddText("No spell here.", BrightRed)
    End If
End Sub
