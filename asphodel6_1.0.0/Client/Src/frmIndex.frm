VERSION 5.00
Begin VB.Form frmIndex 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Index"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   5055
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3480
      Width           =   1575
   End
   Begin VB.ListBox lstIndex 
      Appearance      =   0  'Flat
      Height          =   2670
      ItemData        =   "frmIndex.frx":0000
      Left            =   120
      List            =   "frmIndex.frx":0002
      TabIndex        =   0
      Top             =   720
      Width           =   5055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "search"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Left            =   200
      TabIndex        =   5
      Top             =   95
      Width           =   345
   End
End
Attribute VB_Name = "frmIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ------------------------------------------
' --               Asphodel               --
' ------------------------------------------
Private Sub cmddelete_Click()

    If MsgBox("Are you sure you wish to go through with deleting?", vbYesNo + vbCritical, Game_Name) = vbNo Then Exit Sub
    
    EditorIndex = lstIndex.ListIndex + 1
    SendData CDelete & SEP_CHAR & Editor & SEP_CHAR & EditorIndex & END_CHAR
    
End Sub

Private Sub cmdOk_Click()

    EditorIndex = lstIndex.ListIndex + 1
    
    Select Case Editor
    
        Case GameEditor.Item_
            SendData CEditItem & SEP_CHAR & EditorIndex & END_CHAR
        Case GameEditor.NPC_
            SendData CEditNpc & SEP_CHAR & EditorIndex & END_CHAR
        Case GameEditor.Shop_
            SendData CEditShop & SEP_CHAR & EditorIndex & END_CHAR
        Case GameEditor.Spell_
            SendData CEditSpell & SEP_CHAR & EditorIndex & END_CHAR
        Case GameEditor.Sign_
            SendData CEditSign & SEP_CHAR & EditorIndex & END_CHAR
        Case GameEditor.Anim_
            SendData CEditAnim & SEP_CHAR & EditorIndex & END_CHAR
    End Select
    
    Unload frmIndex
    
End Sub

Private Sub cmdCancel_Click()
    Editor = 0
    Unload frmIndex
End Sub

Private Sub Form_Load()
On Error Resume Next

    txtSearch.SetFocus
    
End Sub

Private Sub txtSearch_Change()
Dim LoopI As Long

    If LenB(Trim$(txtSearch.Text)) < 1 Then
        lstIndex.ListIndex = 0
        lstIndex.Selected(0) = True
        Exit Sub
    End If
    
    For LoopI = 0 To lstIndex.ListCount
        If InStr(1, lstIndex.List(LoopI), Trim$(txtSearch.Text), vbTextCompare) Then
            If LenB(lstIndex.List(LoopI)) > 0 Then
                lstIndex.ListIndex = LoopI
                lstIndex.Selected(LoopI) = True
                Exit For
            End If
        End If
    Next
    
End Sub
