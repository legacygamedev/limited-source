VERSION 5.00
Begin VB.Form frmIndex 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Index"
   ClientHeight    =   4215
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
   ScaleHeight     =   281
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   353
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmddelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3720
      Width           =   1575
   End
   Begin VB.ListBox lstIndex 
      Height          =   3420
      ItemData        =   "frmIndex.frx":0000
      Left            =   120
      List            =   "frmIndex.frx":0002
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ******************************************
' **               rootSource               **
' ******************************************
Private Sub cmddelete_Click()
Dim Buffer As clsBuffer

    If MsgBox("Are you sure you wish to delete this data?", vbYesNo) = vbYes Then
        EditorIndex = lstIndex.ListIndex + 1
        Set Buffer = New clsBuffer
        Buffer.PreAllocate 7
        Buffer.WriteInteger CDelete
        Buffer.WriteByte Editor
        Buffer.WriteLong EditorIndex
        Call SendData(Buffer.ToArray())
    End If
End Sub

Private Sub cmdOk_Click()
Dim Buffer As clsBuffer

    EditorIndex = lstIndex.ListIndex + 1
    
    Set Buffer = New clsBuffer
    Buffer.PreAllocate 6
    Select Case Editor
        Case EDITOR_ITEM
            Buffer.WriteInteger CEditItem
        Case EDITOR_NPC
            Buffer.WriteInteger CEditNpc
        Case EDITOR_SHOP
            Buffer.WriteInteger CEditShop
        Case EDITOR_SPELL
            Buffer.WriteInteger CEditSpell
    End Select
    Buffer.WriteLong EditorIndex
    Call SendData(Buffer.ToArray())
    
    frmIndex.Hide
End Sub

Private Sub cmdCancel_Click()
    Editor = EDITOR_NONE
    Unload frmIndex
End Sub
