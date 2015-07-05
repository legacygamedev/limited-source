VERSION 5.00
Begin VB.Form frmAssigned 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Assignments"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   3540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkUnAssign 
      Caption         =   "UnAssign"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Top             =   1560
      Width           =   735
   End
   Begin VB.ComboBox cboCharacter 
      Height          =   315
      Left            =   960
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1080
      Width           =   2535
   End
   Begin VB.ComboBox cboTo 
      Height          =   315
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
   Begin VB.ComboBox cboFrom 
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Character:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "To:"
      Height          =   195
      Left            =   2040
      TabIndex        =   2
      Top             =   720
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "From:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   390
   End
End
Attribute VB_Name = "frmAssigned"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboType_Change()
Call FillBoxes
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
Dim Packet As String
Packet = "ASSIGN" & SEP_CHAR & Trim(cboType.List(cboType.ListIndex)) & SEP_CHAR & Trim(cboFrom.List(cboFrom.ListIndex)) & SEP_CHAR & Trim(cboTo.List(cboTo.ListIndex)) & SEP_CHAR & Trim(cboCharacter.List(cboCharacter.ListIndex)) & SEP_CHAR & chkUnAssign.Value & SEP_CHAR & END_CHAR
Call SendData(Packet)
Unload Me
End Sub

Private Sub Form_Load()
'cboType.AddItem "Maps"
cboType.AddItem "Items"
'cboType.AddItem "NPCs"
'cboType.AddItem "Spells"

cboType.ListIndex = 0
'Fill the to/from boxes
Call FillBoxes

'Now fill the chars who are online :D

End Sub

Public Sub FillBoxes()
Dim I As Long
Dim m As Long
Dim N As String
N = cboType.List(cboType.ListIndex)

Select Case N
    Case "Items"
        m = MAX_ITEMS
    
    Case "Maps"
        m = MAX_MAPS
        
    Case "NPCs"
        m = MAX_NPCS
    
    Case "Spells"
        m = MAX_SPELLS

End Select

'Clear boxes.
cboTo.Clear
cboFrom.Clear

'Now fill the boxes
For I = 1 To m
    cboTo.AddItem CStr(I)
    cboFrom.AddItem CStr(I)
Next I
cboTo.ListIndex = 0
cboFrom.ListIndex = 0

End Sub

Public Sub FillChars()

End Sub
