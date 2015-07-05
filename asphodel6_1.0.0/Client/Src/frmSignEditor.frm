VERSION 5.00
Begin VB.Form frmSignEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editing Sign"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      MaxLength       =   10
      TabIndex        =   5
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   2400
      Width           =   2535
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "Accept"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   2535
   End
   Begin VB.TextBox txtSign 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   1440
      MaxLength       =   318
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmSignEditor.frx":0000
      Top             =   480
      Width           =   3735
   End
   Begin VB.ListBox lstSections 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lblShrink 
      Caption         =   "[shrink]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblAdd 
      Caption         =   "[add]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label lblName 
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   1320
      X2              =   1320
      Y1              =   465
      Y2              =   2275
   End
End
Attribute VB_Name = "frmSignEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAccept_Click()

    If LenB(Trim$(txtName.Text)) < 1 Then
        MsgBox "You need to add a name for this sign!", , "Error"
        Exit Sub
    End If
    
    SignEditorOk
    
End Sub

Private Sub cmdCancel_Click()
    SignEditorCancel
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Editor = 0
End Sub

Private Sub lblAdd_Click()
Dim LoopI As Long
Dim OldIndex As Long

    ReDim Preserve SignSection(0 To UBound(SignSection) + 1)
    ReDim Preserve Sign(EditorIndex).Section(0 To UBound(SignSection))
    
    OldIndex = lstSections.ListIndex
    
    lstSections.Clear
    
    For LoopI = 0 To UBound(SignSection)
        lstSections.AddItem "Section " & LoopI
    Next
    
    lstSections.ListIndex = OldIndex
    
End Sub

Private Sub lblShrink_Click()
Dim LoopI As Long
Dim OldIndex As Long

    If UBound(SignSection) = 0 Then Exit Sub
    
    ReDim Preserve SignSection(0 To UBound(SignSection) - 1)
    ReDim Preserve Sign(EditorIndex).Section(0 To UBound(SignSection))
    
    OldIndex = lstSections.ListIndex
    
    If OldIndex > UBound(SignSection) Then OldIndex = UBound(SignSection)
    
    lstSections.Clear
    
    For LoopI = 0 To UBound(SignSection)
        lstSections.AddItem "Section " & LoopI
    Next
    
    lstSections.ListIndex = OldIndex
    
End Sub

Private Sub lstSections_Click()
    txtSign.Text = SignSection(lstSections.ListIndex)
End Sub

Private Sub txtSign_Change()
    SignSection(lstSections.ListIndex) = txtSign.Text
End Sub
