VERSION 5.00
Begin VB.Form frmLibrary 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dual Solace"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDesc 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2535
      Width           =   2550
   End
   Begin VB.ListBox lstLibrary 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2310
      ItemData        =   "frmLibraries.frx":0000
      Left            =   120
      List            =   "frmLibraries.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   135
      Width           =   4320
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2865
      TabIndex        =   0
      Top             =   3090
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Author:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2865
      TabIndex        =   3
      Top             =   2490
      Width           =   1530
   End
   Begin VB.Label lblAuthor 
      Alignment       =   2  'Center
      Caption         =   "Unknown"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2865
      TabIndex        =   2
      Top             =   2760
      Width           =   1530
   End
End
Attribute VB_Name = "frmLibrary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
Dim i As Long
Dim FileName As String

For i = 0 To lstLibrary.ListCount - 1
    FileName = App.Path & "\Library\" & lstLibrary.List(i)
    Call PutVar(FileName, "DATA", "Enabled", lstLibrary.Selected(i))
Next i

Call LoadScripts

' Visible=False! Do NOT use Unload, or scripts willstop working due to me not having foresight in coding.
Me.Visible = False

ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("frmLibrary.frm", "cmdClose_Click", Err.Number, Err.Description)
End Sub

Private Sub lstLibrary_Click()
    Dim FileName As String
    
    FileName = App.Path & "\Library\" & frmLibrary.lstLibrary.List(frmLibrary.lstLibrary.ListIndex)
    If GetVar(FileName, "DATA", "Author") = "" Then
        lblAuthor = "Unknown"
    Else
        lblAuthor = GetVar(FileName, "DATA", "Author")
    End If
    txtDesc.Text = GetVar(FileName, "DATA", "Description")
End Sub
