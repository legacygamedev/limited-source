VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmSuggestions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Submit Your Suggestions !"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   3810
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   2160
      Width           =   3855
   End
   Begin VB.CommandButton btnSend 
      Caption         =   "Send"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   1560
      Width           =   3855
   End
   Begin RichTextLib.RichTextBox txtSuggestReport 
      Height          =   1545
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3810
      _ExtentX        =   6720
      _ExtentY        =   2725
      _Version        =   393217
      BackColor       =   16777215
      Appearance      =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmSuggestions.frx":0000
   End
End
Attribute VB_Name = "frmSuggestions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
Unload Me
End Sub

Private Sub btnSend_Click()
    If Not txtSuggestReport.Text = "" Then
        Call SendSuggestions(Trim(txtSuggestReport.Text))
        txtSuggestReport.Text = ""
        frmSuggestions.Visible = False
    End If
End Sub

Private Sub Form_Load()
MsgBox " Improper use of This System will Result in Banishment!"
End Sub
