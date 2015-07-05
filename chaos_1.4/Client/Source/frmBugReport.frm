VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmBugReport 
   Caption         =   "Bug Report"
   ClientHeight    =   2685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3810
   LinkTopic       =   "Form1"
   Picture         =   "frmBugReport.frx":0000
   ScaleHeight     =   2685
   ScaleWidth      =   3810
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   2160
      Width           =   3855
   End
   Begin VB.CommandButton btnSend 
      Caption         =   "Send"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   1560
      Width           =   3855
   End
   Begin RichTextLib.RichTextBox txtBugReport 
      Height          =   1545
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3810
      _ExtentX        =   6720
      _ExtentY        =   2725
      _Version        =   393217
      BackColor       =   16777215
      Appearance      =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmBugReport.frx":2485
   End
End
Attribute VB_Name = "frmBugReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
Unload Me
End Sub

Private Sub btnSend_Click()
    If Not txtBugReport.Text = "" Then
        Call SendBugReport(Trim(txtBugReport.Text))
        txtBugReport.Text = ""
        frmBugReport.Visible = False
    End If
End Sub

Private Sub Form_Load()
MsgBox " DO NOT SPAM OR YOU WILL BE BANNED!"
End Sub
