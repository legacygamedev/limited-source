VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLogs 
   Caption         =   "Log Loader"
   ClientHeight    =   4995
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   ScaleHeight     =   4995
   ScaleWidth      =   6645
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6120
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtLogs 
      Height          =   5055
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   6615
   End
   Begin VB.Menu Logs 
      Caption         =   "File"
      Begin VB.Menu mnuLogs 
         Caption         =   "Load Logs.."
      End
   End
End
Attribute VB_Name = "frmLogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub LoadLogs()
txtLogs.Text = ""
Dim strFileName As String
With Me.CommonDialog1
        .filename = ""
        .Filter = "All Files|*.*"
        .ShowOpen
        
        If Trim(.filename) > " " Then
            strFileName = .filename
        End If
End With
Dim strData As String
Dim lngFile As Long
lngFile = FreeFile()
If strFileName <> "" Then
    Open strFileName For Input As #lngFile
        Do Until (EOF(1))
        Line Input #lngFile, strData
        txtLogs.Text = txtLogs.Text & strData & vbCrLf
        Loop
    Close #1
End If
End Sub

Private Sub Form_Resize()
txtLogs.Width = Me.Width - 100
txtLogs.Height = Me.Height - 800
End Sub

Private Sub mnuLogs_Click()
LoadLogs
End Sub


