VERSION 5.00
Begin VB.Form frmGuide 
   Caption         =   "Form1"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   Picture         =   "frmGuide.frx":0000
   ScaleHeight     =   6615
   ScaleWidth      =   7890
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstTopics 
      Height          =   2790
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   2175
   End
   Begin VB.Frame TopicTitle 
      Caption         =   "Topic Title"
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   7575
      Begin VB.TextBox txtTopic 
         Height          =   3015
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   240
         Width           =   7335
      End
   End
   Begin VB.Label CharInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Close Window"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   0
      Left            =   6240
      TabIndex        =   4
      Top             =   240
      Width           =   1485
   End
   Begin VB.Label CharInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Topics:"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   21
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   510
   End
End
Attribute VB_Name = "frmGuide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CharInfo_Click(Index As Integer)
frmGuide.Visible = False
End Sub

Private Sub lstTopics_Click()
Dim FileName As String, inputdata As String
Dim hFile As Long
Dim X As Long

    txtTopic.text = ""
    TopicTitle.Caption = lstTopics.List(lstTopics.ListIndex)
    FileName = lstTopics.ListIndex + 1 & ".txt"

    X = 0
    
    If FileExist("Guide\" & FileName) = True And FileName <> "" Then
        hFile = FreeFile
        Open App.Path & "\Guide\" & FileName For Input As #hFile
            Do Until EOF(1)
                Line Input #1, inputdata
                If X = 0 Then
                    X = 1
                Else
                    txtTopic.text = txtTopic.text & inputdata & vbCrLf
                End If
            Loop
            
        Close #hFile
    End If
End Sub
