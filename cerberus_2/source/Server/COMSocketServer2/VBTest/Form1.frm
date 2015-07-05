VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Socket Server"
   ClientHeight    =   4344
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   2544
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4344
   ScaleWidth      =   2544
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox SignOnAsUnicode 
      Caption         =   "Send sign on as unicode"
      Height          =   252
      Left            =   240
      TabIndex        =   12
      Top             =   600
      Value           =   1  'Checked
      Width           =   2172
   End
   Begin VB.TextBox ShutdownAfter 
      Enabled         =   0   'False
      Height          =   288
      Left            =   1680
      TabIndex        =   11
      Text            =   "10"
      Top             =   2160
      Width           =   612
   End
   Begin VB.Frame ShutdownFrame 
      Height          =   1452
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   2292
      Begin VB.CheckBox ShutdownEnabled 
         Caption         =   "Shutdown after"
         Height          =   252
         Left            =   120
         TabIndex        =   13
         Top             =   0
         Width           =   1332
      End
      Begin VB.OptionButton CloseSocket 
         Caption         =   "Close Socket"
         Enabled         =   0   'False
         Height          =   252
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   1692
      End
      Begin VB.OptionButton ShutdownSocket 
         Caption         =   "Shutdown Socket"
         Enabled         =   0   'False
         Height          =   252
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1692
      End
      Begin VB.OptionButton ShutdownAfterWrite 
         Caption         =   "Shutdown after write"
         Enabled         =   0   'False
         Height          =   252
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   1692
      End
   End
   Begin VB.CheckBox ShowDataPackets 
      Caption         =   "Show data packets"
      Height          =   252
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Value           =   1  'Checked
      Width           =   1692
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   1560
      TabIndex        =   1
      Text            =   "5001"
      Top             =   240
      Width           =   612
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create Server"
      Height          =   492
      Left            =   120
      TabIndex        =   0
      Top             =   3720
      Width           =   2292
   End
   Begin VB.Frame DataIsFrame 
      Enabled         =   0   'False
      Height          =   1092
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   2292
      Begin VB.OptionButton DataIsString 
         Caption         =   "String"
         Enabled         =   0   'False
         Height          =   252
         Left            =   120
         TabIndex        =   6
         Top             =   708
         Width           =   2052
      End
      Begin VB.OptionButton DataIsBytes 
         Caption         =   "Bytes"
         Enabled         =   0   'False
         Height          =   252
         Left            =   120
         TabIndex        =   5
         Top             =   348
         Value           =   -1  'True
         Width           =   2052
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Port to listen on:"
      Height          =   372
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1452
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    Dim server As JBSOCKETSERVERLib.server
    Set server = CreateSocketServer(CLng(Text1.Text))
    
    Dim frm As Form2
    Set frm = New Form2
    
    frm.SetServer server
    frm.ShowDataPackets.Value = ShowDataPackets.Value
    frm.DataIsBytes.Value = DataIsBytes.Value
    frm.DataIsString.Value = DataIsString.Value
    frm.SignOnAsUnicode.Value = SignOnAsUnicode.Value
    
    If ShutdownEnabled.Value Then
    
        If ShutdownAfterWrite.Value Then
            frm.ShutdownAfterWrite = CLng(ShutdownAfter.Text)
        ElseIf ShutdownSocket.Value Then
            frm.ShutdownAfter = CLng(ShutdownAfter.Text)
        ElseIf CloseSocket.Value Then
            frm.CloseAfter = CLng(ShutdownAfter.Text)
        End If
    
    End If
    
    server.StartListening
    
    frm.Show , Me

End Sub

Private Sub ShowDataPackets_Click()

    DataIsFrame.Enabled = ShowDataPackets.Value
    DataIsBytes.Enabled = ShowDataPackets.Value
    DataIsString.Enabled = ShowDataPackets.Value

End Sub

Private Sub ShutdownEnabled_Click()

    ShutdownAfterWrite.Enabled = ShutdownEnabled.Value
    ShutdownSocket.Enabled = ShutdownEnabled.Value
    CloseSocket.Enabled = ShutdownEnabled.Value
    ShutdownAfter.Enabled = ShutdownEnabled.Value
    
End Sub
