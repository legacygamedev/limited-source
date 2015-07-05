VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AS Configuration File Maker"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3930
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   3930
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReload 
      Caption         =   "Reload"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3735
   End
   Begin VB.CommandButton cmdGenerateFile 
      Caption         =   "Save File"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   3735
   End
   Begin VB.TextBox txtPort 
      Height          =   375
      Left            =   1440
      MaxLength       =   5
      TabIndex        =   1
      Text            =   "Port"
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox txtIP 
      Height          =   375
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   0
      Text            =   "IP"
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox txtPassword 
      Height          =   375
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   3
      Text            =   "admin"
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Caption         =   "Change the IP and Port of your game securely."
      Height          =   615
      Left            =   0
      TabIndex        =   8
      Top             =   600
      Width           =   3975
   End
   Begin VB.Label lblPort 
      Caption         =   "Port:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label lblIPAddress 
      Caption         =   "IP Address:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblPassword 
      Caption         =   "Password:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGenerateFile_Click()
Dim FileName As String
Dim F As Long

    cmdGenerateFile.Enabled = False

    FileName = App.Path & CONFIG_FILE
    
    If Not FileExist(FileName) Then
    
        F = FreeFile
        Open FileName For Output As F
        Close F
        
    End If
    
    With Config
        .Password = Encryption(KEYWORD, Trim$(txtPassword.Text))
        .IP = Encryption(DEFAULT_KEY, Trim$(txtIP.Text))
        .Port = Val(txtPort.Text)
    End With

    F = FreeFile
    Open FileName For Binary As #F
        Put #F, , Config
    Close #F
    
    cmdGenerateFile.Enabled = True
    
End Sub

Private Sub cmdReload_Click()
Dim F As Long
Dim FileName As String

    FileName = App.Path & CONFIG_FILE
    
    F = FreeFile
    Open FileName For Binary As #F
        Get #F, , Config
    Close #F
    
    txtPassword.Text = Encryption(KEYWORD, Trim$(Config.Password))
    txtIP.Text = Encryption(DEFAULT_KEY, Trim$(Config.IP))
    txtPort.Text = Config.Port

End Sub

Private Sub Form_Load()
Dim F As Long
Dim FileName As String

    FileName = App.Path & CONFIG_FILE

    If Not FileExist(FileName) Then
    
        KEYWORD = DEFAULT_KEY
        Config.Password = Encryption(KEYWORD, DEFAULT_KEY)
        Config.IP = Encryption(DEFAULT_KEY, "127.0.0.1")
        Config.Port = 12000
        
        F = FreeFile
        Open FileName For Output As F
        Close F
        
        Open FileName For Binary As #F
            Put #F, , Config
        Close #F
        
    Else
    
Backwards:
        KEYWORD = InputBox("Enter the password to edit this config file.", "Password?", "admin")
    
        F = FreeFile
        Open FileName For Binary As #F
            Get #F, , Config
        Close #F
        
        Config.Password = Config.Password
        KEYWORD = Encryption(KEYWORD, Trim$(Config.Password))
        
    End If
    
    If Encryption(KEYWORD, Trim$(Config.Password)) = DEFAULT_KEY Then
Skip:
        KEYWORD = InputBox("You are still using the default password!" & vbNewLine & _
                           "Please enter a new one and make sure you change" & vbNewLine & _
                           "it in your server as well in Data\config.ini", "Error")
        
        If Len(Trim$(KEYWORD)) > 10 Then
            MsgBox "Re-enter your password! It has to be" & vbNewLine & _
                   "less than or equal to 10 characters.", , "Error"
            GoTo Skip
        End If
        
        Do While KEYWORD = DEFAULT_KEY Or KEYWORD = vbNullString
            GoTo Skip
            DoEvents
        Loop
        Config.Password = Encryption(KEYWORD, KEYWORD)
        Config.IP = Encryption(DEFAULT_KEY, "127.0.0.1")
        Config.Port = Config.Port
        Open FileName For Binary As #F
            Put #F, , Config
        Close #F
        GoTo Backwards
    End If
    
    If KEYWORD <> Trim$(Encryption(KEYWORD, Trim$(Config.Password))) Then
        MsgBox "Invalid password! If you have lost your password, " & vbNewLine & _
               "you can delete the config file to start over.", , "Denied"
        End
    End If
    
    txtPassword.Text = Trim$(Encryption(KEYWORD, Trim$(Config.Password)))
    txtIP.Text = Trim$(Encryption(DEFAULT_KEY, Trim$(Config.IP)))
    txtPort.Text = Config.Port
    
    cmdGenerateFile.Enabled = True

End Sub

Private Sub txtIP_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyDelete Then
        If KeyAscii <> vbKeyBack Then
            Select Case Chr$(KeyAscii)
                Case 0
                Case 1
                Case 2
                Case 3
                Case 4
                Case 5
                Case 6
                Case 7
                Case 8
                Case 9
                Case "."
                Case Else
                    KeyAscii = 0
                    Exit Sub
            End Select
        End If
    End If
End Sub

Private Sub txtPort_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyDelete Then
        If KeyAscii <> vbKeyBack Then
            Select Case Chr$(KeyAscii)
                Case 0
                Case 1
                Case 2
                Case 3
                Case 4
                Case 5
                Case 6
                Case 7
                Case 8
                Case 9
                Case Else
                    KeyAscii = 0
                    Exit Sub
            End Select
        End If
    End If
End Sub

Private Sub txtOffset_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyDelete Then
        If KeyAscii <> vbKeyBack Then
            Select Case Chr$(KeyAscii)
                Case 0
                Case 1
                Case 2
                Case 3
                Case 4
                Case 5
                Case 6
                Case 7
                Case 8
                Case 9
                Case Else
                    KeyAscii = 0
                    Exit Sub
            End Select
        End If
    End If
End Sub
