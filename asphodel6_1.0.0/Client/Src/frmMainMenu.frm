VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMainMenu 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asphodel Source "
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
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
   Icon            =   "frmMainMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   400
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   368
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrCheckTCP 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   5040
      Top             =   5520
   End
   Begin VB.CheckBox chkRemember 
      BackColor       =   &H00FF8080&
      Height          =   240
      Left            =   480
      TabIndex        =   6
      Top             =   4680
      Width           =   200
   End
   Begin RichTextLib.RichTextBox txtNews 
      Height          =   2820
      Left            =   600
      TabIndex        =   5
      Top             =   1110
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   4974
      _Version        =   393217
      BackColor       =   4210752
      BorderStyle     =   0
      Enabled         =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMainMenu.frx":08CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtUsername 
      Height          =   285
      Left            =   2370
      TabIndex        =   1
      Top             =   4515
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   503
      _Version        =   393217
      BackColor       =   4210752
      BorderStyle     =   0
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      MaxLength       =   20
      Appearance      =   0
      TextRTF         =   $"frmMainMenu.frx":0951
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtPassword 
      Height          =   285
      Left            =   2370
      TabIndex        =   2
      Top             =   4920
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   503
      _Version        =   393217
      BackColor       =   4210752
      BorderStyle     =   0
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      MaxLength       =   20
      Appearance      =   0
      TextRTF         =   $"frmMainMenu.frx":09CC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblStatus 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Server: Checking..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   2430
      TabIndex        =   8
      Top             =   3915
      Width           =   2520
   End
   Begin VB.Label lblRemember 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Remember"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   4920
      UseMnemonic     =   0   'False
      Width           =   795
   End
   Begin VB.Label lblCancel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3675
      TabIndex        =   4
      Top             =   5355
      Width           =   1440
   End
   Begin VB.Label lblNewAccount 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   1935
      TabIndex        =   3
      Top             =   5355
      Width           =   1635
   End
   Begin VB.Label lblLogin 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   420
      TabIndex        =   0
      Top             =   5355
      Width           =   1440
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ------------------------------------------
' --               Asphodel               --
' ------------------------------------------

Private Sub Form_Load()

    Me.Picture = LoadPicture(App.Path & GFX_PATH & "interface\mainmenuwindow.bmp")
    SetWindowLong txtNews.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT
    SetWindowLong txtUsername.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT
    SetWindowLong txtPassword.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT
    
    If Remember Then
        frmMainMenu.txtUsername.Text = GetVar(App.Path & "/info.ini", "BASIC", "Account")
        frmMainMenu.txtPassword.Text = GetVar(App.Path & "/info.ini", "BASIC", "Password")
        frmMainMenu.chkRemember.Value = 1
    End If
    
    ' make the text centered
    txtNews.SelAlignment = 2
    
    If ConnectToServer Then
        tmrCheckTCP.Enabled = False
        
        If Not CheckedStuff And Not CheckedTwice Then
            Check_Password
            
            Do Until Password_Confirmed And Config_Received
                DoEvents
                Sleep 1
            Loop
        End If
        
        lblStatus.Caption = "Server: Online"
        lblStatus.ForeColor = &H8000&
    Else
        lblStatus.Caption = "Server: Offline"
        lblStatus.ForeColor = &H80&
        tmrCheckTCP.Enabled = True
    End If
    
End Sub

Private Sub chkRemember_Click()
Dim FileName As String

    If LenB(Trim$(txtUsername.Text)) > 0 Then
        If LenB(Trim$(txtPassword.Text)) > 0 Then
            
            FileName = App.Path & "/info.ini"
            
            If chkRemember.Value = 1 Then
                PutVar FileName, "BASIC", "Account", Trim$(txtUsername.Text)
                PutVar FileName, "BASIC", "Password", Trim$(txtPassword.Text)
                PutVar FileName, "BASIC", "Remember", CStr(1)
                Remember = True
            Else
                PutVar FileName, "BASIC", "Account", vbNullString
                PutVar FileName, "BASIC", "Password", vbNullString
                PutVar FileName, "BASIC", "Remember", CStr(0)
                Remember = False
            End If
            Exit Sub
        
        End If
    End If
    
    chkRemember.Value = 0
    PutVar FileName, "BASIC", "Remember", CStr(0)
    Remember = False
    
End Sub

Private Sub Form_GotFocus()
    txtUsername.SetFocus
End Sub

Private Sub lblLogin_Click()
Dim FileName As String

    FileName = App.Path & "/info.ini"
    
    If Remember Then
        If GetVar(FileName, "BASIC", "Account") <> txtUsername.Text Then PutVar FileName, "BASIC", "Account", txtUsername.Text
        If GetVar(FileName, "BASIC", "Password") <> txtPassword.Text Then PutVar FileName, "BASIC", "Password", txtPassword.Text
    End If
    
    If isLoginLegal(txtUsername.Text, txtPassword.Text) Then
        
        ' checking if server is online block
        Me.Visible = False
        
        SetStatus "Connecting to server..."
        
        If Not ConnectToServer Then
            MsgBox "The server appears to be down." & vbNewLine & _
                   "Please check back later!", , "Error"
            Me.Visible = True
            frmStatus.Visible = False
            Exit Sub
        End If
        
        MenuState Menu_State.Login_
        
    End If
    
End Sub

Private Sub lblNewAccount_Click()
Dim Name As String
Dim Password As String

    Name = Trim$(txtUsername.Text)
    Password = Trim$(txtPassword.Text)
    
    If isLoginLegal(Name, Password) Then
    
        If Not isStringLegal(Name) Then Exit Sub
        
        ' checking if server is online block
        Me.Visible = False
        
        SetStatus "Connecting to server..."
        
        If Not ConnectToServer Then
            MsgBox "The server appears to be down." & vbNewLine & _
                   "Please check back later!", , "Error"
            Me.Visible = True
            frmStatus.Visible = False
            Exit Sub
        End If
        
        MenuState Menu_State.NewAccount_
        
    End If
    
End Sub

Private Sub lblCancel_Click()
    DestroyGame
End Sub

Private Sub tmrCheckTCP_Timer()

    lblStatus.Caption = "Server: Checking..."
    lblStatus.ForeColor = &H404040
    DoEvents
    
    If ConnectToServer Then
        If Not CheckedStuff And Not CheckedTwice Then
            Check_Password
            
            Do Until Password_Confirmed And Config_Received
                DoEvents
                Sleep 1
            Loop
        End If
        
        lblStatus.Caption = "Server: Online"
        lblStatus.ForeColor = &H8000&
        tmrCheckTCP.Enabled = False
    Else
        lblStatus.Caption = "Server: Offline"
        lblStatus.ForeColor = &H80&
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            lblLogin_Click
        Case vbKeyEscape
            lblCancel_Click
    End Select
End Sub

Private Sub txtNews_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub txtUsername_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub
