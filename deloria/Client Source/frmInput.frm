VERSION 5.00
Begin VB.Form frmInput 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configure Input"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   480
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   120
      ScaleHeight     =   2415
      ScaleWidth      =   4455
      TabIndex        =   40
      Top             =   120
      Width           =   4455
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Let go of all your joypads keys..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   0
         TabIndex        =   42
         Top             =   0
         Width           =   4095
      End
      Begin VB.Label lblSex 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(3 Secs)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3120
         TabIndex        =   41
         Top             =   360
         Width           =   1020
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
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
      Left            =   3480
      TabIndex        =   38
      Top             =   2640
      Width           =   1095
   End
   Begin VB.PictureBox picKey 
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   120
      ScaleHeight     =   2415
      ScaleWidth      =   4455
      TabIndex        =   34
      Top             =   120
      Visible         =   0   'False
      Width           =   4455
      Begin VB.Label lblSecs 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(3 Secs)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3000
         TabIndex        =   39
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label lblKey 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1200
         TabIndex        =   37
         Top             =   0
         Width           =   405
      End
      Begin VB.Label lblMsg2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hold desired key now..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   0
         TabIndex        =   36
         Top             =   240
         Width           =   2985
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Configure Key:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   0
         TabIndex        =   35
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Joypad"
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
      Left            =   1200
      TabIndex        =   33
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Keyboard"
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
      TabIndex        =   32
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Keyboard"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   4455
      Begin VB.TextBox txtKUp 
         Alignment       =   2  'Center
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
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtKLeft 
         Alignment       =   2  'Center
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
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtKDown 
         Alignment       =   2  'Center
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
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtKRight 
         Alignment       =   2  'Center
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
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtKAttack 
         Alignment       =   2  'Center
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
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txtKRun 
         Alignment       =   2  'Center
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
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txtKEnter 
         Alignment       =   2  'Center
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
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Up Button"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1560
         TabIndex        =   31
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Down Button"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1560
         TabIndex        =   30
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Left Button"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Right Button"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3000
         TabIndex        =   28
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Attack"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Run"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1560
         TabIndex        =   26
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Pick Up"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3000
         TabIndex        =   25
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Click on the text boxes to set the buttons."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   2160
         Width           =   4215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Joypad"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtEnter 
         Alignment       =   2  'Center
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
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txtRun 
         Alignment       =   2  'Center
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
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txtAttack 
         Alignment       =   2  'Center
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
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txtRight 
         Alignment       =   2  'Center
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
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtDown 
         Alignment       =   2  'Center
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
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtLeft 
         Alignment       =   2  'Center
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
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtUp 
         Alignment       =   2  'Center
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
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Click on the text boxes to set the buttons."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   2160
         Width           =   4215
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Pick Up"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3000
         TabIndex        =   11
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Run"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1560
         TabIndex        =   10
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Attack"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Right Button"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3000
         TabIndex        =   4
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Left Button"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Down Button"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1560
         TabIndex        =   2
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Up Button"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1560
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Control As String
Dim TC As String
Dim TempX As Long
Dim TempY As Long
Dim TempB As Long

Private Sub Command1_Click()
    MsgBox "Not Implemented Yet!"
    'Frame2.Visible = True
    'Frame1.Visible = False
End Sub

Private Sub Command2_Click()
    Frame2.Visible = False
    Frame1.Visible = True
End Sub

Private Sub Command3_Click()
Dim FileName As String
    FileName = App.Path & "\Controls.ini"
    
    WriteINI "KEYBOARD", "Up", 0, FileName
    WriteINI "KEYBOARD", "Down", 0, FileName
    WriteINI "KEYBOARD", "Left", 0, FileName
    WriteINI "KEYBOARD", "Right", 0, FileName
    WriteINI "KEYBOARD", "Attack", 0, FileName
    WriteINI "KEYBOARD", "Run", 0, FileName
    WriteINI "KEYBOARD", "Enter", 0, FileName
    
    Unload Me
End Sub

Private Sub Form_Load()
    txtUp.Text = JUp
    txtDown.Text = JDown
    txtLeft.Text = JLeft
    txtRight.Text = JRight
    txtAttack.Text = JAttack
    txtRun.Text = JRun
    txtEnter.Text = JEnter
    TC = 3
    Picture1.Visible = True
    Timer2.Enabled = True
End Sub

Private Sub Timer1_Timer()
Dim x As Long, y As Long
Dim CurrJoy As JOYINFO
Dim Chosen As Long, Chose As Byte
Dim FileName As String

    FileName = App.Path & "\Controls.ini"

    joyGetPos 0, CurrJoy
    x = CurrJoy.wXpos
    y = CurrJoy.wYpos
    
    If x <> TempX Then
        Chosen = x
        Chose = 1
    ElseIf y <> TempY Then
        Chosen = y
        Chose = 2
    ElseIf CurrJoy.wButtons <> TempB Then
        Chosen = CurrJoy.wButtons
        Chose = 3
    End If
    
    TC = TC - 0.1
    lblSecs.Caption = "(" & TC & " Secs)"
    
    If TC <= 0 Then
        If Control = "Joypad Up" Then
            WriteINI "JOYPAD", "Up", Val(Chosen), FileName
            WriteINI "JOYPAD", "UpC", Val(Chose), FileName
            JUp = Val(Chosen)
            JUpC = Val(Chose)
            txtUp.Text = JUp
        ElseIf Control = "Joypad Down" Then
            WriteINI "JOYPAD", "Down", Val(Chosen), FileName
            WriteINI "JOYPAD", "DownC", Val(Chose), FileName
            JDown = Val(Chosen)
            JDownC = Val(Chose)
            txtDown.Text = JDown
        ElseIf Control = "Joypad Right" Then
            WriteINI "JOYPAD", "Right", Val(Chosen), FileName
            WriteINI "JOYPAD", "RightC", Val(Chose), FileName
            JRight = Val(Chosen)
            JRightC = Val(Chose)
            txtRight.Text = JRight
        ElseIf Control = "Joypad Left" Then
            WriteINI "JOYPAD", "Left", Val(Chosen), FileName
            WriteINI "JOYPAD", "LeftC", Val(Chose), FileName
            JLeft = Val(Chosen)
            JLeftC = Val(Chose)
            txtLeft.Text = JLeft
        ElseIf Control = "Joypad Attack" Then
            WriteINI "JOYPAD", "Attack", Val(Chosen), FileName
            WriteINI "JOYPAD", "AttackC", Val(Chose), FileName
            JAttack = Val(Chosen)
            JAttackC = Val(Chose)
            txtAttack.Text = JAttack
        ElseIf Control = "Joypad Run" Then
            WriteINI "JOYPAD", "Run", Val(Chosen), FileName
            WriteINI "JOYPAD", "RunC", Val(Chose), FileName
            JRun = Val(Chosen)
            JRunC = Val(Chose)
            txtRun.Text = JRun
        ElseIf Control = "Joypad Enter" Then
            WriteINI "JOYPAD", "Enter", Val(Chosen), FileName
            WriteINI "JOYPAD", "EnterC", Val(Chose), FileName
            JEnter = Val(Chosen)
            JEnterC = Val(Chose)
            txtEnter.Text = JEnter
        End If
        Timer1.Enabled = False
        picKey.Visible = False
        Control = ""
    End If
End Sub

Private Sub Timer2_Timer()
Dim x As Long, y As Long
Dim CurrJoy As JOYINFO
Dim Chosen As Long
    joyGetPos 0, CurrJoy
    x = CurrJoy.wXpos
    y = CurrJoy.wYpos
    
    TempX = x
    TempY = y
    TempB = CurrJoy.wButtons
    
    TC = TC - 0.1
    lblSex.Caption = "(" & TC & " Secs)"
    
    If TC <= 0 Then
        Picture1.Visible = False
        Timer2.Enabled = False
    End If
End Sub

Private Sub txtUp_Click()
    Control = "Joypad Up"
    picKey.Visible = True
    Timer1.Enabled = True
    lblKey.Caption = Trim(Control)
    TC = 3
End Sub

Private Sub txtDown_Click()
    Control = "Joypad Down"
    picKey.Visible = True
    Timer1.Enabled = True
    lblKey.Caption = Trim(Control)
    TC = 3
End Sub

Private Sub txtRight_Click()
    Control = "Joypad Right"
    picKey.Visible = True
    Timer1.Enabled = True
    lblKey.Caption = Trim(Control)
    TC = 3
End Sub

Private Sub txtLeft_Click()
    Control = "Joypad Left"
    picKey.Visible = True
    Timer1.Enabled = True
    lblKey.Caption = Trim(Control)
    TC = 3
End Sub

Private Sub txtAttack_Click()
    Control = "Joypad Attack"
    picKey.Visible = True
    Timer1.Enabled = True
    lblKey.Caption = Trim(Control)
    TC = 3
End Sub

Private Sub txtRun_Click()
    Control = "Joypad Run"
    picKey.Visible = True
    Timer1.Enabled = True
    lblKey.Caption = Trim(Control)
    TC = 3
End Sub

Private Sub txtEnter_Click()
    Control = "Joypad Enter"
    picKey.Visible = True
    Timer1.Enabled = True
    lblKey.Caption = Trim(Control)
    TC = 3
End Sub
