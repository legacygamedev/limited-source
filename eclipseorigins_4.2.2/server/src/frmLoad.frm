VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL32.OCX"
Begin VB.Form frmLoad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Eclipse Origins Server Loading..."
   ClientHeight    =   6600
   ClientLeft      =   6375
   ClientTop       =   4110
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLoad.frx":0000
   ScaleHeight     =   440
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   600
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrNotifications 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame fraConnecting 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Connecting..."
      Height          =   4095
      Left            =   0
      TabIndex        =   1
      Top             =   1920
      Width           =   9015
      Begin MSComctlLib.ProgressBar pbarLoading 
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   3600
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label lblProg 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Loading...."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   3240
         Width           =   8535
      End
      Begin VB.Label lblConnecting 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Connecting!!!"
         ForeColor       =   &H00FFFFFF&
         Height          =   1695
         Left            =   1920
         TabIndex        =   2
         Top             =   0
         Width           =   5175
      End
   End
   Begin VB.Label lblNotifications 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6360
      Width           =   9015
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim i As Long, x As Long
   On Error GoTo errorhandler
    InitMessages
    ' load options, set if they dont exist
    If Not FileExist(App.path & "\data\options.ini", True) Then
        Options.Game_Name = "Eclipse Origins"
        Options.Port = 7001
        Options.MOTD = "Welcome to Eclipse Origins."
        Options.Website = "http://www.eclipseorigins.com"
        Options.SilentStartup = 0
        Options.Key = GenerateOptionsKey
        PutVar App.path & "\data\options.ini", "OPTIONS", "MapCount", "300"
        SaveOptions
    Else
        LoadOptions
    End If
    
    If Options.SilentStartup = 1 Then
        Me.Hide
    Else
        Me.Show
    End If
    
    InitServer


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Form_Load", "frmLogin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

   On Error GoTo errorhandler

    End


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Form_QueryUnload", "frmLogin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub lblExitServer_Click()
    

   On Error GoTo errorhandler
    End
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "lblExitServer_Click", "frmLogin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub lblExitServer_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    

   On Error GoTo errorhandler
   
   lblExitServer.Font.Bold = True
   lblNewUser.Font.Bold = False
   lblExistingUser.Font.Bold = False
    


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "lblExitServer_MouseMove", "frmLogin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Private Sub tmrNotifications_Timer()
    Static LastNotification As String
    Static TimeShown As Long
    

   On Error GoTo errorhandler

    If lblNotifications.Caption <> "" Then
        If lblNotifications.Caption = LastNotification Then
            If TimeShown >= 6 Then
                LastNotification = ""
                lblNotifications.Caption = ""
                TimeShown = 0
            Else
                TimeShown = TimeShown + 1
            End If
        Else
            LastNotification = lblNotifications.Caption
            TimeShown = 0
        End If
    End If


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "tmrNotifications_Timer", "frmLogin", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

