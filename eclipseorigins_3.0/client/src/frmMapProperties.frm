VERSION 5.00
Begin VB.Form frmEditor_MapProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Properties"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   487
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   441
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame7 
      Caption         =   "Fog"
      Height          =   1935
      Left            =   4440
      TabIndex        =   48
      Top             =   4800
      Width           =   2055
      Begin VB.HScrollBar scrlFogOpacity 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   53
         Top             =   1620
         Width           =   1575
      End
      Begin VB.HScrollBar ScrlFog 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   51
         Top             =   480
         Width           =   1575
      End
      Begin VB.HScrollBar ScrlFogSpeed 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   49
         Top             =   1050
         Width           =   1575
      End
      Begin VB.Label lblFogOpacity 
         Caption         =   "Fog Opacity: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   1380
         Width           =   1815
      End
      Begin VB.Label lblFog 
         Caption         =   "Fog: None"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblFogSpeed 
         Caption         =   "Fog Speed: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   810
         Width           =   1815
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Weather"
      Height          =   1455
      Left            =   120
      TabIndex        =   41
      Top             =   4800
      Width           =   2055
      Begin VB.HScrollBar scrlWeatherIntensity 
         Height          =   255
         Left            =   120
         Max             =   100
         TabIndex        =   44
         Top             =   1080
         Width           =   1815
      End
      Begin VB.ComboBox CmbWeather 
         Height          =   315
         ItemData        =   "frmMapProperties.frx":0000
         Left            =   120
         List            =   "frmMapProperties.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lblWeatherIntensity 
         Caption         =   "Intensity: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Weather Type:"
         Height          =   195
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   1275
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Map Sound Effect"
      Height          =   975
      Left            =   120
      TabIndex        =   39
      Top             =   6240
      Width           =   2055
      Begin VB.ComboBox cmbSound 
         Height          =   315
         ItemData        =   "frmMapProperties.frx":0045
         Left            =   120
         List            =   "frmMapProperties.frx":0047
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Map Overlay"
      Height          =   2415
      Left            =   2280
      TabIndex        =   32
      Top             =   4800
      Width           =   2055
      Begin VB.HScrollBar scrlA 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   46
         Top             =   1920
         Width           =   855
      End
      Begin VB.HScrollBar ScrlR 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   35
         Top             =   480
         Width           =   855
      End
      Begin VB.HScrollBar ScrlG 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   34
         Top             =   960
         Width           =   855
      End
      Begin VB.HScrollBar ScrlB 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   33
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lblA 
         Caption         =   "Opacity: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label lblR 
         Caption         =   "Red: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblG 
         Caption         =   "Green: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblB 
         Caption         =   "Blue: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1200
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Music"
      Height          =   3255
      Left            =   4440
      TabIndex        =   27
      Top             =   1440
      Width           =   2055
      Begin VB.CommandButton cmdPlay 
         Caption         =   "Play"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   2520
         Width           =   1815
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   2880
         Width           =   1815
      End
      Begin VB.ListBox lstMusic 
         Height          =   2010
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame frmMaxSizes 
      Caption         =   "Max Sizes"
      Height          =   975
      Left            =   120
      TabIndex        =   22
      Top             =   3720
      Width           =   2055
      Begin VB.TextBox txtMaxX 
         Height          =   285
         Left            =   1080
         TabIndex        =   24
         Text            =   "0"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtMaxY 
         Height          =   285
         Left            =   1080
         TabIndex        =   23
         Text            =   "0"
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Max X:"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   270
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Max Y:"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   630
         Width           =   585
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Map Links"
      Height          =   1455
      Left            =   120
      TabIndex        =   16
      Top             =   480
      Width           =   2055
      Begin VB.TextBox txtUp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   720
         TabIndex        =   20
         Text            =   "0"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtDown 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   720
         TabIndex        =   19
         Text            =   "0"
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtRight 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1320
         TabIndex        =   18
         Text            =   "0"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   120
         TabIndex        =   17
         Text            =   "0"
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblMap 
         BackStyle       =   0  'Transparent
         Caption         =   "Current map: 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame fraMapSettings 
      Caption         =   "Map Settings"
      Height          =   855
      Left            =   2280
      TabIndex        =   13
      Top             =   480
      Width           =   4215
      Begin VB.ComboBox cmbMoral 
         Height          =   315
         ItemData        =   "frmMapProperties.frx":0049
         Left            =   960
         List            =   "frmMapProperties.frx":0053
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Moral:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Boot Settings"
      Height          =   1575
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   2055
      Begin VB.TextBox txtBootMap 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Text            =   "0"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtBootX 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         TabIndex        =   8
         Text            =   "0"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtBootY 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         Text            =   "0"
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Boot Map:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   870
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Boot X:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   645
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Boot Y:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   630
      End
   End
   Begin VB.Frame fraNPCs 
      Caption         =   "NPCs"
      Height          =   3255
      Left            =   2280
      TabIndex        =   4
      Top             =   1440
      Width           =   2055
      Begin VB.CheckBox chkDoNotAutoSpawn 
         Caption         =   "Don't AutoSpawn"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   2880
         Width           =   1815
      End
      Begin VB.ListBox lstNpcs 
         Height          =   2010
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1815
      End
      Begin VB.ComboBox cmbNpc 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2400
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5640
      TabIndex        =   3
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   6840
      Width           =   975
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "frmEditor_MapProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkDoNotAutoSpawn_Click()
    If lstNpcs.ListIndex > -1 Then
        Map.NpcSpawnType(lstNpcs.ListIndex + 1) = chkDoNotAutoSpawn.value
    End If
End Sub

Private Sub cmdPlay_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    StopMusic
    PlayMusic lstMusic.List(lstMusic.ListIndex)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdPlay_Click", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdStop_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    StopMusic
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdStop_Click", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdOk_Click()
    Dim i As Long
    Dim sTemp As Long
    Dim x As Long, x2 As Long
    Dim y As Long, y2 As Long
    Dim tempArr() As TileRec
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not IsNumeric(txtMaxX.Text) Then txtMaxX.Text = Map.MaxX
    If Val(txtMaxX.Text) < MAX_MAPX Then txtMaxX.Text = MAX_MAPX
    If Val(txtMaxX.Text) > MAX_BYTE Then txtMaxX.Text = MAX_BYTE
    If Not IsNumeric(txtMaxY.Text) Then txtMaxY.Text = Map.MaxY
    If Val(txtMaxY.Text) < MAX_MAPY Then txtMaxY.Text = MAX_MAPY
    If Val(txtMaxY.Text) > MAX_BYTE Then txtMaxY.Text = MAX_BYTE

    With Map
        .name = Trim$(txtName.Text)
        If lstMusic.ListIndex >= 0 Then
            .Music = lstMusic.List(lstMusic.ListIndex)
        Else
            .Music = vbNullString
        End If
        If cmbSound.ListIndex >= 0 Then
            .BGS = cmbSound.List(cmbSound.ListIndex)
        Else
            .BGS = vbNullString
        End If
        .Up = Val(txtUp.Text)
        .Down = Val(txtDown.Text)
        .Left = Val(txtLeft.Text)
        .Right = Val(txtRight.Text)
        .Moral = cmbMoral.ListIndex
        .BootMap = Val(txtBootMap.Text)
        .BootX = Val(txtBootX.Text)
        .BootY = Val(txtBootY.Text)
        
        .Weather = CmbWeather.ListIndex
        .WeatherIntensity = scrlWeatherIntensity.value
        
        .Fog = ScrlFog.value
        .FogSpeed = ScrlFogSpeed.value
        .FogOpacity = scrlFogOpacity.value
        
        .Red = ScrlR.value
        .Green = ScrlG.value
        .Blue = ScrlB.value
        .Alpha = scrlA.value
        
        ' set the data before changing it
        tempArr = Map.Tile
        x2 = Map.MaxX
        y2 = Map.MaxY
        ' change the data
        .MaxX = Val(txtMaxX.Text)
        .MaxY = Val(txtMaxY.Text)
        ReDim Map.Tile(0 To .MaxX, 0 To .MaxY)

        If x2 > .MaxX Then x2 = .MaxX
        If y2 > .MaxY Then y2 = .MaxY

        For x = 0 To x2
            For y = 0 To y2
                .Tile(x, y) = tempArr(x, y)
            Next
        Next

        ClearTempTile
    End With

    Call UpdateDrawMapName
    initAutotiles
    Unload frmEditor_MapProperties
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdOk_Click", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Unload frmEditor_MapProperties
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lstNpcs_Click()
Dim tmpString() As String
Dim npcNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' exit out if needed
    If Not cmbNpc.ListCount > 0 Then Exit Sub
    If Not lstNpcs.ListCount > 0 Then Exit Sub
    
    ' set the combo box properly
    tmpString = Split(lstNpcs.List(lstNpcs.ListIndex))
    npcNum = CLng(Left$(tmpString(0), Len(tmpString(0)) - 1))
    cmbNpc.ListIndex = Map.Npc(npcNum)
    chkDoNotAutoSpawn.value = Map.NpcSpawnType(npcNum)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstNpcs_Click", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbNpc_Click()
Dim tmpString() As String
Dim npcNum As Long
Dim x As Long, tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' exit out if needed
    If Not cmbNpc.ListCount > 0 Then Exit Sub
    If Not lstNpcs.ListCount > 0 Then Exit Sub
    
    ' set the combo box properly
    tmpString = Split(cmbNpc.List(cmbNpc.ListIndex))
    ' make sure it's not a clear
    If Not cmbNpc.List(cmbNpc.ListIndex) = "No NPC" Then
        npcNum = CLng(Left$(tmpString(0), Len(tmpString(0)) - 1))
        Map.Npc(lstNpcs.ListIndex + 1) = npcNum
    Else
        Map.Npc(lstNpcs.ListIndex + 1) = 0
    End If
    
    ' re-load the list
    tmpIndex = lstNpcs.ListIndex
    lstNpcs.Clear
    For x = 1 To MAX_MAP_NPCS
        If Map.Npc(x) > 0 Then
        lstNpcs.AddItem x & ": " & Trim$(Npc(Map.Npc(x)).name)
        Else
            lstNpcs.AddItem x & ": No NPC"
        End If
    Next
    lstNpcs.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbNpc_Click", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlA_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblA.Caption = "Opacity: " & scrlA.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ScrlA_Change", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub ScrlB_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblB.Caption = "Blue: " & ScrlB.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ScrlB_Change", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub ScrlFog_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If ScrlFog.value = 0 Then
        lblFog.Caption = "None."
    Else
        lblFog.Caption = "Fog: " & ScrlFog.value
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ScrlFog_Change", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlFogOpacity_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblFogOpacity.Caption = "Fog Opacity: " & scrlFogOpacity.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ScrlFogOpacity_Change", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub ScrlFogSpeed_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblFogSpeed.Caption = "Fog Speed: " & ScrlFogSpeed.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ScrlFogSpeed_Change", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub ScrlG_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblG.Caption = "Green: " & ScrlG.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ScrlG_Change", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub ScrlR_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblR.Caption = "Red: " & ScrlR.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ScrlR_Change", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlWeatherIntensity_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblWeatherIntensity.Caption = "Intensity: " & scrlWeatherIntensity.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ScrlWeatherIntensity_Change", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
