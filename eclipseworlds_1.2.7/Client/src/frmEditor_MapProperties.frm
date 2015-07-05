VERSION 5.00
Begin VB.Form frmEditor_MapProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Properties"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12105
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditor_MapProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   518
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   807
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame7 
      Caption         =   "Fog"
      Height          =   2055
      Left            =   2280
      TabIndex        =   28
      Top             =   3600
      Width           =   1815
      Begin VB.HScrollBar ScrlFogSpeed 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   31
         Top             =   1080
         Width           =   1455
      End
      Begin VB.HScrollBar ScrlFog 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   30
         Top             =   480
         Width           =   1455
      End
      Begin VB.HScrollBar scrlFogOpacity 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   29
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label lblFogOpacity 
         Caption         =   "Fog Opacity: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label lblFogSpeed 
         Caption         =   "Fog Speed: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label lblFog 
         Caption         =   "Fog: None"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Panorama"
      ForeColor       =   &H80000007&
      Height          =   5655
      Left            =   6600
      TabIndex        =   59
      Top             =   0
      Width           =   5415
      Begin VB.PictureBox picPanorama 
         Appearance      =   0  'Flat
         BackColor       =   &H80000008&
         ForeColor       =   &H80000008&
         Height          =   4695
         Left            =   120
         ScaleHeight     =   311
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   343
         TabIndex        =   61
         Top             =   240
         Width           =   5175
      End
      Begin VB.HScrollBar scrlPanorama 
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   5280
         Width           =   5175
      End
      Begin VB.Label lblPanorama 
         Caption         =   "Panorama: 0"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   5040
         Width           =   2295
      End
   End
   Begin VB.Frame frmMaxSizes 
      Caption         =   "Max Sizes"
      Height          =   1215
      Left            =   120
      TabIndex        =   54
      Top             =   3600
      Width           =   2055
      Begin VB.TextBox txtMaxX 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   56
         Text            =   "0"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtMaxY 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   55
         Text            =   "0"
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblMaxX 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Max X:"
         Height          =   195
         Left            =   120
         TabIndex        =   58
         Top             =   360
         Width           =   600
      End
      Begin VB.Label lblMaxY 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Max Y:"
         Height          =   195
         Left            =   120
         TabIndex        =   57
         Top             =   720
         Width           =   585
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   9480
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   10800
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      Caption         =   "Map Overlay"
      Height          =   2055
      Left            =   4200
      TabIndex        =   42
      Top             =   3600
      Width           =   2295
      Begin VB.HScrollBar ScrlB 
         Height          =   255
         Left            =   1200
         Max             =   255
         TabIndex        =   46
         Top             =   1080
         Width           =   975
      End
      Begin VB.HScrollBar ScrlG 
         Height          =   255
         Left            =   1200
         Max             =   255
         TabIndex        =   45
         Top             =   720
         Width           =   975
      End
      Begin VB.HScrollBar ScrlR 
         Height          =   255
         Left            =   1200
         Max             =   255
         TabIndex        =   44
         Top             =   360
         Width           =   975
      End
      Begin VB.HScrollBar scrlA 
         Height          =   255
         Left            =   1200
         Max             =   255
         TabIndex        =   43
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lblB 
         Caption         =   "Blue: 0"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblG 
         Caption         =   "Green: 0"
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblR 
         Caption         =   "Red: 0"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblA 
         Caption         =   "Alpha: 0"
         ForeColor       =   &H80000006&
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   1560
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Map Sound Effect"
      Height          =   855
      Left            =   120
      TabIndex        =   40
      Top             =   4800
      Width           =   2055
      Begin VB.ComboBox cmbSound 
         Height          =   315
         ItemData        =   "frmEditor_MapProperties.frx":038A
         Left            =   120
         List            =   "frmEditor_MapProperties.frx":038C
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Weather"
      Height          =   1575
      Left            =   120
      TabIndex        =   35
      Top             =   5640
      Width           =   2055
      Begin VB.ComboBox cmbWeather 
         Height          =   315
         ItemData        =   "frmEditor_MapProperties.frx":038E
         Left            =   120
         List            =   "frmEditor_MapProperties.frx":03A4
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   480
         Width           =   1815
      End
      Begin VB.HScrollBar scrlWeatherIntensity 
         Height          =   255
         Left            =   120
         Max             =   100
         TabIndex        =   36
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Weather Type:"
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label lblWeatherIntensity 
         Caption         =   "Intensity: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Music"
      Height          =   1575
      Left            =   2280
      TabIndex        =   23
      Top             =   5640
      Width           =   9735
      Begin VB.CheckBox chkAutoPlay 
         Caption         =   "Autoplay"
         Height          =   375
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   375
         Left            =   8640
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdPlay 
         Caption         =   "Play"
         Height          =   375
         Left            =   8640
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
      Begin VB.ListBox lstMusic 
         Height          =   1230
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   8415
      End
   End
   Begin VB.Frame fraNPCs 
      Caption         =   "NPCs"
      Height          =   2295
      Left            =   2280
      TabIndex        =   18
      Top             =   1320
      Width           =   4215
      Begin VB.CheckBox chkAutoSpawn 
         BackColor       =   &H80000004&
         Caption         =   "No Auto Spawn"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         MaskColor       =   &H00000000&
         TabIndex        =   53
         Top             =   1920
         Width           =   1815
      End
      Begin VB.ListBox lstNpcs 
         Height          =   1230
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   3975
      End
      Begin VB.CommandButton cmbClear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   3120
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton cmbUpdate 
         Caption         =   "Update"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1560
         Width           =   975
      End
      Begin VB.ComboBox cmbNpcs 
         Height          =   315
         ItemData        =   "frmEditor_MapProperties.frx":03D3
         Left            =   1200
         List            =   "frmEditor_MapProperties.frx":03D5
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1560
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Links"
      Height          =   1575
      Left            =   120
      TabIndex        =   16
      Top             =   480
      Width           =   2055
      Begin VB.TextBox txtUp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   720
         TabIndex        =   2
         Text            =   "0"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtDown 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   720
         TabIndex        =   5
         Text            =   "0"
         Top             =   1170
         Width           =   615
      End
      Begin VB.TextBox txtRight 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1320
         TabIndex        =   4
         Text            =   "0"
         Top             =   885
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   120
         TabIndex        =   3
         Text            =   "0"
         Top             =   885
         Width           =   615
      End
      Begin VB.Label lblMap 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Current Map: 0"
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
         Left            =   0
         TabIndex        =   17
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame fraMapSettings 
      Caption         =   "Settings"
      Height          =   855
      Left            =   2280
      TabIndex        =   14
      Top             =   480
      Width           =   4215
      Begin VB.ComboBox cmbMoral 
         Height          =   315
         ItemData        =   "frmEditor_MapProperties.frx":03D7
         Left            =   960
         List            =   "frmEditor_MapProperties.frx":03DE
         Style           =   2  'Dropdown List
         TabIndex        =   9
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
      Caption         =   "Bootmap"
      Height          =   1575
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   2055
      Begin VB.TextBox txtBootMap 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   840
         TabIndex        =   6
         Text            =   "0"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtBootX 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   840
         TabIndex        =   7
         Text            =   "0"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtBootY 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   840
         TabIndex        =   8
         Text            =   "0"
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Map:"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   420
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "X:"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   195
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Y:"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   180
      End
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
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "frmEditor_MapProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAutoPlay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lstMusic.SetFocus
End Sub

Private Sub chkAutoSpawn_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If lstNpcs.ListIndex > -1 Then
        Map.NPCSpawnType(lstNpcs.ListIndex + 1) = chkAutoSpawn.Value
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "chkAutoSpawn_Click", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmbSound_Change()
    If cmbSound.ListIndex < 0 Then Exit Sub
    Audio.StopSounds
    Audio.PlaySound Map.BGS, -1, -1, True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    frmMain.cmdSave.Enabled = True
    frmMain.cmdRevert.Enabled = True
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Form_Unload", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmbClear_Click()
    Dim I As Long
    Dim TmpIndex As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Exit if the list Index is subscript out of range
    If lstNpcs.ListIndex + 1 < 1 Or lstNpcs.ListIndex + 1 > MAX_MAP_NPCS Then Exit Sub
    
    ' Clear the NPCs from the list
    For I = 1 To MAX_MAP_NPCS
        Map.NPC(I) = 0
    Next
    
    TmpIndex = lstNpcs.ListIndex
        
    ' Clear the list
    lstNpcs.Clear
    
    ' Reload the list NPCs
    Call LoadMapPropertiesNPCs
    
    ' Set the list Index to TmpIndex
    lstNpcs.ListIndex = TmpIndex
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmbClear_Click", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmbUpdate_Click()
    Dim TmpIndex As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Exit if the list index is subscript out of range
    If lstNpcs.ListIndex < 0 Or lstNpcs.ListIndex > MAX_MAP_NPCS - 1 Then Exit Sub
    
    ' Exit early if we don't need to move it down
    If lstNpcs.ListIndex = MAX_MAP_NPCS - 1 Then
        ' Save the npc to the map
        Map.NPC(lstNpcs.ListIndex + 1) = cmbNpcs.ListIndex
        
        ' Clear the list
        lstNpcs.Clear
        
        ' Reload the list NPCs
        Call LoadMapPropertiesNPCs
        lstNpcs.ListIndex = MAX_MAP_NPCS - 1
        Exit Sub
    End If
    
    ' Make sure it has a name
    If Not cmbNpcs.ListIndex = 0 Then
        If Trim$(NPC(cmbNpcs.ListIndex).Name) = vbNullString Then Exit Sub
    End If
    
    Map.NPC_HighIndex = lstNpcs.ListIndex + 1
    
    ' Set the temporary index for when it reloads it after it rebuilds the list
    TmpIndex = lstNpcs.ListIndex
    
    ' Save the npc to the map
    Map.NPC(lstNpcs.ListIndex + 1) = cmbNpcs.ListIndex
    
    ' Clear the list
    lstNpcs.Clear
    
    ' Reload the npcs into the list
    Call LoadMapPropertiesNPCs
    
    ' Set the new index based on the old temporary Index
    lstNpcs.ListIndex = TmpIndex + 1
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmbUpdate_Click", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdPlay_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call Audio.StopMusic
    Call Audio.PlayMusic(lstMusic.List(lstMusic.ListIndex), True)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdPlay_Click", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdStop_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call Audio.StopMusic
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdStop_Click", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lstMusic_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If chkAutoPlay.Value = 1 Then
        Call Audio.StopMusic
        
        ' Don't play none
        If lstMusic.List(lstMusic.ListIndex) = vbNullString Then Exit Sub
        
        Call Audio.PlayMusic(lstMusic.List(lstMusic.ListIndex), True)
    End If
    Exit Sub

' Error handler
ErrorHandler:
    HandleError "lstMusic_DblClick", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lstMusic_DblClick()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call Audio.StopMusic
    Call Audio.PlayMusic(lstMusic.List(lstMusic.ListIndex), True)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "lstMusic_DblClick", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Set max values
    txtName.MaxLength = NAME_LENGTH
    scrlPanorama.max = NumPanoramas
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Form_Load", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub cmdOK_Click()
    Dim I As Long
    Dim sTemp As Long
    Dim X As Long, X2 As Long
    Dim Y As Long, Y2 As Long
    Dim TempArr() As TileRec
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Not IsNumeric(txtMaxX.text) Then txtMaxX.text = Map.MaxX
    If val(txtMaxX.text) < MIN_MAPX Then txtMaxX.text = MIN_MAPX
    If val(txtMaxX.text) > MAX_BYTE Then txtMaxX.text = MAX_BYTE
    If Not IsNumeric(txtMaxY.text) Then txtMaxY.text = Map.MaxY
    If val(txtMaxY.text) < MIN_MAPY Then txtMaxY.text = MIN_MAPY
    If val(txtMaxY.text) > MAX_BYTE Then txtMaxY.text = MAX_BYTE
    
    With Map
        ' Set values
        .Name = Trim$(txtName.text)
        
        ' Save music
        If lstMusic.ListIndex > 0 Then
            .Music = lstMusic.List(lstMusic.ListIndex)
        Else
            .Music = vbNullString
        End If
        
        ' Save BGS
        If cmbSound.ListIndex >= 0 Then
            Audio.StopSounds
            .BGS = cmbSound.List(cmbSound.ListIndex)
            Audio.PlaySound cmbSound.List(cmbSound.ListIndex), -1, -1, True
        Else
            .BGS = vbNullString
        End If
        
        ' Other things to save
        .Up = val(txtUp.text)
        .Down = val(txtDown.text)
        .Left = val(txtLeft.text)
        .Right = val(txtRight.text)
        .Moral = cmbMoral.ListIndex + 1

        .Weather = cmbWeather.ListIndex
        .WeatherIntensity = scrlWeatherIntensity.Value
        
        .Fog = ScrlFog.Value
        .FogSpeed = ScrlFogSpeed.Value
        .FogOpacity = scrlFogOpacity.Value
        
        .Panorama = scrlPanorama.Value
        
        .Red = ScrlR.Value
        .Green = ScrlG.Value
        .Blue = ScrlB.Value
        .Alpha = scrlA.Value

        TempArr = Map.Tile
        X2 = Map.MaxX
        Y2 = Map.MaxY
        
        ' Set the data before changing it
        .MaxX = val(txtMaxX.text)
        .MaxY = val(txtMaxY.text)
        
        ReDim Map.Tile(0 To .MaxX, 0 To .MaxY)

        If X2 > .MaxX Then X2 = .MaxX
        If Y2 > .MaxY Then Y2 = .MaxY

        For X = 0 To X2
            For Y = 0 To Y2
                .Tile(X, Y) = TempArr(X, Y)
            Next
        Next
        
        ' Set Bootmap
        .BootMap = val(txtBootMap.text)
        .BootX = val(txtBootX.text)
        .BootY = val(txtBootY.text)
    End With
    
    ' Update the map Name
    Call UpdateDrawMapName
    InitAutotiles
    Unload frmEditor_MapProperties
    
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdOK_Click", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Unload frmEditor_MapProperties
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdCancel_Click", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lstNPCs_dblClick()
    Dim I As Long
    Dim Index As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Index = frmEditor_MapProperties.lstNpcs.ListIndex + 1
    
    If Index < 1 Or Index > MAX_NPCS Then Exit Sub
    
    For I = 1 To MAX_NPCS
        If frmEditor_MapProperties.lstNpcs.List(Index - 1) = Index & ": None" Then
            cmbNpcs.ListIndex = 0
            Exit For
        ElseIf frmEditor_MapProperties.lstNpcs.List(Index - 1) = Index & ": " & Trim$(NPC(I).Name) Then
            cmbNpcs.ListIndex = I
            Exit For
        End If
    Next
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "lstNPCs_dblClick", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lstNPCs_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    chkAutoSpawn.Value = Map.NPCSpawnType(lstNpcs.ListIndex + 1)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "lstNPCs_Click", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlA_Change()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblA.Caption = "Alpha: " & scrlA.Value
    Exit Sub
     
' Error handler
ErrorHandler:
    HandleError "ScrlA_Change", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub ScrlB_Change()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblB.Caption = "Blue: " & ScrlB.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ScrlB_Change", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub ScrlFog_Change()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblFog.Caption = "Fog: " & ScrlFog.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ScrlFog_Change", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlFogOpacity_Change()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblFogOpacity.Caption = "Fog Opacity: " & scrlFogOpacity.Value
    Exit Sub
        
' Error handler
ErrorHandler:
    HandleError "ScrlFogOpacity_Change", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub ScrlFogSpeed_Change()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblFogSpeed.Caption = "Fog Speed: " & ScrlFogSpeed.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ScrlFogSpeed_Change", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub ScrlG_Change()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblG.Caption = "Green: " & ScrlG.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ScrlG_Change", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlPanorama_Change()
    lblPanorama.Caption = "Panorama: " & CStr(scrlPanorama.Value)
End Sub

Private Sub ScrlR_Change()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblR.Caption = "Red: " & ScrlR.Value
    Exit Sub

' Error handler
ErrorHandler:
    HandleError "ScrlR_Change", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlWeatherIntensity_Change()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblWeatherIntensity.Caption = "Intensity: " & scrlWeatherIntensity.Value
    Exit Sub
        
' Error handler
ErrorHandler:
    HandleError "ScrlWeatherIntensity_Change", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtBootMap_Change()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Not IsNumeric(txtBootMap.text) Then txtBootMap = 0
    If frmEditor_MapProperties.txtBootMap.text < 0 Then frmEditor_MapProperties.txtBootMap.text = 0
    If frmEditor_MapProperties.txtBootMap.text > MAX_MAPS Then frmEditor_MapProperties.txtBootMap.text = MAX_MAPS
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtBootMap_Change", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtBootX_Change()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Not IsNumeric(txtBootX.text) Then txtBootX = 0
    If frmEditor_MapProperties.txtBootX.text < 0 Then frmEditor_MapProperties.txtBootX.text = 0
    If frmEditor_MapProperties.txtBootX.text > MAX_BYTE Then frmEditor_MapProperties.txtBootX.text = MAX_BYTE
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtBootX_Change", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtBootY_Change()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Not IsNumeric(txtBootY.text) Then txtBootY = 0
    If frmEditor_MapProperties.txtBootY.text < 0 Then frmEditor_MapProperties.txtBootY.text = 0
    If frmEditor_MapProperties.txtBootY.text > MAX_BYTE Then frmEditor_MapProperties.txtBootY.text = MAX_BYTE
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtBootY_Change", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtDown_Change()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Not IsNumeric(txtDown.text) Then txtDown = 0
    If frmEditor_MapProperties.txtDown.text < 0 Then frmEditor_MapProperties.txtDown.text = 0
    If frmEditor_MapProperties.txtDown.text > MAX_MAPS Then frmEditor_MapProperties.txtDown.text = MAX_MAPS
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtDown_Change", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtLeft_Change()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Not IsNumeric(txtLeft.text) Then txtLeft = 0
    If frmEditor_MapProperties.txtLeft.text < 0 Then frmEditor_MapProperties.txtLeft.text = 0
    If frmEditor_MapProperties.txtLeft.text > MAX_MAPS Then frmEditor_MapProperties.txtLeft.text = MAX_MAPS
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtLeft_Change", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtName_GotFocus()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    txtName.SelStart = Len(txtName)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtName_GotFocus", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtRight_Change()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Not IsNumeric(txtRight.text) Then txtRight = 0
    If frmEditor_MapProperties.txtRight.text < 0 Then frmEditor_MapProperties.txtRight.text = 0
    If frmEditor_MapProperties.txtRight.text > MAX_MAPS Then frmEditor_MapProperties.txtRight.text = MAX_MAPS
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtRight_Change", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtUp_Change()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Not IsNumeric(txtUp.text) Then txtUp = 0
    If frmEditor_MapProperties.txtUp.text < 0 Then frmEditor_MapProperties.txtUp.text = 0
    If frmEditor_MapProperties.txtUp.text > MAX_MAPS Then frmEditor_MapProperties.txtUp.text = MAX_MAPS
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtUp_Change", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtUp_GotFocus()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    txtUp.SelStart = Len(txtUp)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtUp_GotFocus", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtDown_GotFocus()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    txtDown.SelStart = Len(txtDown)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtDown_GotFocus", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtLeft_GotFocus()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    txtLeft.SelStart = Len(txtLeft)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtLeft_GotFocus", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtRight_GotFocus()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    txtRight.SelStart = Len(txtRight)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtRight_GotFocus", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub
