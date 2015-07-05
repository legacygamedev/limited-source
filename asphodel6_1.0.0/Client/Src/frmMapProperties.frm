VERSION 5.00
Begin VB.Form frmMapProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Properties"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMusic 
      Caption         =   "Music"
      Height          =   1815
      Left            =   3120
      TabIndex        =   21
      Top             =   3480
      Width           =   4215
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1200
         Width           =   735
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton cmdPlay 
         Caption         =   "Play"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   735
      End
      Begin VB.FileListBox flMusic 
         Appearance      =   0  'Flat
         Height          =   990
         Left            =   960
         TabIndex        =   22
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label lblCurrentMusic 
         Alignment       =   2  'Center
         Caption         =   "Current Music: None"
         Height          =   255
         Left            =   960
         TabIndex        =   23
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Map Links"
      Height          =   1695
      Left            =   120
      TabIndex        =   15
      Top             =   600
      Width           =   2895
      Begin VB.TextBox txtUp 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1080
         TabIndex        =   19
         Text            =   "0"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtDown 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1080
         TabIndex        =   18
         Text            =   "0"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtRight 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1920
         TabIndex        =   17
         Text            =   "0"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtLeft 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   240
         TabIndex        =   16
         Text            =   "0"
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblMap 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame fraMapSettings 
      Caption         =   "Map Settings"
      Height          =   975
      Left            =   120
      TabIndex        =   12
      Top             =   4320
      Width           =   2895
      Begin VB.ComboBox cmbMoral 
         Height          =   360
         ItemData        =   "frmMapProperties.frx":0000
         Left            =   840
         List            =   "frmMapProperties.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Moral"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Boot Settings"
      Height          =   1815
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   2895
      Begin VB.TextBox txtBootMap 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1440
         TabIndex        =   8
         Text            =   "0"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtBootX 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1440
         TabIndex        =   7
         Text            =   "0"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtBootY 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1440
         TabIndex        =   6
         Text            =   "0"
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Boot Map"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Boot X"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Boot Y"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   1320
         Width           =   1095
      End
   End
   Begin VB.Frame fraNPCs 
      Caption         =   "NPCs"
      Height          =   2775
      Left            =   3120
      TabIndex        =   4
      Top             =   600
      Width           =   4215
      Begin VB.CommandButton cmdClearSpawn 
         Caption         =   "ClearSpawn"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   35
         Top             =   1920
         Width           =   735
      End
      Begin VB.CommandButton cmdSetspawn 
         Caption         =   "Set Spawn"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   34
         Top             =   1680
         Width           =   735
      End
      Begin VB.ListBox lstUseNpcs 
         Appearance      =   0  'Flat
         Height          =   1950
         Left            =   1920
         TabIndex        =   28
         Top             =   600
         Width           =   1335
      End
      Begin VB.ListBox lstNpcs 
         Appearance      =   0  'Flat
         Height          =   2190
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblMinus 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "[-]"
         Enabled         =   0   'False
         Height          =   240
         Left            =   2775
         TabIndex        =   37
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblPlus 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "[+]"
         Enabled         =   0   'False
         Height          =   240
         Left            =   2175
         TabIndex        =   36
         Top             =   360
         Width           =   285
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Spawn Data"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   33
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblSpawnY 
         Caption         =   "Y: None"
         Height          =   255
         Left            =   3360
         TabIndex        =   32
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblSpawnX 
         Caption         =   "X: None"
         Height          =   255
         Left            =   3360
         TabIndex        =   31
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblRemove 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "[<]"
         Height          =   240
         Left            =   1560
         TabIndex        =   30
         Top             =   1440
         Width           =   315
      End
      Begin VB.Label lblAdd 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "[>]"
         Height          =   240
         Left            =   1560
         TabIndex        =   29
         Top             =   1080
         Width           =   315
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   5400
      Width           =   2895
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   5400
      Width           =   2895
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "frmMapProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ------------------------------------------
' --               Asphodel               --
' ------------------------------------------

Private MapNpcUse() As Long

Private Sub cmdClearSpawn_Click()
    MapSpawnY(lstUseNpcs.ListIndex + 1) = -1
    MapSpawnX(lstUseNpcs.ListIndex + 1) = -1
    lstUseNpcs_Click
End Sub

Private Sub cmdSetspawn_Click()

    If MapNpcUse(lstUseNpcs.ListIndex + 1) = 0 Then Exit Sub
    
    SettingSpawn = True
    Me.Hide
    frmMainGame.Visible = True
    
    AddText "Click a spot on the map to choose where the NPC will spawn!", Color.BrightGreen
    
End Sub

Private Sub Form_Load()
Dim X As Long
Dim Y As Long
Dim i As Long

    ReDim Preserve MapSpawn.Npc(1 To 10)
    
    txtName.Text = Trim$(Map.Name)
    txtUp.Text = CStr(Map.Up)
    txtDown.Text = CStr(Map.Down)
    txtLeft.Text = CStr(Map.Left)
    txtRight.Text = CStr(Map.Right)
    cmbMoral.ListIndex = Map.Moral
    txtBootMap.Text = CStr(Map.BootMap)
    txtBootX.Text = CStr(Map.BootX)
    txtBootY.Text = CStr(Map.BootY)
    
    flMusic.Path = App.Path & MUSIC_PATH
    
    lstNpcs.Clear
    
    For i = 1 To MAX_NPCS
        lstNpcs.AddItem i & ": " & Trim$(Npc(i).Name)
    Next
    
    lstNpcs.ListIndex = 0
    
    lstUseNpcs.Clear
    
    ReDim MapNpcUse(1 To UBound(MapSpawn.Npc))
    ReDim MapSpawnX(1 To UBound(MapSpawn.Npc))
    ReDim MapSpawnY(1 To UBound(MapSpawn.Npc))
    
    For i = 1 To UBound(MapSpawn.Npc)
        MapSpawnX(i) = -1
        MapSpawnY(i) = -1
        If MapNpc(i).Num > 0 Then
            lstUseNpcs.AddItem i & ": " & Trim$(Npc(MapNpc(i).Num).Name)
            MapNpcUse(i) = MapNpc(i).Num
            MapSpawnY(i) = MapSpawn.Npc(i).Y
            MapSpawnX(i) = MapSpawn.Npc(i).X
        Else
            lstUseNpcs.AddItem i & ": (blank)"
            MapNpcUse(i) = 0
            MapSpawnY(i) = -1
            MapSpawnX(i) = -1
        End If
    Next
    
    lstUseNpcs.ListIndex = 0
    
    If LenB(Trim$(Map.Music)) > 0 Then
        lblCurrentMusic.Caption = "Current Music: " & Trim$(Map.Music) & MUSIC_EXT
        EditorMapMusic = Trim$(Map.Music)
        If flMusic.ListCount > 0 Then
            For i = 0 To flMusic.ListCount
                If flMusic.List(i) = Trim$(Map.Music) & MUSIC_EXT Then
                    flMusic.Selected(i) = True
                    flMusic.ListIndex = i
                End If
            Next
        End If
    Else
        lblCurrentMusic.Caption = "Current Music: None"
        EditorMapMusic = vbNullString
    End If
    
    lblMap.Caption = "Current map: " & GetPlayerMap(MyIndex)
    
End Sub

Private Sub cmdOk_Click()
Dim i As Long
Dim sTemp As Long

    With Map
        .Name = Trim$(txtName.Text)
        .Up = Val(txtUp.Text)
        .Down = Val(txtDown.Text)
        .Left = Val(txtLeft.Text)
        .Right = Val(txtRight.Text)
        .Moral = cmbMoral.ListIndex
        .Music = EditorMapMusic
        .BootMap = Val(txtBootMap.Text)
        .BootX = Val(txtBootX.Text)
        .BootY = Val(txtBootY.Text)
Rewind:
        ' get the high_npc_index
        For i = 1 To UBound(MapSpawn.Npc)
            If lstUseNpcs.List(i - 1) <> i & ": (blank)" Then
                sTemp = i
            End If
        Next
        
        For i = 1 To UBound(MapSpawn.Npc)
            If lstUseNpcs.List(i - 1) = i & ": (blank)" Then
                If i < sTemp Then
                    lstUseNpcs.List(i - 1) = i & ": " & Trim$(Npc(MapNpcUse(sTemp)).Name)
                    MapNpcUse(i) = MapNpcUse(sTemp)
                    lstUseNpcs.List(sTemp - 1) = sTemp & ": (blank)"
                    MapNpcUse(sTemp) = 0
                    GoTo Rewind
                End If
            End If
        Next
        
        For i = 1 To UBound(MapSpawn.Npc)
            MapSpawn.Npc(i).Num = MapNpcUse(i)
            
            If MapNpcUse(i) = 0 Then
                MapSpawn.Npc(i).X = -1
                MapSpawn.Npc(i).Y = -1
            Else
                MapSpawn.Npc(i).X = MapSpawnX(i)
                MapSpawn.Npc(i).Y = MapSpawnY(i)
            End If
        Next
        
    End With
    
    UpdateDrawMapName
    Unload Me
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub flMusic_Click()
Dim FileName() As String
Dim Ending As String

    If flMusic.ListIndex < 0 Then Exit Sub

    FileName = Split(flMusic.List(flMusic.ListIndex), ".", , vbTextCompare)
    
    If UBound(FileName) > 1 Then
        MsgBox "Invalid file name! Cannot contain any periods!"
        flMusic.ListIndex = -1
        Exit Sub
    End If
    
    Ending = FileName(1)
    
    If "." & Ending <> MUSIC_EXT Then
        MsgBox "." & UCase$(Ending) & " files are not supported. Please select another!"
        flMusic.ListIndex = -1
        Exit Sub
    End If
    
    lblCurrentMusic.Caption = "Current Music: " & flMusic.List(flMusic.ListIndex)
    EditorMapMusic = FileName(0)
    DirectMusic_PlayMidi flMusic.List(flMusic.ListIndex)

End Sub

Private Sub cmdPlay_Click()
    If LenB(EditorMapMusic) > 0 Then
        DirectMusic_PlayMidi EditorMapMusic & MUSIC_EXT
    End If
End Sub

Private Sub cmdStop_Click()
    DirectMusic_StopMidi
End Sub

Private Sub cmdClear_Click()
    DirectMusic_StopMidi
    EditorMapMusic = vbNullString
    flMusic.ListIndex = -1
    lblCurrentMusic.Caption = "Current Music: None"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReDim MapNpcUse(1 To 1)
    ReDim MapSpawnX(1 To 1)
    ReDim MapSpawnY(1 To 1)
    Me.Hide
    frmMainGame.Show
    Unload Me
End Sub

Private Sub lblAdd_Click()

    If LenB(Trim$(Npc(lstNpcs.ListIndex + 1).Name)) = 40 Then
        lblRemove_Click
    Else
        lstUseNpcs.List(lstUseNpcs.ListIndex) = lstUseNpcs.ListIndex + 1 & ": " & Trim$(Npc(lstNpcs.ListIndex + 1).Name)
        MapNpcUse(lstUseNpcs.ListIndex + 1) = lstNpcs.ListIndex + 1
    End If
    
    lstUseNpcs_Click
    
End Sub

Private Sub lblMinus_Click()
Dim i As Long

    If UBound(MapNpcUse) = 1 Then Exit Sub
    
    ReDim Preserve MapNpcUse(1 To UBound(MapNpcUse) - 1)
    ReDim Preserve MapSpawnX(1 To UBound(MapNpcUse))
    ReDim Preserve MapSpawnY(1 To UBound(MapNpcUse))
    ReDim Preserve MapSpawn.Npc(1 To UBound(MapNpcUse))
    'ReDim Preserve MapNpc(1 To UBound(MapNpcUse))
    
    lstUseNpcs.Clear
    
    For i = 1 To UBound(MapSpawn.Npc)
        If MapNpc(i).Num > 0 Then
            lstUseNpcs.AddItem i & ": " & Trim$(Npc(MapNpc(i).Num).Name)
            MapNpcUse(i) = MapNpc(i).Num
            MapSpawnY(i) = MapSpawn.Npc(i).Y
            MapSpawnX(i) = MapSpawn.Npc(i).X
        Else
            lstUseNpcs.AddItem i & ": (blank)"
            MapNpcUse(i) = 0
            MapSpawnY(i) = -1
            MapSpawnX(i) = -1
        End If
    Next
    
    lstUseNpcs.ListIndex = 0
    
End Sub

Private Sub lblPlus_Click()

    If UBound(MapNpcUse) = (MAX_MAPX * MAX_MAPY) - 1 Then
        MsgBox "You have reached the maximum number of NPCs per map! Good job!", , "Error"
        Exit Sub
    End If
    
    ReDim Preserve MapNpcUse(1 To UBound(MapNpcUse) + 1)
    ReDim Preserve MapSpawnX(1 To UBound(MapNpcUse))
    ReDim Preserve MapSpawnY(1 To UBound(MapNpcUse))
    ReDim Preserve MapSpawn.Npc(1 To UBound(MapNpcUse))
    'ReDim Preserve MapNpc(1 To UBound(MapNpcUse))
    
    MapSpawnX(UBound(MapNpcUse)) = -1
    MapSpawnY(UBound(MapNpcUse)) = -1
    
    lstUseNpcs.AddItem UBound(MapNpcUse) & ": (blank)"
    
End Sub

Private Sub lblRemove_Click()

    lstUseNpcs.List(lstUseNpcs.ListIndex) = lstUseNpcs.ListIndex + 1 & ": (blank)"
    
    MapNpcUse(lstUseNpcs.ListIndex + 1) = 0
    MapSpawnY(lstUseNpcs.ListIndex + 1) = -1
    MapSpawnX(lstUseNpcs.ListIndex + 1) = -1
    
    lstUseNpcs_Click
    
End Sub

Private Sub lstUseNpcs_Click()

    If MapSpawnY(lstUseNpcs.ListIndex + 1) = -1 Then
        lblSpawnY.Caption = "Y: None"
    Else
        lblSpawnY.Caption = "Y: " & MapSpawnY(lstUseNpcs.ListIndex + 1)
    End If
    
    If MapSpawnX(lstUseNpcs.ListIndex + 1) = -1 Then
        lblSpawnX.Caption = "X: None"
    Else
        lblSpawnX.Caption = "X: " & MapSpawnX(lstUseNpcs.ListIndex + 1)
    End If
    
End Sub
