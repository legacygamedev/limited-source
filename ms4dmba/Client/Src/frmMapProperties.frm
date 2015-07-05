VERSION 5.00
Begin VB.Form frmMapProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Properties"
   ClientHeight    =   6150
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
   ScaleHeight     =   410
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   498
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmMaxSizes 
      Caption         =   "Max Sizes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   31
      Top             =   4560
      Width           =   1935
      Begin VB.TextBox txtMaxX 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   33
         Text            =   "0"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtMaxY 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   32
         Text            =   "0"
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Max X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   270
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Max Y"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   630
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Map Links"
      Height          =   1695
      Left            =   120
      TabIndex        =   24
      Top             =   600
      Width           =   2895
      Begin VB.TextBox txtUp 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1080
         TabIndex        =   28
         Text            =   "0"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtDown 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1080
         TabIndex        =   27
         Text            =   "0"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtRight 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1920
         TabIndex        =   26
         Text            =   "0"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtLeft 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   240
         TabIndex        =   25
         Text            =   "0"
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblMap 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame fraMapSettings 
      Caption         =   "Map Settings"
      Height          =   2055
      Left            =   3120
      TabIndex        =   16
      Top             =   3480
      Width           =   4215
      Begin VB.ComboBox cmbMoral 
         Height          =   360
         ItemData        =   "frmMapProperties.frx":0000
         Left            =   960
         List            =   "frmMapProperties.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   360
         Width           =   2415
      End
      Begin VB.HScrollBar scrlMusic 
         Height          =   255
         Left            =   960
         Max             =   255
         TabIndex        =   18
         Top             =   1440
         Value           =   1
         Width           =   2415
      End
      Begin VB.ComboBox cmbShop 
         Height          =   360
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label6 
         Caption         =   "Moral"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Music"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lblMusic 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   375
         Left            =   3480
         TabIndex        =   21
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label12 
         Caption         =   "Shop"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Boot Settings"
      Height          =   2055
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   2895
      Begin VB.TextBox txtBootMap 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1440
         TabIndex        =   12
         Text            =   "0"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtBootX 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1440
         TabIndex        =   11
         Text            =   "0"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtBootY 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1440
         TabIndex        =   10
         Text            =   "0"
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Boot Map"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Boot X"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Boot Y"
         Height          =   255
         Left            =   360
         TabIndex        =   13
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
      Begin VB.ComboBox cmbNpc 
         Height          =   360
         Index           =   4
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1800
         Width           =   3975
      End
      Begin VB.ComboBox cmbNpc 
         Height          =   360
         Index           =   1
         ItemData        =   "frmMapProperties.frx":001F
         Left            =   120
         List            =   "frmMapProperties.frx":0021
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   3975
      End
      Begin VB.ComboBox cmbNpc 
         Height          =   360
         Index           =   2
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   840
         Width           =   3975
      End
      Begin VB.ComboBox cmbNpc 
         Height          =   360
         Index           =   3
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1320
         Width           =   3975
      End
      Begin VB.ComboBox cmbNpc 
         Height          =   360
         Index           =   5
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2280
         Width           =   3975
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   5640
      Width           =   2895
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   5640
      Width           =   2895
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   1080
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

' ******************************************
' **            Mirage Source 4           **
' ******************************************

Private Sub Form_Load()
Dim X As Long
Dim Y As Long
Dim i As Long

    txtName.Text = Trim$(Map.Name)
    txtUp.Text = CStr(Map.Up)
    txtDown.Text = CStr(Map.Down)
    txtLeft.Text = CStr(Map.Left)
    txtRight.Text = CStr(Map.Right)
    cmbMoral.ListIndex = Map.Moral
    scrlMusic.Value = Map.Music
    txtBootMap.Text = CStr(Map.BootMap)
    txtBootX.Text = CStr(Map.BootX)
    txtBootY.Text = CStr(Map.BootY)
    
    cmbShop.AddItem "No Shop"
    For X = 1 To MAX_SHOPS
        cmbShop.AddItem X & ": " & Trim$(Shop(X).Name)
    Next
    cmbShop.ListIndex = Map.Shop
    
    For X = 1 To MAX_MAP_NPCS
        cmbNpc(X).AddItem "No NPC"
    Next
    
    For Y = 1 To MAX_NPCS
        For X = 1 To MAX_MAP_NPCS
            cmbNpc(X).AddItem Y & ": " & Trim$(Npc(Y).Name)
        Next
    Next
    
    For i = 1 To MAX_MAP_NPCS
        cmbNpc(i).ListIndex = Map.Npc(i)
    Next
    
    lblMap.Caption = "Current map: " & GetPlayerMap(MyIndex)
    
    txtMaxX.Text = Map.MaxX
    txtMaxY.Text = Map.MaxY
End Sub

Private Sub scrlMusic_Change()
    lblMusic.Caption = CStr(scrlMusic.Value)
    'Call DirectMusic_PlayMidi(scrlMusic.Value)
End Sub

Private Sub cmdOk_Click()
Dim i As Long
Dim sTemp As Long
Dim X As Long, X2 As Long
Dim Y As Long, Y2 As Long
Dim tempArr() As TileRec

    If Not IsNumeric(txtMaxX.Text) Then txtMaxX.Text = Map.MaxX
    If Val(txtMaxX.Text) < MAX_MAPX Then txtMaxX.Text = MAX_MAPX
    If Val(txtMaxX.Text) > MAX_BYTE Then txtMaxX.Text = MAX_BYTE
    
    If Not IsNumeric(txtMaxY.Text) Then txtMaxY.Text = Map.MaxY
    If Val(txtMaxY.Text) < MAX_MAPY Then txtMaxY.Text = MAX_MAPY
    If Val(txtMaxY.Text) > MAX_BYTE Then txtMaxY.Text = MAX_BYTE
    
    With Map
        .Name = Trim$(txtName.Text)
        .Up = Val(txtUp.Text)
        .Down = Val(txtDown.Text)
        .Left = Val(txtLeft.Text)
        .Right = Val(txtRight.Text)
        .Moral = cmbMoral.ListIndex
        .Music = scrlMusic.Value
        .BootMap = Val(txtBootMap.Text)
        .BootX = Val(txtBootX.Text)
        .BootY = Val(txtBootY.Text)
        .Shop = cmbShop.ListIndex
        
        For i = 1 To MAX_MAP_NPCS
            If cmbNpc(i).ListIndex > 0 Then
                
                sTemp = InStr(1, Trim$(cmbNpc(i).Text), ":", vbTextCompare)
                
                If Len(Trim$(cmbNpc(i).Text)) = sTemp Then
                    cmbNpc(i).ListIndex = 0
                End If
            End If
        Next
        
        ' get the high_npc_index
        For i = 1 To MAX_MAP_NPCS
            If cmbNpc(i).ListIndex = 0 Then
                High_Npc_Index = i - 1
                Exit For
            End If
            High_Npc_Index = MAX_MAP_NPCS
        Next
            
        ' Check to see if there are spaces in between NPC slots
        If High_Npc_Index = 0 Then
            For i = 1 To MAX_MAP_NPCS
                If cmbNpc(i).ListIndex > 0 Then
                    MsgBox "Clear all NPC slots"
                    Exit Sub
                End If
            Next
        Else
            For i = High_Npc_Index To MAX_MAP_NPCS - 1
                If cmbNpc(i + 1).ListIndex > 0 Then
                    MsgBox "There cannot be empty spaces in between NPC slots"
                    Exit Sub
                End If
            Next
        End If
    
        For i = 1 To MAX_MAP_NPCS
            .Npc(i) = cmbNpc(i).ListIndex
        Next
        
        ' set the data before changing it
        tempArr = Map.Tile
        X2 = Map.MaxX
        Y2 = Map.MaxY
        
        ' change the data
        .MaxX = Val(txtMaxX.Text)
        .MaxY = Val(txtMaxY.Text)
        ReDim Map.Tile(0 To .MaxX, 0 To .MaxY)
        
        If X2 > .MaxX Then X2 = .MaxX
        If Y2 > .MaxY Then Y2 = .MaxY
        
        For X = 0 To X2
            For Y = 0 To Y2
                .Tile(X, Y) = tempArr(X, Y)
            Next
        Next
        
        ClearTempTile
    End With
    
    Call UpdateDrawMapName
    
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    
End Sub
