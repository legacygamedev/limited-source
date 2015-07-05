VERSION 5.00
Begin VB.Form frmMapProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Properties"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11535
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   11535
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtStreet 
      Height          =   390
      Left            =   960
      TabIndex        =   56
      Top             =   360
      Width           =   3015
   End
   Begin VB.Frame frmBank 
      Caption         =   "Bank"
      Height          =   615
      Left            =   120
      TabIndex        =   52
      Top             =   2760
      Width           =   3975
      Begin VB.OptionButton optNoBank 
         Caption         =   "No"
         Height          =   270
         Left            =   1920
         TabIndex        =   54
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton optBank 
         Caption         =   "Yes"
         Height          =   270
         Left            =   360
         TabIndex        =   53
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.FileListBox fileMusic 
      Appearance      =   0  'Flat
      Height          =   5970
      Left            =   8400
      Pattern         =   "music*.*"
      TabIndex        =   48
      Top             =   75
      Width           =   3015
   End
   Begin VB.Frame frmNight 
      Caption         =   "Night Options"
      Height          =   1335
      Left            =   240
      TabIndex        =   46
      Top             =   5760
      Width           =   3735
      Begin VB.OptionButton optNight 
         Caption         =   "Night Only"
         Height          =   270
         Left            =   120
         TabIndex        =   51
         Top             =   1080
         Width           =   2775
      End
      Begin VB.OptionButton optDay 
         Caption         =   "Day Only"
         Height          =   270
         Left            =   120
         TabIndex        =   50
         Top             =   720
         Width           =   3255
      End
      Begin VB.OptionButton optDayNight 
         Caption         =   "Normal (Day and Night)"
         Height          =   270
         Left            =   120
         TabIndex        =   49
         Top             =   360
         Width           =   3495
      End
      Begin VB.HScrollBar scrlNight 
         Height          =   255
         Left            =   120
         Max             =   3
         TabIndex        =   47
         Top             =   1320
         Visible         =   0   'False
         Width           =   3495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Map Respawns ?"
      Height          =   855
      Left            =   240
      TabIndex        =   43
      Top             =   4905
      Width           =   3735
      Begin VB.OptionButton optnNoRespawn 
         Caption         =   "No"
         Height          =   375
         Left            =   1800
         TabIndex        =   45
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton optnYesRespawn 
         Caption         =   "Yes"
         Height          =   375
         Left            =   240
         TabIndex        =   44
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   375
      Left            =   10215
      TabIndex        =   41
      Top             =   6105
      Width           =   1215
   End
   Begin VB.CommandButton cmdTestMusic 
      Caption         =   "Test Music"
      Height          =   375
      Left            =   8415
      TabIndex        =   40
      Top             =   6105
      Width           =   1275
   End
   Begin VB.ComboBox cmbNpc 
      Height          =   390
      Index           =   13
      ItemData        =   "frmMapProperties.frx":0000
      Left            =   4200
      List            =   "frmMapProperties.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   39
      Top             =   6720
      Width           =   4095
   End
   Begin VB.ComboBox cmbNpc 
      Height          =   390
      Index           =   12
      ItemData        =   "frmMapProperties.frx":0004
      Left            =   4200
      List            =   "frmMapProperties.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   38
      Top             =   6240
      Width           =   4095
   End
   Begin VB.ComboBox cmbNpc 
      Height          =   390
      Index           =   11
      ItemData        =   "frmMapProperties.frx":0008
      Left            =   4200
      List            =   "frmMapProperties.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   37
      Top             =   5760
      Width           =   4095
   End
   Begin VB.ComboBox cmbNpc 
      Height          =   390
      Index           =   10
      ItemData        =   "frmMapProperties.frx":000C
      Left            =   4200
      List            =   "frmMapProperties.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   36
      Top             =   5280
      Width           =   4095
   End
   Begin VB.ComboBox cmbNpc 
      Height          =   390
      Index           =   9
      ItemData        =   "frmMapProperties.frx":0010
      Left            =   4200
      List            =   "frmMapProperties.frx":0012
      Style           =   2  'Dropdown List
      TabIndex        =   35
      Top             =   4800
      Width           =   4095
   End
   Begin VB.ComboBox cmbNpc 
      Height          =   390
      Index           =   8
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   34
      Top             =   4320
      Width           =   4095
   End
   Begin VB.ComboBox cmbNpc 
      Height          =   390
      Index           =   7
      ItemData        =   "frmMapProperties.frx":0014
      Left            =   4200
      List            =   "frmMapProperties.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   3840
      Width           =   4095
   End
   Begin VB.ComboBox cmbNpc 
      Height          =   390
      Index           =   6
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   32
      Top             =   3360
      Width           =   4095
   End
   Begin VB.ComboBox cmbNpc 
      Height          =   390
      Index           =   5
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   2880
      Width           =   4095
   End
   Begin VB.ComboBox cmbShop 
      Height          =   390
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   2400
      Width           =   2415
   End
   Begin VB.HScrollBar scrlMusic 
      Height          =   375
      Left            =   960
      Max             =   255
      TabIndex        =   26
      Top             =   3480
      Value           =   1
      Width           =   2415
   End
   Begin VB.TextBox txtBootY 
      Alignment       =   1  'Right Justify
      Height          =   390
      Left            =   3360
      TabIndex        =   24
      Text            =   "0"
      Top             =   4425
      Width           =   735
   End
   Begin VB.TextBox txtBootX 
      Alignment       =   1  'Right Justify
      Height          =   390
      Left            =   1320
      TabIndex        =   23
      Text            =   "0"
      Top             =   4425
      Width           =   735
   End
   Begin VB.TextBox txtBootMap 
      Alignment       =   1  'Right Justify
      Height          =   390
      Left            =   1320
      TabIndex        =   20
      Text            =   "0"
      Top             =   3945
      Width           =   735
   End
   Begin VB.ComboBox cmbNpc 
      Height          =   390
      Index           =   4
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   2400
      Width           =   4095
   End
   Begin VB.ComboBox cmbNpc 
      Height          =   390
      Index           =   3
      ItemData        =   "frmMapProperties.frx":0018
      Left            =   4200
      List            =   "frmMapProperties.frx":001A
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   1920
      Width           =   4095
   End
   Begin VB.ComboBox cmbNpc 
      Height          =   390
      Index           =   2
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   1440
      Width           =   4095
   End
   Begin VB.ComboBox cmbNpc 
      Height          =   390
      Index           =   1
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   960
      Width           =   4095
   End
   Begin VB.ComboBox cmbNpc 
      Height          =   390
      Index           =   0
      ItemData        =   "frmMapProperties.frx":001C
      Left            =   4200
      List            =   "frmMapProperties.frx":001E
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   480
      Width           =   4095
   End
   Begin VB.ComboBox cmbMoral 
      Height          =   390
      ItemData        =   "frmMapProperties.frx":0020
      Left            =   960
      List            =   "frmMapProperties.frx":0030
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   1920
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4200
      TabIndex        =   6
      Top             =   7200
      Width           =   3975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   7200
      Width           =   3975
   End
   Begin VB.TextBox txtLeft 
      Alignment       =   1  'Right Justify
      Height          =   390
      Left            =   2640
      TabIndex        =   3
      Text            =   "0"
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox txtRight 
      Alignment       =   1  'Right Justify
      Height          =   390
      Left            =   2640
      TabIndex        =   4
      Text            =   "0"
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox txtDown 
      Alignment       =   1  'Right Justify
      Height          =   390
      Left            =   960
      TabIndex        =   2
      Text            =   "0"
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox txtUp 
      Alignment       =   1  'Right Justify
      Height          =   390
      Left            =   960
      TabIndex        =   1
      Text            =   "0"
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox txtName 
      Height          =   390
      Left            =   960
      TabIndex        =   0
      Top             =   0
      Width           =   3015
   End
   Begin VB.Label Label14 
      Caption         =   "Street"
      Height          =   375
      Left            =   120
      TabIndex        =   55
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label13 
      Caption         =   "Music takes about 3 seconds to start."
      Height          =   600
      Left            =   8385
      TabIndex        =   42
      Top             =   6570
      Width           =   3105
   End
   Begin VB.Label Label12 
      Caption         =   "Shop"
      Height          =   375
      Left            =   120
      TabIndex        =   29
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "NPC's"
      Height          =   375
      Left            =   4200
      TabIndex        =   28
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label lblMusic 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   3480
      TabIndex        =   27
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label10 
      Caption         =   "Music"
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "Boot Y"
      Height          =   375
      Left            =   2520
      TabIndex        =   22
      Top             =   4425
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Boot X"
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   4425
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Boot Map"
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   3945
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Moral"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Right"
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Left"
      Height          =   375
      Left            =   1920
      TabIndex        =   10
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Down"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Up"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "frmMapProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit






Private Sub cmdStop_Click()
    Call StopMidi
End Sub

Private Sub cmdTestMusic_Click()
Dim filename As String
filename = fileMusic.filename
filename = Mid(filename, 6, Len(filename) - 9)
    Call PlayMidi("music" & filename)
End Sub

Private Sub fileMusic_Click()
Dim filename As String
filename = fileMusic.filename
filename = Mid(filename, 6, Len(filename) - 9)
scrlMusic.value = filename
End Sub

Private Sub Form_Load()
Dim x As Long, y As Long, i As Long

fileMusic.Path = App.Path & "\data\audio\music\"

    txtName.text = Trim(map.Name)
    txtStreet.text = Trim(map.street)
    txtUp.text = str(map.Up)
    txtDown.text = str(map.Down)
    txtLeft.text = str(map.Left)
    txtRight.text = str(map.Right)
    cmbMoral.ListIndex = map.Moral
    scrlMusic.value = map.music
    txtBootMap.text = str(map.BootMap)
    txtBootX.text = str(map.BootX)
    txtBootY.text = str(map.BootY)
    optBank.value = map.Bank
    optNoBank.value = Not optBank.value
    cmbShop.AddItem "No Shop"
    For x = 1 To MAX_SHOPS
        cmbShop.AddItem x & ": " & Trim(Shop(x).Name)
    Next x
    cmbShop.ListIndex = map.Shop
    
    For x = 1 To MAX_MAP_NPCS
        cmbNpc(x - 1).AddItem "No NPC"
    Next x
    
    For y = 1 To MAX_NPCS
        For x = 1 To MAX_MAP_NPCS
            cmbNpc(x - 1).AddItem y & ": " & Trim(Npc(y).Name)
        Next x
    Next y
    optnYesRespawn = map.Respawn
    optnNoRespawn = Not optnYesRespawn
    For i = 1 To MAX_MAP_NPCS
        cmbNpc(i - 1).ListIndex = map.Npc(i)
    Next i
    scrlNight.value = map.Night
    Call changeNightCap(map.Night)
End Sub





Private Sub optDay_Click()
scrlNight = 2
End Sub

Private Sub optDayNight_Click()
scrlNight = 0
End Sub

Private Sub optNight_Click()
scrlNight = 1
End Sub



Private Sub scrlMusic_Change()
    lblMusic.Caption = str(scrlMusic.value)
End Sub

Private Sub cmdOK_Click()
Dim x As Long, y As Long, i As Long
    
    map.Name = txtName.text
    map.street = txtStreet.text
    map.Up = Val(txtUp.text)
    map.Down = Val(txtDown.text)
    map.Left = Val(txtLeft.text)
    map.Right = Val(txtRight.text)
    map.Moral = cmbMoral.ListIndex
    map.music = scrlMusic.value
    map.BootMap = Val(txtBootMap.text)
    map.BootX = Val(txtBootX.text)
    map.BootY = Val(txtBootY.text)
    map.Shop = cmbShop.ListIndex
    map.Respawn = optnYesRespawn.value
    map.Night = scrlNight.value
    map.Bank = optBank.value
    
    'Map.Night = optNight.Item.value
    
    For i = 1 To MAX_MAP_NPCS
        map.Npc(i) = cmbNpc(i - 1).ListIndex
    Next i
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub scrlNight_Change()
    Call changeNightCap(scrlNight.value)
End Sub

Public Sub changeNightCap(ByVal value As Long)
Select Case value
    Case Is = 0
        optDayNight.value = True
    Case Is = 1
        optNight.value = True
    Case Is = 2
        optDay.value = True
End Select
End Sub
