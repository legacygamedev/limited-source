VERSION 5.00
Begin VB.Form frmMapProperties 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Map Properties"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   8445
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
   ScaleHeight     =   4545
   ScaleWidth      =   8445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   270
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   270
      Left            =   2400
      TabIndex        =   6
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   270
      Left            =   3240
      TabIndex        =   33
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test"
      Height          =   270
      Left            =   2160
      TabIndex        =   32
      Top             =   2760
      Width           =   1095
   End
   Begin VB.ComboBox cmbMusic 
      Height          =   390
      Left            =   960
      TabIndex        =   31
      Text            =   "cmbMusic"
      Top             =   2280
      Width           =   3135
   End
   Begin VB.ComboBox cmbNpc 
      Height          =   390
      Index           =   7
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   3840
      Width           =   4095
   End
   Begin VB.ComboBox cmbNpc 
      Height          =   390
      Index           =   6
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   3360
      Width           =   4095
   End
   Begin VB.ComboBox cmbNpc 
      Height          =   390
      Index           =   5
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   2880
      Width           =   4095
   End
   Begin VB.ComboBox cmbShop 
      Height          =   390
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   4035
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox txtBootY 
      Alignment       =   1  'Right Justify
      Height          =   390
      Left            =   3360
      TabIndex        =   24
      Text            =   "0"
      Top             =   3240
      Width           =   735
   End
   Begin VB.TextBox txtBootX 
      Alignment       =   1  'Right Justify
      Height          =   390
      Left            =   1320
      TabIndex        =   23
      Text            =   "0"
      Top             =   3240
      Width           =   735
   End
   Begin VB.TextBox txtBootMap 
      Alignment       =   1  'Right Justify
      Height          =   390
      Left            =   1320
      TabIndex        =   20
      Text            =   "0"
      Top             =   2760
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
      ItemData        =   "frmMapProperties.frx":0000
      Left            =   4200
      List            =   "frmMapProperties.frx":0002
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
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   480
      Width           =   4095
   End
   Begin VB.ComboBox cmbMoral 
      Height          =   390
      ItemData        =   "frmMapProperties.frx":0004
      Left            =   960
      List            =   "frmMapProperties.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   1800
      Width           =   2415
   End
   Begin VB.TextBox txtLeft 
      Alignment       =   1  'Right Justify
      Height          =   390
      Left            =   2640
      TabIndex        =   3
      Text            =   "0"
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox txtRight 
      Alignment       =   1  'Right Justify
      Height          =   390
      Left            =   2640
      TabIndex        =   4
      Text            =   "0"
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox txtDown 
      Alignment       =   1  'Right Justify
      Height          =   390
      Left            =   960
      TabIndex        =   2
      Text            =   "0"
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox txtUp 
      Alignment       =   1  'Right Justify
      Height          =   390
      Left            =   960
      TabIndex        =   1
      Text            =   "0"
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox txtName 
      Height          =   390
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "NPC's"
      Height          =   375
      Left            =   4200
      TabIndex        =   26
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label10 
      Caption         =   "Music"
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "Boot Y"
      Height          =   375
      Left            =   2160
      TabIndex        =   22
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Boot X"
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Boot Map"
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Moral"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Right"
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Left"
      Height          =   375
      Left            =   1920
      TabIndex        =   10
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Down"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Up"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   120
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
Call DirectMusic.StopMusic
End Sub

Private Sub cmdTest_Click()
Call DirectMusic.StopMusic
Call DirectMusic.PlayMusic(cmbMusic.List(cmbMusic.ListIndex) & ".mid")
End Sub

Private Sub Form_Load()
Dim X As Long, Y As Long, i As Long, MusicName As String
Dim FSys As Object, Folder As Object, FolderFiles As Object, File As Object, FileName As String
Set FSys = CreateObject("Scripting.FileSystemObject")

    txtName.Text = Trim$(Map.Name)
    txtUp.Text = STR(Map.Up)
    txtDown.Text = STR(Map.Down)
    txtLeft.Text = STR(Map.Left)
    txtRight.Text = STR(Map.Right)
    cmbMoral.ListIndex = Map.Moral
    txtBootMap.Text = STR(Map.BootMap)
    txtBootX.Text = STR(Map.BootX)
    txtBootY.Text = STR(Map.BootY)
    
    cmbShop.AddItem "No Shop"
    For X = 1 To MAX_SHOPS
        cmbShop.AddItem X & ": " & Trim$(Shop(X).Name)
    Next X
    cmbShop.ListIndex = Map.Shop
    
    For X = 1 To MAX_MAP_NPCS
        cmbNpc(X - 1).AddItem "No NPC"
    Next X
    
    For Y = 1 To MAX_NPCS
        For X = 1 To MAX_MAP_NPCS
            cmbNpc(X - 1).AddItem Y & ": " & Trim$(Npc(Y).Name)
        Next X
    Next Y
    
    For i = 1 To MAX_MAP_NPCS
        cmbNpc(i - 1).ListIndex = Map.Npc(i)
    Next i
    
    'Set the folder objects
    Set Folder = FSys.GetFolder(App.Path & "\music")
    Set FolderFiles = Folder.Files
    
    cmbMusic.AddItem "No Music"
    For Each File In FolderFiles
        FileName = Mid(File, Len(App.Path & "\music\") + 1, ((Len(File) - Len(App.Path & "\music\"))))
        If UCase$(Right$(FileName, 3)) = "MP3" Or UCase$(Right$(FileName, 3)) = "MID" Then cmbMusic.AddItem FileName
    Next File
    
    'Destroy the folder objects
    Set File = Nothing
    Set FolderFiles = Nothing
    Set Folder = Nothing
    Set FSys = Nothing
    
    For X = 0 To cmbMusic.ListCount - 1
        If cmbMusic.List(X) = Map.Music Then
            cmbMusic.ListIndex = X
            Exit Sub
        End If
    Next X
    
    cmbMusic.ListIndex = 0
End Sub

Private Sub cmdOk_Click()
Dim X As Long, Y As Long, i As Long

    Map.Name = txtName.Text
    Map.Up = Val(txtUp.Text)
    Map.Down = Val(txtDown.Text)
    Map.Left = Val(txtLeft.Text)
    Map.Right = Val(txtRight.Text)
    Map.Moral = cmbMoral.ListIndex
    Map.Music = cmbMusic.List(cmbMusic.ListIndex)
    Map.BootMap = Val(txtBootMap.Text)
    Map.BootX = Val(txtBootX.Text)
    Map.BootY = Val(txtBootY.Text)
    Map.Shop = cmbShop.ListIndex
    
    For i = 1 To MAX_MAP_NPCS
        Map.Npc(i) = cmbNpc(i - 1).ListIndex
    Next i
    
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

