VERSION 5.00
Begin VB.Form frmNpcEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Npc Editor"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10320
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
   ScaleHeight     =   6210
   ScaleWidth      =   10320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraReflection 
      Caption         =   "Reflection"
      Height          =   1455
      Left            =   5280
      TabIndex        =   56
      Top             =   2640
      Width           =   4815
      Begin VB.HScrollBar scrlReflection 
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   2880
         Max             =   300
         Min             =   1
         TabIndex        =   62
         Top             =   960
         Value           =   100
         Width           =   1455
      End
      Begin VB.HScrollBar scrlReflection 
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   360
         Max             =   300
         Min             =   1
         TabIndex        =   59
         Top             =   960
         Value           =   100
         Width           =   1455
      End
      Begin VB.CheckBox chkReflection 
         Caption         =   "Melee Reflection"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   58
         Top             =   360
         Width           =   1815
      End
      Begin VB.CheckBox chkReflection 
         Caption         =   "Magic Reflection"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   57
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblReflection 
         Alignment       =   2  'Center
         Caption         =   "100%"
         ForeColor       =   &H8000000A&
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   61
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblReflection 
         Alignment       =   2  'Center
         Caption         =   "100%"
         ForeColor       =   &H8000000A&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   60
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.Frame fraAttackSound 
      Caption         =   "Sounds"
      Height          =   1695
      Left            =   240
      TabIndex        =   50
      Top             =   4320
      Width           =   4815
      Begin VB.ListBox lstSoundTypes 
         Height          =   780
         ItemData        =   "frmNpcEditor.frx":0000
         Left            =   120
         List            =   "frmNpcEditor.frx":000D
         TabIndex        =   55
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   255
         Left            =   840
         TabIndex        =   54
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton cmdPlay 
         Caption         =   "Play"
         Height          =   255
         Left            =   840
         TabIndex        =   53
         Top             =   720
         Width           =   735
      End
      Begin VB.FileListBox flSound 
         Appearance      =   0  'Flat
         Height          =   990
         Left            =   1680
         TabIndex        =   51
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label lblCurrentSound 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Attack Sound: None"
         Height          =   255
         Left            =   1680
         TabIndex        =   52
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame fraItemDrop 
      Caption         =   "Item Drop"
      Height          =   1575
      Left            =   5280
      TabIndex        =   32
      Top             =   4080
      Width           =   4815
      Begin VB.HScrollBar scrlNum 
         Height          =   255
         Left            =   960
         Max             =   255
         TabIndex        =   34
         Top             =   720
         Value           =   1
         Width           =   3135
      End
      Begin VB.HScrollBar scrlValue 
         Height          =   255
         Left            =   960
         Max             =   255
         TabIndex        =   33
         Top             =   960
         Value           =   1
         Width           =   3135
      End
      Begin VB.HScrollBar scrlChance 
         Height          =   255
         Left            =   1440
         Max             =   100
         Min             =   1
         TabIndex        =   42
         Top             =   1200
         Value           =   1
         Width           =   2415
      End
      Begin VB.Label lblChance 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "100%"
         Height          =   240
         Left            =   4080
         TabIndex        =   43
         Top             =   1200
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Drop Chance"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   1200
         UseMnemonic     =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Value"
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   960
         UseMnemonic     =   0   'False
         Width           =   735
      End
      Begin VB.Label lblValue 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   375
         Left            =   4080
         TabIndex        =   35
         Top             =   960
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label lblNum 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   375
         Left            =   4080
         TabIndex        =   40
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "Num"
         Height          =   375
         Left            =   120
         TabIndex        =   39
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Item:"
         Height          =   375
         Left            =   120
         TabIndex        =   38
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblItemName 
         Caption         =   "None"
         Height          =   375
         Left            =   840
         TabIndex        =   37
         Top             =   360
         Width           =   3735
      End
   End
   Begin VB.Frame fraStats 
      Caption         =   "Stats"
      Height          =   1935
      Left            =   240
      TabIndex        =   15
      Top             =   2280
      Width           =   4815
      Begin VB.HScrollBar scrlExp 
         Height          =   255
         Left            =   1080
         TabIndex        =   46
         Top             =   480
         Width           =   3135
      End
      Begin VB.HScrollBar scrlHP 
         Height          =   255
         Left            =   1080
         Min             =   1
         TabIndex        =   45
         Top             =   240
         Value           =   1
         Width           =   3135
      End
      Begin VB.HScrollBar scrlStrength 
         Height          =   255
         Left            =   1080
         Max             =   255
         TabIndex        =   22
         Top             =   720
         Width           =   3135
      End
      Begin VB.HScrollBar scrlDefense 
         Height          =   255
         Left            =   1080
         Max             =   255
         TabIndex        =   21
         Top             =   960
         Width           =   3135
      End
      Begin VB.HScrollBar scrlSpeed 
         Enabled         =   0   'False
         Height          =   255
         Left            =   1080
         Max             =   255
         TabIndex        =   20
         Top             =   1200
         Width           =   3135
      End
      Begin VB.HScrollBar scrlMagic 
         Enabled         =   0   'False
         Height          =   255
         Left            =   1080
         Max             =   255
         TabIndex        =   19
         Top             =   1440
         Width           =   3135
      End
      Begin VB.Label Label12 
         Caption         =   "Magic"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Speed"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   1200
         UseMnemonic     =   0   'False
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Defense"
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   960
         UseMnemonic     =   0   'False
         Width           =   855
      End
      Begin VB.Label lblMagic 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   375
         Left            =   4200
         TabIndex        =   24
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lblSpeed 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   375
         Left            =   4200
         TabIndex        =   26
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label lblDefense 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   375
         Left            =   4200
         TabIndex        =   28
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblStrength 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   375
         Left            =   4200
         TabIndex        =   30
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Strength"
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   720
         UseMnemonic     =   0   'False
         Width           =   855
      End
      Begin VB.Label lblExp 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   375
         Left            =   4200
         TabIndex        =   47
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblHP 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   375
         Left            =   4200
         TabIndex        =   48
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label18 
         Caption         =   "Exp"
         Height          =   375
         Left            =   120
         TabIndex        =   49
         Top             =   480
         UseMnemonic     =   0   'False
         Width           =   855
      End
      Begin VB.Label Label17 
         Caption         =   "HP"
         Height          =   375
         Left            =   120
         TabIndex        =   44
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   855
      End
   End
   Begin VB.Frame FraNpcView 
      Caption         =   "Npc View"
      Height          =   2535
      Left            =   5280
      TabIndex        =   13
      Top             =   120
      Width           =   4815
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Left            =   1200
         Max             =   255
         TabIndex        =   16
         Top             =   240
         Width           =   2895
      End
      Begin VB.PictureBox picSprite 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   240
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   14
         Top             =   600
         Width           =   480
      End
      Begin VB.Label lblSprite 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   375
         Left            =   4080
         TabIndex        =   18
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Sprite"
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox txtSpawnSecs 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   3120
      TabIndex        =   12
      Text            =   "0"
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox txtAttackSay 
      Height          =   345
      Left            =   960
      TabIndex        =   10
      Top             =   480
      Width           =   4095
   End
   Begin VB.PictureBox picSprites 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   120
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   8
      Top             =   9360
      Width           =   480
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7800
      TabIndex        =   7
      Top             =   5640
      Width           =   2295
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   5280
      TabIndex        =   6
      Top             =   5640
      Width           =   2295
   End
   Begin VB.HScrollBar scrlRange 
      Height          =   255
      Left            =   2520
      Max             =   255
      TabIndex        =   3
      Top             =   1800
      Value           =   1
      Width           =   1935
   End
   Begin VB.ComboBox cmbBehavior 
      Height          =   360
      ItemData        =   "frmNpcEditor.frx":0027
      Left            =   1320
      List            =   "frmNpcEditor.frx":003A
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   840
      Width           =   3735
   End
   Begin VB.TextBox txtName 
      Height          =   360
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
   Begin VB.CheckBox chkGivesGuild 
      Caption         =   "Gives Guild"
      Height          =   255
      Left            =   240
      TabIndex        =   63
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Sight"
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   1800
      UseMnemonic     =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   375
      Left            =   240
      TabIndex        =   31
      Top             =   150
      Width           =   495
   End
   Begin VB.Label Label16 
      Caption         =   "Spawn Rate (in seconds):"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   1320
      UseMnemonic     =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label14 
      Caption         =   "Say"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   480
      UseMnemonic     =   0   'False
      Width           =   735
   End
   Begin VB.Label lblRange 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Behavior"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   840
      UseMnemonic     =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "frmNpcEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ------------------------------------------
' --               Asphodel               --
' ------------------------------------------

Private Sub chkGivesGuild_Click()
    If chkGivesGuild.Value = vbChecked Then
        cmbBehavior.ListIndex = 2
        cmbBehavior.Enabled = False
    Else
        cmbBehavior.Enabled = True
    End If
End Sub

Private Sub chkReflection_Click(Index As Integer)

    scrlReflection(Index).Enabled = chkReflection(Index).Value
    
    If scrlReflection(Index).Enabled Then
        lblReflection(Index).ForeColor = &H80000012
    Else
        lblReflection(Index).ForeColor = &H8000000A
    End If
    
End Sub

Private Sub lstSoundTypes_Click()
Dim i As Long

    If lstSoundTypes.ListIndex < 0 Then Exit Sub
    
    Select Case lstSoundTypes.ListIndex
        Case NpcSound.Attack_
            If LenB(EditorNpcAttackSound) > 0 Then
                lblCurrentSound.Caption = "Attack Sound: " & EditorNpcAttackSound & SOUND_EXT
                If flSound.ListCount > 0 Then
                    For i = 0 To flSound.ListCount
                        If flSound.List(i) = EditorNpcAttackSound & SOUND_EXT Then
                            flSound.Selected(i) = True
                            flSound.ListIndex = i
                        End If
                    Next
                End If
            Else
                lblCurrentSound.Caption = "Attack Sound: None"
                flSound.ListIndex = -1
            End If
        Case NpcSound.Spawn_
            If LenB(EditorNpcSpawnSound) > 0 Then
                lblCurrentSound.Caption = "On Spawn Sound: " & EditorNpcSpawnSound & SOUND_EXT
                If flSound.ListCount > 0 Then
                    For i = 0 To flSound.ListCount
                        If flSound.List(i) = EditorNpcSpawnSound & SOUND_EXT Then
                            flSound.Selected(i) = True
                            flSound.ListIndex = i
                        End If
                    Next
                End If
            Else
                lblCurrentSound.Caption = "On Spawn Sound: None"
                flSound.ListIndex = -1
            End If
        Case NpcSound.Death_
            If LenB(EditorNpcDeathSound) > 0 Then
                lblCurrentSound.Caption = "On Death Sound: " & EditorNpcDeathSound & SOUND_EXT
                If flSound.ListCount > 0 Then
                    For i = 0 To flSound.ListCount
                        If flSound.List(i) = EditorNpcSpawnSound & SOUND_EXT Then
                            flSound.Selected(i) = True
                            flSound.ListIndex = i
                        End If
                    Next
                End If
            Else
                lblCurrentSound.Caption = "On Death Sound: None"
                flSound.ListIndex = -1
            End If
    End Select
    
End Sub

Private Sub scrlChance_Change()
    lblChance.Caption = Round((1 / scrlChance.Value) * 100, 1) & "%"
End Sub

Private Sub scrlChance_Scroll()
    scrlChance_Change
End Sub

Private Sub scrlDefense_Scroll()
    scrlDefense_Change
End Sub

Private Sub scrlExp_Change()
    lblExp.Caption = CStr(scrlExp.Value)
End Sub

Private Sub scrlExp_Scroll()
    scrlExp_Change
End Sub

Private Sub scrlHP_Change()
    lblHP.Caption = CStr(scrlHP.Value)
End Sub

Private Sub scrlHP_Scroll()
    scrlHP_Change
End Sub

Private Sub scrlMagic_Scroll()
    scrlMagic_Change
End Sub

Private Sub scrlNum_Scroll()
    scrlNum_Change
End Sub

Private Sub scrlRange_Scroll()
    scrlRange_Change
End Sub

Private Sub scrlReflection_Change(Index As Integer)
    lblReflection(Index).Caption = scrlReflection(Index).Value & "%"
End Sub

Private Sub scrlReflection_Scroll(Index As Integer)
    scrlReflection_Change Index
End Sub

Private Sub scrlSpeed_Scroll()
    scrlSpeed_Change
End Sub

Private Sub scrlSprite_Change()
    lblSprite.Caption = CStr(scrlSprite.Value)
    
    NpcEditorBltSprite
End Sub

Private Sub scrlRange_Change()
    lblRange.Caption = CStr(scrlRange.Value)
End Sub

Private Sub scrlSprite_Scroll()
    scrlSprite_Change
End Sub

Private Sub scrlStrength_Change()
    lblStrength.Caption = CStr(scrlStrength.Value)
End Sub

Private Sub scrlDefense_Change()
    lblDefense.Caption = CStr(scrlDefense.Value)
End Sub

Private Sub scrlSpeed_Change()
    lblSpeed.Caption = CStr(scrlSpeed.Value)
End Sub

Private Sub scrlMagic_Change()
    lblMagic.Caption = CStr(scrlMagic.Value)
End Sub

Private Sub scrlNum_Change()
    lblNum.Caption = CStr(scrlNum.Value)
    
    If scrlNum.Value > 0 Then
        lblItemName.Caption = Trim$(Item(scrlNum.Value).Name)
    Else
        lblItemName.Caption = "None"
    End If
    
End Sub

Private Sub scrlStrength_Scroll()
    scrlStrength_Change
End Sub

Private Sub scrlValue_Change()
    lblValue.Caption = CStr(scrlValue.Value)
End Sub

Private Sub flSound_Click()
Dim FileName() As String
Dim Ending As String

    If lstSoundTypes.ListIndex < 0 Then Exit Sub
    
    If flSound.ListIndex < 0 Then Exit Sub
    
    FileName = Split(flSound.List(flSound.ListIndex), ".", , vbTextCompare)
    
    If UBound(FileName) > 1 Then
        MsgBox "Invalid file name! Cannot contain any periods!"
        flSound.ListIndex = -1
        Exit Sub
    End If
    
    Ending = FileName(1)
    
    If "." & Ending <> SOUND_EXT Then
        MsgBox "." & UCase$(Ending) & " files are not supported. Please select another!"
        flSound.ListIndex = -1
        Exit Sub
    End If
    
    Select Case lstSoundTypes.ListIndex
    
        Case 0
            lblCurrentSound.Caption = "Attack Sound: " & flSound.List(flSound.ListIndex)
            EditorNpcAttackSound = FileName(0)
        
        Case 1
            lblCurrentSound.Caption = "On Spawn Sound: " & flSound.List(flSound.ListIndex)
            EditorNpcSpawnSound = FileName(0)
        Case 2
            lblCurrentSound.Caption = "On Death Sound: " & flSound.List(flSound.ListIndex)
            EditorNpcDeathSound = FileName(0)
        
    End Select
    
    SoundPlay flSound.List(flSound.ListIndex)

End Sub

Private Sub cmdPlay_Click()
    If lstSoundTypes.ListIndex < 0 Then Exit Sub
    Select Case lstSoundTypes.ListIndex
        Case NpcSound.Attack_
            If LenB(EditorNpcAttackSound) > 0 Then
                SoundPlay EditorNpcAttackSound & SOUND_EXT
            End If
        Case NpcSound.Spawn_
            If LenB(EditorNpcSpawnSound) > 0 Then
                SoundPlay EditorNpcSpawnSound & SOUND_EXT
            End If
        Case NpcSound.Death_
            If LenB(EditorNpcDeathSound) > 0 Then
                SoundPlay EditorNpcDeathSound & SOUND_EXT
            End If
    End Select
End Sub

Private Sub cmdClear_Click()
    If lstSoundTypes.ListIndex < 0 Then Exit Sub
    Select Case lstSoundTypes.ListIndex
        Case NpcSound.Attack_
            EditorNpcAttackSound = vbNullString
            flSound.ListIndex = -1
            lblCurrentSound.Caption = "Attack Sound: None"
        Case NpcSound.Spawn_
            EditorNpcSpawnSound = vbNullString
            flSound.ListIndex = -1
            lblCurrentSound.Caption = "On Spawn Sound: None"
        Case NpcSound.Death_
            EditorNpcDeathSound = vbNullString
            flSound.ListIndex = -1
            lblCurrentSound.Caption = "On Death Sound: None"
    End Select
End Sub

Private Sub cmdOk_Click()
    If LenB(Trim$(txtName)) = 0 Then
        MsgBox ("Name required.")
    Else
        NpcEditorOk
    End If
End Sub

Private Sub cmdCancel_Click()
    NpcEditorCancel
End Sub

Private Sub scrlValue_Scroll()
    scrlValue_Change
End Sub
