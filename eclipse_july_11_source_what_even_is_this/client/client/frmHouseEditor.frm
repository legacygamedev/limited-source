VERSION 5.00
Begin VB.Form frmHouseEditor 
   Caption         =   "House Editor"
   ClientHeight    =   5355
   ClientLeft      =   165
   ClientTop       =   870
   ClientWidth     =   6975
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   357
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   465
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar scrlPicture 
      Height          =   5385
      LargeChange     =   10
      Left            =   0
      Max             =   512
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox picBack 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   5400
      Left            =   255
      ScaleHeight     =   360
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   448
      TabIndex        =   0
      Top             =   0
      Width           =   6720
      Begin VB.PictureBox picBackSelect 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   5400
         Left            =   0
         ScaleHeight     =   360
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   448
         TabIndex        =   1
         Top             =   0
         Width           =   6720
         Begin VB.Shape shpSelected 
            BorderColor     =   &H000000FF&
            Height          =   480
            Left            =   0
            Top             =   0
            Width           =   480
         End
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
      Begin VB.Menu mnuBreak2 
         Caption         =   "----------"
      End
      Begin VB.Menu mnuMinimize 
         Caption         =   "Minimze"
      End
   End
   Begin VB.Menu mnuTileSheet 
      Caption         =   "Tile Sheet"
      Begin VB.Menu mnuSet 
         Caption         =   "Tile Set 0"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuSet 
         Caption         =   "Tile Set 1"
         Index           =   1
      End
      Begin VB.Menu mnuSet 
         Caption         =   "Tile Set 2"
         Index           =   2
      End
      Begin VB.Menu mnuSet 
         Caption         =   "Tile Set 3"
         Index           =   3
      End
      Begin VB.Menu mnuSet 
         Caption         =   "Tile Set 4"
         Index           =   4
      End
      Begin VB.Menu mnuSet 
         Caption         =   "Tile Set 5"
         Index           =   5
      End
      Begin VB.Menu mnuSet 
         Caption         =   "Tile Set 6"
         Index           =   6
      End
   End
   Begin VB.Menu mnuTypes 
      Caption         =   "Select Type"
      Begin VB.Menu mnuType 
         Caption         =   "Layers"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu mnuType 
         Caption         =   "Blocked"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmHouseEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' // MAP EDITOR STUFF //
Dim KeyShift As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyShift Then
        KeyShift = True
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyShift = False
End Sub

Private Sub Form_Resize()
    If frmHouseEditor.WindowState = 0 Then
        If frmHouseEditor.Width > picBack.Width + scrlPicture.Width Then frmHouseEditor.Width = (picBack.Width + scrlPicture.Width + 8) * Screen.TwipsPerPixelX
        picBack.Height = (frmHouseEditor.Height - 800) / Screen.TwipsPerPixelX
        scrlPicture.Height = (frmHouseEditor.Height - 800) / Screen.TwipsPerPixelX
        frmHouseEditor.scrlPicture.Max = ((frmHouseEditor.picBackSelect.Height - frmHouseEditor.picBack.Height) / PIC_Y)
        If frmHouseEditor.Height > (picBackSelect.Height * Screen.TwipsPerPixelX) + 800 Then frmHouseEditor.Height = (picBackSelect.Height * Screen.TwipsPerPixelX) + 800
        
        frmAttributes.WindowState = 0
    End If
End Sub

Private Sub mnuDayNight_Click()
    If mnuDayNight.Checked = True Then
        mnuDayNight.Checked = False
    Else
        mnuDayNight.Checked = True
    End If
End Sub

Private Sub mnuExit_Click()
Dim x As Long

    x = MsgBox("Are you sure you want to discard your changes?", vbYesNo)
    If x = vbNo Then
        Exit Sub
    End If
    
    ScreenMode = 0
    Call HouseEditorCancel
End Sub

Private Sub mnuEyeDropper_Click()
    If frmHouseEditor.MousePointer = 2 Or frmMirage.MousePointer = 2 Then
        frmHouseEditor.MousePointer = 1
        frmMirage.MousePointer = 1
    Else
        frmHouseEditor.MousePointer = 2
        frmMirage.MousePointer = 2
    End If
End Sub

Private Sub mnuFill_Click()
Dim Y As Long
Dim x As Long

x = MsgBox("Are you sure you want to fill the map?", vbYesNo)
If x = vbNo Then
    Exit Sub
End If

If frmMapEditor.mnuType(1).Checked = True Then
    For Y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            With Map(GetPlayerMap(MyIndex)).Tile(x, Y)
                If frmAttributes.optGround.Value = True Then
                    .Ground = EditorTileY * TilesInSheets + EditorTileX
                    .GroundSet = EditorSet
                End If
                If frmAttributes.optMask.Value = True Then
                    .Mask = EditorTileY * TilesInSheets + EditorTileX
                    .MaskSet = EditorSet
                End If
                If frmAttributes.optAnim.Value = True Then
                    .Anim = EditorTileY * TilesInSheets + EditorTileX
                    .AnimSet = EditorSet
                End If
                If frmAttributes.optMask2.Value = True Then
                    .Mask2 = EditorTileY * TilesInSheets + EditorTileX
                    .Mask2Set = EditorSet
                End If
                If frmAttributes.optM2Anim.Value = True Then
                    .M2Anim = EditorTileY * TilesInSheets + EditorTileX
                    .M2AnimSet = EditorSet
                End If
                If frmAttributes.optFringe.Value = True Then
                    .Fringe = EditorTileY * TilesInSheets + EditorTileX
                    .FringeSet = EditorSet
                End If
                If frmAttributes.optFAnim.Value = True Then
                    .FAnim = EditorTileY * TilesInSheets + EditorTileX
                    .FAnimSet = EditorSet
                End If
                If frmAttributes.optFringe2.Value = True Then
                    .Fringe2 = EditorTileY * TilesInSheets + EditorTileX
                    .Fringe2Set = EditorSet
                End If
                If frmAttributes.optF2Anim.Value = True Then
                    .F2Anim = EditorTileY * TilesInSheets + EditorTileX
                    .F2AnimSet = EditorSet
                End If
            End With
        Next x
    Next Y
End Sub

Private Sub mnuMapGrid_Click()
    If mnuMapGrid.Checked = True Then
        WriteINI "CONFIG", "MapGrid", 0, App.Path & "\config.ini"
        mnuMapGrid.Checked = False
    Else
        WriteINI "CONFIG", "MapGrid", 1, App.Path & "\config.ini"
        mnuMapGrid.Checked = True
    End If
End Sub

Private Sub mnuMinimize_Click()
    frmMapEditor.WindowState = 1
    frmAttributes.WindowState = 1
End Sub


Private Sub mnuSave_Click()
Dim x As Long

    x = MsgBox("Are you sure you want to make these changes?", vbYesNo)
    If x = vbNo Then
        Exit Sub
    End If
    
    ScreenMode = 0
    Call HouseEditorSend
End Sub

Private Sub mnuSet_Click(Index As Integer)

    If mnuSet(Index).Checked = False Then
        mnuSet(Index).Checked = True
        picBackSelect.Picture = LoadPicture(App.Path & "\GFX\Tiles" & Index & ".bmp")
        EditorSet = Index
        
        scrlPicture.Max = ((picBackSelect.Height - picBack.Height) / PIC_Y)
        frmHouseEditor.picBack.Width = frmHouseEditor.picBackSelect.Width
        If frmHouseEditor.Width > picBack.Width + scrlPicture.Width Then frmHouseEditor.Width = (picBack.Width + scrlPicture.Width + 8) * Screen.TwipsPerPixelX
        If frmHouseEditor.Height > (picBackSelect.Height * Screen.TwipsPerPixelX) + 800 Then frmHouseEditor.Height = (picBackSelect.Height * Screen.TwipsPerPixelX) + 800
    End If
    
    Dim i As Byte
    For i = 0 To ExtraSheets
        If i <> Index Then mnuSet(i).Checked = False
    Next i
End Sub

Private Sub mnuType_Click(Index As Integer)
Dim i As Byte

    mnuType(Index).Checked = True
    If Index = 1 Then
        If mnuType(1).Checked = True Then
            frmAttributes.fraLayers.Visible = True
            frmAttributes.fraAttribs.Visible = False
            mnuTileSheet.Enabled = True
            frmAttributes.Visible = True
        End If
    End If
   If mnuType(2).Checked = True Then
            shpSelected.Width = 32
            shpSelected.Height = 32
            mnuTileSheet.Enabled = True
            frmAttributes.Visible = True
        End If
            For i = 1 To 2
        If i <> Index Then mnuType(i).Checked = False
    Next i
End Sub



Private Sub picBackSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyShift Then
        KeyShift = True
    End If
End Sub

Private Sub picBackSelect_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyShift = False
End Sub

Private Sub picBackSelect_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        If KeyShift = False Then
            Call HouseEditorChooseTile(Button, Shift, x, Y)
            shpSelected.Width = 32
            shpSelected.Height = 32
        Else
            EditorTileX = Int(x / PIC_X)
            EditorTileY = Int(Y / PIC_Y)
            
            If Int(EditorTileX * PIC_X) >= shpSelected.Left + shpSelected.Width Then
                EditorTileX = Int(EditorTileX * PIC_X + PIC_X) - (shpSelected.Left + shpSelected.Width)
                shpSelected.Width = shpSelected.Width + Int(EditorTileX)
            Else
                If shpSelected.Width > PIC_X Then
                    If Int(EditorTileX * PIC_X) >= shpSelected.Left Then
                        EditorTileX = (EditorTileX * PIC_X + PIC_X) - (shpSelected.Left + shpSelected.Width)
                        shpSelected.Width = shpSelected.Width + Int(EditorTileX)
                    End If
                End If
            End If
            
            If Int(EditorTileY * PIC_Y) >= shpSelected.Top + shpSelected.Height Then
                EditorTileY = Int(EditorTileY * PIC_Y + PIC_Y) - (shpSelected.Top + shpSelected.Height)
                shpSelected.Height = shpSelected.Height + Int(EditorTileY)
            Else
                If shpSelected.Height > PIC_Y Then
                    If Int(EditorTileY * PIC_Y) >= shpSelected.Top Then
                        EditorTileY = (EditorTileY * PIC_Y + PIC_Y) - (shpSelected.Top + shpSelected.Height)
                        shpSelected.Height = shpSelected.Height + Int(EditorTileY)
                    End If
                End If
            End If
        End If
    End If
    
    
    EditorTileX = Int((shpSelected.Left + PIC_X) / PIC_X)
    EditorTileY = Int((shpSelected.Top + PIC_Y) / PIC_Y)
End Sub

Private Sub picBackSelect_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        If KeyShift = False Then
            Call HouseEditorChooseTile(Button, Shift, x, Y)
            shpSelected.Width = 32
            shpSelected.Height = 32
        Else
            EditorTileX = Int(x / PIC_X)
            EditorTileY = Int(Y / PIC_Y)
            
            If Int(EditorTileX * PIC_X) >= shpSelected.Left + shpSelected.Width Then
                EditorTileX = Int(EditorTileX * PIC_X + PIC_X) - (shpSelected.Left + shpSelected.Width)
                shpSelected.Width = shpSelected.Width + Int(EditorTileX)
            Else
                If shpSelected.Width > PIC_X Then
                    If Int(EditorTileX * PIC_X) >= shpSelected.Left Then
                        EditorTileX = (EditorTileX * PIC_X + PIC_X) - (shpSelected.Left + shpSelected.Width)
                        shpSelected.Width = shpSelected.Width + Int(EditorTileX)
                    End If
                End If
            End If
            
            If Int(EditorTileY * PIC_Y) >= shpSelected.Top + shpSelected.Height Then
                EditorTileY = Int(EditorTileY * PIC_Y + PIC_Y) - (shpSelected.Top + shpSelected.Height)
                shpSelected.Height = shpSelected.Height + Int(EditorTileY)
            Else
                If shpSelected.Height > PIC_Y Then
                    If Int(EditorTileY * PIC_Y) >= shpSelected.Top Then
                        EditorTileY = (EditorTileY * PIC_Y + PIC_Y) - (shpSelected.Top + shpSelected.Height)
                        shpSelected.Height = shpSelected.Height + Int(EditorTileY)
                    End If
                End If
            End If
        End If
    End If
    
    
    EditorTileX = Int(shpSelected.Left / PIC_X)
    EditorTileY = Int(shpSelected.Top / PIC_Y)
End Sub

Private Sub scrlPicture_Change()
    Call HouseEditorTileScroll
End Sub


