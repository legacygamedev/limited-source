Attribute VB_Name = "modGeneral"
Option Explicit

' ------------------------------------------
' --               Asphodel               --
' ------------------------------------------

Private Sub Main()

    frmStatus.Picture = LoadPicture(App.Path & GFX_PATH & "interface\statuswindow.bmp")
    
    SetStatus "Loading config..."
    
    Randomize
    
    TOTAL_SPRITES = -1
    TOTAL_ANIMGFX = -1
    
    vbQuote = ChrW$(34)
    
    GettingMap = True
    
    Load_Config
    
    Build_Lookups
    
    SetStatus "Initializing TCP..."
    
    TcpInit
    
    Load_GameConfig
    Load frmMainGame
    
    CheckTiles
    CheckItems
    
    SetStatus "Initializing DirectX 7..."
    
    ' load up DX7
    InitDirectDraw
    InitSurfaces
    
    If Sound_On Then InitDirectSound
    If Music_On Then InitDirectMusic
    
    ' late binding
    If DX7 Is Nothing Then Set DX7 = New DirectX7
    
    SetStatus "Starting menus..."
    
    Load frmMainMenu
    
    CurrentWindow = Window_State.Main_Menu
    LoadWindows
    
    frmMainMenu.Caption = Game_Name & " [Main Menu]"
    frmMainMenu.Visible = True
    frmStatus.Visible = False
    
End Sub

Public Sub CheckItems()
Dim i As Long

    While FileExist(GFX_PATH & "Items\item" & i & GFX_EXT)
        MAX_ITEMSETS = MAX_ITEMSETS + 1
        i = i + 1
    Wend
    
    MAX_ITEMSETS = MAX_ITEMSETS - 1
    
    ReDim DDS_Item(0 To MAX_ITEMSETS)
    ReDim DDSD_Item(0 To MAX_ITEMSETS)
    
End Sub

Public Sub CheckTiles()
Dim i As Long

    While FileExist(GFX_PATH & "tiles" & i & GFX_EXT)
        MAX_TILESETS = MAX_TILESETS + 1
        i = i + 1
    Wend
    
    MAX_TILESETS = MAX_TILESETS - 1
    
    ReDim DDS_Tile(0 To MAX_TILESETS)
    ReDim DDSD_Tile(0 To MAX_TILESETS)
    ReDim TILESHEET_WIDTH(0 To MAX_TILESETS)
    
    frmMainGame.scrlTileSet.Max = MAX_TILESETS
    
    ' getting the width of all tile sheets
    With frmMainGame.picCheckSize
       For i = 0 To MAX_TILESETS
          .Picture = LoadPicture(App.Path & GFX_PATH & "tiles" & i & GFX_EXT)
          LoadTileSheetWidth i, .Width
          .Picture = LoadPicture()
          .Width = 1
          .Height = 1
       Next
    End With
    
End Sub

Public Sub LoadTileSheetWidth(ByVal TileNum As Long, ByVal UseWidth As Long)
Dim LoopI As Long

    TILESHEET_WIDTH(TileNum) = (UseWidth \ PIC_X)
    
End Sub

Public Sub MenuState(ByVal State As Long)

    SetStatus "Connecting to server..."
    
    Select Case State
        Case Menu_State.NewAccount_
        
            frmMainMenu.Visible = False
            
            If ConnectToServer Then
                SetStatus "Connected, sending account information..."
                SendNewAccount frmMainMenu.txtUsername.Text, frmMainMenu.txtPassword.Text
            End If
            
        Case Menu_State.Login_
        
            frmMainMenu.Visible = False
            
            If ConnectToServer Then
                SetStatus "Connected, sending login information..."
                SendLogin frmMainMenu.txtUsername.Text, frmMainMenu.txtPassword.Text
                Exit Sub
            End If
            
        Case Menu_State.NewChar_
        
            frmChars.Visible = False
            
            If ConnectToServer Then
                SetStatus ("Connected, getting available classes...")
                SendGetClasses
            End If
            
        Case Menu_State.AddChar_
        
            frmNewChar.Visible = False
            
            If ConnectToServer Then
                SetStatus "Connected, sending character addition data..."
                If frmNewChar.optMale.Value Then
                    SendAddChar frmNewChar.txtName, GenderType.Male_, frmNewChar.cmbClass.ListIndex + 1, Char_Selected
                Else
                    SendAddChar frmNewChar.txtName, GenderType.Female_, frmNewChar.cmbClass.ListIndex + 1, Char_Selected
                End If
            End If
            
        Case Menu_State.DelChar_
        
            frmChars.Visible = False
            
            If ConnectToServer Then
                SetStatus "Connected, sending character deletion request..."
                SendDelChar Char_Selected
            End If
            
        Case Menu_State.UseChar_
        
            frmChars.Visible = False
            
            If ConnectToServer Then
                SetStatus "Connected, sending char data..."
                SendUseChar Char_Selected
            End If
            
    End Select
    
    If Not IsConnected Then
        frmMainMenu.Visible = True
        frmStatus.Visible = False
        MsgBox "Sorry, the server seems to be down.  Please try to reconnect in a few minutes or visit " & GAME_WEBSITE, vbOKOnly, Game_Name
    End If
    
End Sub

Public Sub GameInit()

    Unload frmMainMenu
    
    CurrentWindow = Window_State.Main_Game
    
    frmMainGame.picPlayerSpells.Picture = LoadPicture(App.Path & GFX_PATH & "Interface/miscwindow.bmp")
    frmMainGame.picInv.Picture = LoadPicture(App.Path & GFX_PATH & "Interface/miscwindow.bmp")
    frmMainGame.picStatWindow.Picture = LoadPicture(App.Path & GFX_PATH & "Interface/miscwindow.bmp")
    frmMainGame.picGuildCP.Picture = LoadPicture(App.Path & GFX_PATH & "Interface/miscwindow.bmp")
    frmMainGame.picOptions.Picture = LoadPicture(App.Path & GFX_PATH & "Interface/miscwindow.bmp")
    frmMainGame.picShop.Picture = LoadPicture(App.Path & GFX_PATH & "Interface/shopwindow.bmp")
    frmMainGame.Picture = LoadPicture(App.Path & GFX_PATH & "Interface/ingamewindow.bmp")
    frmMainGame.picTNL.Picture = LoadPicture(App.Path & GFX_PATH & "Interface/tnl_over.bmp")
    frmMainGame.picHP.Picture = LoadPicture(App.Path & GFX_PATH & "Interface/hpbar_over.bmp")
    frmMainGame.picMP.Picture = LoadPicture(App.Path & GFX_PATH & "Interface/mpbar_over.bmp")
    frmMainGame.picItemDesc.Picture = LoadPicture(App.Path & GFX_PATH & "Interface/itemdescwindow.bmp")
    frmMainGame.picItemDescBottom.Picture = LoadPicture(App.Path & GFX_PATH & "Interface/itemdescwindowbottom.bmp")
    frmMainGame.picSpellDesc.Picture = LoadPicture(App.Path & GFX_PATH & "Interface/spelldescwindow.bmp")
    frmMainGame.picSpellDescBottom.Picture = LoadPicture(App.Path & GFX_PATH & "Interface/spelldescwindowbottom.bmp")
    
    If Music_On Then
        frmMainGame.chkMusic.Value = 1
    Else
        frmMainGame.chkMusic.Value = 0
    End If
    
    If Sound_On Then
        frmMainGame.chkSound.Value = 1
    Else
        frmMainGame.chkSound.Value = 0
    End If
    
    If ShowPNames Then
        frmMainGame.chkPlayerNames.Value = 1
    Else
        frmMainGame.chkPlayerNames.Value = 0
    End If
    
    If ShowNNames Then
        frmMainGame.chkNPCNames.Value = 1
    Else
        frmMainGame.chkNPCNames.Value = 0
    End If
    
    frmMainGame.chkPing.Value = PingEnabled
    
    ' Set font
    SetFont FONT_NAME, FONT_SIZE
    
    frmStatus.Visible = False
    
    frmMainGame.Caption = Game_Name & " [In-Game]"
    
    frmMainGame.Show
    
    ' Set the focus
    SetFocusOnChat
    
End Sub

Public Sub DestroyGame()

    ' break out of GameLoop
    InGame = False
    
    DestroyTCP
    
    'destroy objects in reverse order
    DestroyDirectMusic
    DestroyDirectSound
    DestroyDirectDraw
    
    ' destroy DirectX7 master object
    If Not DX7 Is Nothing Then Set DX7 = Nothing
    
    UnloadAllForms
    End
    
End Sub

Public Sub UnloadAllForms()
Dim frm As Form

    For Each frm In VB.Forms
        Unload frm
    Next
    
End Sub

Public Sub SetStatus(ByVal Caption As String)

    If Not frmStatus.Visible Then frmStatus.Show
    
    frmStatus.lblStatus.Caption = Caption
    DoEvents
    
End Sub

Public Sub AddText(ByVal Msg As String, ByVal Color As Integer)
Dim s As String

    s = vbNewLine & Msg
    frmMainGame.txtChat.SelStart = Len(frmMainGame.txtChat.Text)
    frmMainGame.txtChat.SelColor = ColorTable(Color)
    frmMainGame.txtChat.SelText = s
    frmMainGame.txtChat.SelStart = Len(frmMainGame.txtChat.Text) - 1
    
    ' Prevent players from name spoofing
    frmMainGame.txtChat.SelHangingIndent = 10
    
End Sub

' Used for adding text to packet debugger
Public Sub TextAdd(ByVal Textbox As Textbox, ByVal Message As String)
    
    With Textbox
        .Text = .Text + Message + vbNewLine
        .SelStart = Len(.Text) - 1
    End With
    
End Sub

Public Sub SetFocusOnChat()
On Error Resume Next
' Prevent hardware related errors, no way to handle

    frmMainGame.txtMyChat.SetFocus
    
End Sub

Function Random(Lowerbound As Integer, Upperbound As Integer) As Integer
    Random = Int((Upperbound - Lowerbound + 1) * Rnd) + Lowerbound
End Function

Public Function isLoginLegal(ByVal Username As String, ByVal Password As String) As Boolean
    If LenB(Trim$(Username)) >= 3 Then
        If LenB(Trim$(Password)) >= 3 Then
            isLoginLegal = True
        End If
    End If
End Function

Public Function isStringLegal(ByVal sInput As String) As Boolean
Dim i As Long

    ' Prevent high ascii chars
    For i = 1 To Len(sInput)
        If Asc(Mid$(sInput, i, 1)) < vbKeySpace Or Asc(Mid$(sInput, i, 1)) > vbKeyF15 Then
            MsgBox "You cannot use high ASCII characters in your name, please re-enter.", vbOKOnly, Game_Name
            Exit Function
        End If
    Next
    
    isStringLegal = True
    
End Function

Public Sub Load_SpriteSizes()
Dim i As Long

    For i = 0 To TOTAL_SPRITES
        Sprite_Size(i).SizeX = DDSD_Sprite(i).lWidth \ ((Total_SpriteFrames) * 4)
        Sprite_Size(i).SizeY = DDSD_Sprite(i).lHeight
    Next
    
End Sub

Public Sub MovePicture(PB As PictureBox, Button As Integer, X As Single, Y As Single)

    If Button = 1 Then
        PB.Left = PB.Left + X - SOffsetX
        PB.Top = PB.Top + Y - SOffsetY
    End If
    
End Sub

Public Sub ResetWindows()

    With frmMainGame
        .picSpellDesc.Visible = False
        .picItemDesc.Visible = False
    End With
    
End Sub

Public Function GetTickCountNew() As Currency
    GetSysTimeMS GetTickCountNew
End Function
