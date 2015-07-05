Attribute VB_Name = "General"
Option Explicit

' ******************************************
' **               rootSource               **
' ******************************************

' halts thread of execution
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

 ' get system uptime in milliseconds (32-bit)
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long

'For Clear functions
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Public Sub Main()

    frmSendGetData.Visible = True
    Call SetStatus("Loading...")
    
    Call SetStatus("Loading Game Data...")
    Call LoadDataFile
    
    GettingMap = True
    'vbQuote = ChrW$(34) ' "
    
    Load frmMainGame
    
        
    Set DX8 = New clsDX8
    
    VerProcess = 0
    
    If GameData.VerProcess = -1 Then
        If DX8.InitDirectX(frmMainGame.picScreen.hWnd, D3DCREATE_HARDWARE_VERTEXPROCESSING) = 0 Then
            VerProcess = 1
            If DX8.InitDirectX(frmMainGame.picScreen.hWnd, D3DCREATE_MIXED_VERTEXPROCESSING) = 0 Then
                VerProcess = 2
                If DX8.InitDirectX(frmMainGame.picScreen.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING) = 0 Then
                VerProcess = -1
                    MsgBox "Couldn't start DX8!"
                End If
            End If
        End If
    Else
        Select Case GameData.VerProcess
            Case 0
                If DX8.InitDirectX(frmMainGame.picScreen.hWnd, D3DCREATE_HARDWARE_VERTEXPROCESSING) = 0 Then
                    VerProcess = 1
                    If DX8.InitDirectX(frmMainGame.picScreen.hWnd, D3DCREATE_MIXED_VERTEXPROCESSING) = 0 Then
                        VerProcess = 2
                        If DX8.InitDirectX(frmMainGame.picScreen.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING) = 0 Then
                        VerProcess = -1
                            MsgBox "Couldn't start DX8!"
                        End If
                    End If
                End If
            Case 1
                VerProcess = 1
                If DX8.InitDirectX(frmMainGame.picScreen.hWnd, D3DCREATE_MIXED_VERTEXPROCESSING) = 0 Then
                    VerProcess = 0
                    If DX8.InitDirectX(frmMainGame.picScreen.hWnd, D3DCREATE_HARDWARE_VERTEXPROCESSING) = 0 Then
                        VerProcess = 2
                        If DX8.InitDirectX(frmMainGame.picScreen.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING) = 0 Then
                        VerProcess = -1
                            MsgBox "Couldn't start DX8!"
                        End If
                    End If
                End If
            Case 2
                VerProcess = 2
                If DX8.InitDirectX(frmMainGame.picScreen.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING) = 0 Then
                    VerProcess = 1
                    If DX8.InitDirectX(frmMainGame.picScreen.hWnd, D3DCREATE_MIXED_VERTEXPROCESSING) = 0 Then
                        VerProcess = 0
                        If DX8.InitDirectX(frmMainGame.picScreen.hWnd, D3DCREATE_HARDWARE_VERTEXPROCESSING) = 0 Then
                        VerProcess = -1
                            MsgBox "Couldn't start DX8!"
                        End If
                    End If
                End If
        End Select
    End If
    ' Update the form with the game's name
    frmMainGame.Caption = Trim$(GameData.GameName)
    frmMainGame.lblName.Caption = Trim$(GameData.GameName)

    ' randomize rnd's seed
    Randomize
    
    'Call InitFont

    Call SetStatus("Initializing TCP settings...")
    GAME_IP = Trim$(GameData.IP)
    GAME_PORT = GameData.Port
    Call TcpInit

    Call SetStatus("Initializing DirectX...")
    ' DX7 Master Object is already created, early binding
    Call CheckTiles
    Call CheckSprites
    Call CheckSpells
    Call CheckItems
        
   
    frmSendGetData.Visible = False

    Load frmMainMenu ' this line also initalizes directX
    
    InitSoundSys
    
    frmMainMenu.Visible = True
End Sub

'Public Sub MenuState(ByVal State As Long)
'    frmSendGetData.Visible = True
'
'    Select Case State
'        Case MENU_STATE_NEWACCOUNT
'            frmNewAccount.Visible = False
'            If ConnectToServer Then
'                Call SetStatus("Connected, sending new account information...")
'                Call SendNewAccount(frmNewAccount.txtName.Text, frmNewAccount.txtPassword.Text)
'            End If
'
'        Case MENU_STATE_DELACCOUNT
'            frmDeleteAccount.Visible = False
'            If ConnectToServer Then
'                Call SetStatus("Connected, sending account deletion request ...")
'                Call SendDelAccount(frmDeleteAccount.txtName.Text, frmDeleteAccount.txtPassword.Text)
'                Exit Sub
'            End If
'
'        Case MENU_STATE_LOGIN
'            frmLogin.Visible = False
'            If ConnectToServer Then
'                Call SetStatus("Connected, sending login information...")
'                Call SendLogin(frmLogin.txtName.Text, frmLogin.txtPassword.Text)
'                Exit Sub
'            End If
'
'        Case MENU_STATE_NEWCHAR
'            frmChars.Visible = False
'            Call SetStatus("Connected, getting available classes...")
'            Call SendGetClasses
'
'        Case MENU_STATE_ADDCHAR
'            frmNewChar.Visible = False
'            If ConnectToServer Then
'                Call SetStatus("Connected, sending character addition data...")
'                If frmNewChar.optMale.Value Then
'
'                    Call SendAddChar(frmNewChar.txtName, SEX_MALE, frmNewChar.cmbClass.ListIndex + 1, frmChars.lstChars.ListIndex + 1)
'                Else
'                    Call SendAddChar(frmNewChar.txtName, SEX_FEMALE, frmNewChar.cmbClass.ListIndex + 1, frmChars.lstChars.ListIndex + 1)
'                End If
'            End If
'
'        Case MENU_STATE_DELCHAR
'            frmChars.Visible = False
'            If ConnectToServer Then
'                Call SetStatus("Connected, sending character deletion request...")
'                Call SendDelChar(frmChars.lstChars.ListIndex + 1)
'            End If
'
'        Case MENU_STATE_USECHAR
'            frmChars.Visible = False
'            If ConnectToServer Then
'                Call SetStatus("Connected, sending char data...")
'                Call SendUseChar(frmChars.lstChars.ListIndex + 1)
'            End If
'    End Select
'
'    If frmSendGetData.Visible Then
'        If Not IsConnected Then
'            frmMainMenu.Visible = True
'            frmSendGetData.Visible = False
'            Call MsgBox("Sorry, the server seems to be down.  Please try to reconnect in a few minutes or visit " & GAME_WEBSITE, vbOKOnly, GAME_NAME)
'        End If
'    End If
'
'End Sub

Public Sub GameInit()
    Unload frmMainMenu

    frmSendGetData.Visible = False
    
    InitMapTables
    frmMainGame.Show
    frmMainGame.PreviewTimer.Enabled = True
    ' Set the focus
    Call SetFocusOnChat

End Sub

Public Sub DestroyGame()
    ' break out of GameLoop
    InGame = False
    
    KillSoundSys
    
    Call DestroyTCP

    
    Call UnloadAllForms
    End
End Sub

Public Sub UnloadAllForms()
    Dim frm As Form

    For Each frm In VB.Forms
        Unload frm
    Next
End Sub

Public Sub SetStatus(ByVal Caption As String)
    frmSendGetData.lblStatus.Caption = Caption
    'DoEvents
End Sub

Public Sub AddText(ByVal msg As String, ByVal Color As Integer)
Dim s As String
  
    s = vbNewLine & msg
    
    With frmMainGame.txtChat
        .SelStart = Len(.Text)
        .SelColor = QBColor(Color)
        .SelText = s
        
        .SelStart = Len(.Text) - 1
    
    
        ' Prevent players from name spoofing
        .SelHangingIndent = 15
    End With
    
End Sub

' Used for adding text to packet debugger
Public Sub TextAdd(ByVal Txt As TextBox, msg As String, NewLine As Boolean)
    If NewLine Then
        Txt.Text = Txt.Text + msg + vbCrLf
    Else
        Txt.Text = Txt.Text + msg
    End If
    
    Txt.SelStart = Len(Txt.Text) - 1
End Sub

Public Sub SetFocusOnChat()
On Error Resume Next 'prevent RTE5, no way to handle error
    frmMainGame.txtMyChat.SetFocus
End Sub

Public Function Rand(ByVal Low As Long, ByVal High As Long) As Long
  Rand = Int((High - Low + 1) * Rnd) + Low
End Function

Public Sub MovePicture(PB As PictureBox, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim GlobalX As Integer
    Dim GlobalY As Integer

    GlobalX = PB.Left
    GlobalY = PB.Top

    If Button = 1 Then
        PB.Left = GlobalX + x - SOffsetX
        PB.Top = GlobalY + y - SOffsetY
    End If
End Sub

Public Sub ResetUI()
    With frmMainGame
        .picInvList.Left = 0
        .picInvList.Top = 0
        
        .picSpellsList.Left = 0
        .picSpellsList.Top = 0
    End With
End Sub

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
            Call MsgBox("You cannot use high ASCII characters in your name, please re-enter.", vbOKOnly, GAME_NAME)
            Exit Function
        End If
    Next
    
    isStringLegal = True
        
End Function

