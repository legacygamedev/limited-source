Attribute VB_Name = "Menu"
Option Explicit

Public Sub MoveForm(F As Form)
    ReleaseCapture
    SendMessage F.hWnd, WM_NCLBUTTONDOWN, 2, 0
End Sub

Public Sub NewAccountConnect()
    Dim msg As String
    Dim i As Long

    If LenB(Trim$(frmMainMenu.txtNewAcctName.Text)) > 0 And LenB(Trim$(frmMainMenu.txtNewAcctPassword.Text)) > 0 Then
        msg = Trim$(frmMainMenu.txtNewAcctName.Text)

        ' Prevent high ascii chars
        For i = 1 To Len(msg)
            If Asc(Mid$(msg, i, 1)) < 32 Or Asc(Mid$(msg, i, 1)) > 126 Then
                Call MsgBox("You cannot use high ascii chars in your name, please re-enter.", vbOKOnly, GAME_NAME)
                frmMainMenu.txtNewAcctName.Text = vbNullString
                Exit Sub
            End If
        Next i

        Call MenuState(MENU_STATE_NEWACCOUNT)
    End If
End Sub

Public Sub LoginConnect()
    If LenB(Trim$(frmMainMenu.txtLoginName.Text)) > 0 And LenB(Trim$(frmMainMenu.txtLoginPassword.Text)) > 0 Then
        Call MenuState(MENU_STATE_LOGIN)
    ElseIf LenB(Trim$(frmMainMenu.txtLoginName.Text)) = 0 And LenB(Trim$(frmMainMenu.txtLoginPassword.Text)) = 0 Then
        Call MsgBox("Please enter your login name and password!", vbOKOnly)
    ElseIf LenB(Trim$(frmMainMenu.txtLoginName.Text)) = 0 Then
        Call MsgBox("Please enter your login name!", vbOKOnly)
    ElseIf LenB(Trim$(frmMainMenu.txtLoginPassword.Text)) = 0 Then
        Call MsgBox("Please enter your password!", vbOKOnly)
    End If
End Sub

Public Sub AddCharClick()
    Dim msg As String
    Dim i As Long

    If LenB(Trim$(frmMainMenu.txtNewCharName.Text)) > 0 Then
        msg = Trim$(frmMainMenu.txtNewCharName.Text)

        ' Prevent high ascii chars
        For i = 1 To Len(msg)
            If Asc(Mid$(msg, i, 1)) < 32 Or Asc(Mid$(msg, i, 1)) > 126 Then
                Call MsgBox("You cannot use high ascii chars in your name, please reenter.", vbOKOnly, GAME_NAME)
                frmMainMenu.txtNewCharName.Text = vbNullString
                Exit Sub
            End If
        Next i

        Call MenuState(MENU_STATE_ADDCHAR)
    End If
End Sub

Public Sub CloseSideMenu()
    ' Close Mirage Menus
    'With frmMainGame
    '    .picMnuGear.Visible = False
    '    .picPlayerSpells.Visible = False
    '    .picInv.Visible = False
    '    .picKeepNotes.Visible = False
    '    .picMnuTrain.Visible = False
    '    .picLiveStats.Visible = False
    'End With
End Sub

Public Sub MenuState(ByVal State As Long)

    frmSendGetData.Visible = True
    Call SetStatus("Connecting to server...")
    Select Case State
        Case MENU_STATE_NEWACCOUNT
            frmMainMenu.mnuNewAccount.Visible = False
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending new account information...")
                Call SendNewAccount(frmMainMenu.txtNewAcctName.Text, frmMainMenu.txtNewAcctPassword.Text)
            End If
            Exit Sub

        Case MENU_STATE_LOGIN
            frmMainMenu.mnuLogin.Visible = False
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending login information...")
                Call SendLogin(frmMainMenu.txtLoginName.Text, frmMainMenu.txtLoginPassword.Text)
            End If
            Exit Sub

        Case MENU_STATE_NEWCHAR
            frmMainMenu.mnuChars.Visible = False
            Call SetStatus("Connected, getting available classes...")
            Call SendGetClasses
            Exit Sub

        Case MENU_STATE_ADDCHAR
            frmMainMenu.mnuNewCharacter.Visible = False
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending character addition data...")
                If frmMainMenu.optMale.Value = True Then
                    Call SendAddChar(frmMainMenu.txtNewCharName, 0, frmMainMenu.cmbClass.ListIndex, frmMainMenu.lstChars.ListIndex + 1)
                Else
                    Call SendAddChar(frmMainMenu.txtNewCharName, 1, frmMainMenu.cmbClass.ListIndex, frmMainMenu.lstChars.ListIndex + 1)
                End If
            End If
            Exit Sub

        Case MENU_STATE_USECHAR
            frmMainMenu.Visible = False
            frmMainMenu.mnuChars.Visible = False
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending char data...")
                Call SendUseChar(frmMainMenu.lstChars.ListIndex + 1)
                Call Unload(frmMainMenu)
            End If
            Exit Sub
            
        Case MENU_STATE_DELCHAR
            If ConnectToServer Then
                Call SetStatus("Connected, sending character deletion request...")
                Call SendDelChar(frmMainMenu.lstChars.ListIndex + 1)
                frmMainMenu.mnuChars.Visible = False
            End If
            Exit Sub
            
    End Select

    If Not IsConnected Then
        frmMainMenu.Visible = True
        frmSendGetData.Visible = False
        Call MsgBox("Sorry, the server seems to be down.  Please try to reconnect in a few minutes or visit " & GAME_WEBSITE, vbOKOnly, GAME_NAME)
    End If
End Sub

