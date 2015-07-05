Attribute VB_Name = "modGeneral"
Option Explicit

' Client Executes Here.
Public Sub Main()
    If FileExists("debug") Then
        frmDebug.Visible = True
    End If

    frmSendGetData.Visible = True

    ' Check to make sure all the folder exist.
    Call SetStatus("Checking Folders...")
    Call CheckFolders

    ' Check to make sure all the files exist.
    Call SetStatus("Checking Files...")
    Call SystemFileChecker

    If Not FileExists("config.ini") Then
        Call FileCreateConfigINI
    End If

    If Not FileExists("News.ini") Then
        Call FileCreateNewsINI
    End If

    If Not FileExists("Font.ini") Then
        Call FileCreateFontINI
    End If

    If Not FileExists("GUI\Colors.txt") Then
        Call FileCreateColorsTXT
    End If
    
    ' Initialize global variables.
    LAST_DIR = -1

    ' Load the configuration settings.
    Call SetStatus("Loading Configuration...")
    Call LoadConfig
    Call LoadColors
    Call LoadFont

    ' Prepare the socket for communication.
    Call SetStatus("Preparing Socket...")
    Call TcpInit

    frmMainMenu.lblVersion.Caption = "Version: " & App.Major & "." & App.Minor

    frmSendGetData.Visible = False
    frmMainMenu.Visible = True
End Sub

Private Sub CheckFolders()

    If LCase$(Dir$(App.Path & "\Maps", vbDirectory)) <> "maps" Then
        Call MkDir$(App.Path & "\Maps")
    End If

    If UCase$(Dir$(App.Path & "\GFX", vbDirectory)) <> "GFX" Then
        Call MkDir$(App.Path & "\GFX")
    End If

    If UCase$(Dir$(App.Path & "\GUI", vbDirectory)) <> "GUI" Then
        Call MkDir$(App.Path & "\GUI")
    End If

    If UCase$(Dir$(App.Path & "\Music", vbDirectory)) <> "MUSIC" Then
        Call MkDir$(App.Path & "\Music")
    End If

    If UCase$(Dir$(App.Path & "\SFX", vbDirectory)) <> "SFX" Then
        Call MkDir$(App.Path & "\SFX")
    End If

    If UCase$(Dir$(App.Path & "\Flashs", vbDirectory)) <> "FLASHS" Then
        Call MkDir$(App.Path & "\Flashs")
    End If

    If UCase$(Dir$(App.Path & "\BGS", vbDirectory)) <> "BGS" Then
        Call MkDir$(App.Path & "\BGS")
    End If

    If UCase$(Dir$(App.Path & "\DATA", vbDirectory)) <> "DATA" Then
        Call MkDir$(App.Path & "\Data")
    End If

End Sub

Private Sub LoadConfig()
    Dim FileName As String

    On Error GoTo ErrorHandle

    FileName = App.Path & "\config.ini"

    frmMirage.chkBubbleBar.Value = CLng(ReadINI("CONFIG", "SpeechBubbles", FileName))
    frmMirage.chkNpcBar.Value = CLng(ReadINI("CONFIG", "NpcBar", FileName))
    frmMirage.chkNpcName.Value = CLng(ReadINI("CONFIG", "NPCName", FileName))
    frmMirage.chkPlayerBar.Value = CLng(ReadINI("CONFIG", "PlayerBar", FileName))
    frmMirage.chkPlayerName.Value = CLng(ReadINI("CONFIG", "PlayerName", FileName))
    frmMirage.chkPlayerDamage.Value = CLng(ReadINI("CONFIG", "NPCDamage", FileName))
    frmMirage.chkNpcDamage.Value = ReadINI("CONFIG", "PlayerDamage", FileName)
   ' frmMirage.chkMusic.Value = CLng(ReadINI("CONFIG", "Music", FileName)) <-- This caused connectivity issues upon disabling music [Devil Of Duce]
    frmMirage.chkSound.Value = CLng(ReadINI("CONFIG", "Sound", FileName))
    frmMirage.chkAutoScroll.Value = CLng(ReadINI("CONFIG", "AutoScroll", FileName))
    AutoLogin = CLng(ReadINI("CONFIG", "Auto", FileName))

    Exit Sub

ErrorHandle:
    Call MsgBox("Error reading from config.ini, re-creating file.")
    Kill "config.ini"
    Call FileCreateConfigINI

End Sub

Private Sub FileCreateConfigINI()
    WriteINI "IPCONFIG", "IP", "127.0.0.1", App.Path & "\config.ini"
    WriteINI "IPCONFIG", "PORT", 4001, App.Path & "\config.ini"

    WriteINI "CONFIG", "Account", vbNullString, App.Path & "\config.ini"
    WriteINI "CONFIG", "Password", vbNullString, App.Path & "\config.ini"
    WriteINI "CONFIG", "WebSite", vbNullString, App.Path & "\config.ini"
    WriteINI "CONFIG", "SpeechBubbles", 1, App.Path & "\config.ini"
    WriteINI "CONFIG", "NpcBar", 1, App.Path & "\config.ini"
    WriteINI "CONFIG", "NPCName", 1, App.Path & "\config.ini"
    WriteINI "CONFIG", "NPCDamage", 1, App.Path & "\config.ini"
    WriteINI "CONFIG", "PlayerBar", 1, App.Path & "\config.ini"
    WriteINI "CONFIG", "PlayerName", 1, App.Path & "\config.ini"
    WriteINI "CONFIG", "PlayerDamage", 1, App.Path & "\config.ini"
    WriteINI "CONFIG", "MapGrid", 1, App.Path & "\config.ini"
    WriteINI "CONFIG", "Music", 1, App.Path & "\config.ini"
    WriteINI "CONFIG", "Sound", 1, App.Path & "\config.ini"
    WriteINI "CONFIG", "AutoScroll", 1, App.Path & "\config.ini"
    WriteINI "CONFIG", "Auto", 0, App.Path & "\config.ini"
End Sub

Private Sub FileCreateNewsINI()
    WriteINI "DATA", "News", vbNullString, App.Path & "\News.ini"
    WriteINI "DATA", "Desc", vbNullString, App.Path & "\News.ini"

    WriteINI "COLOR", "Red", 255, App.Path & "\News.ini"
    WriteINI "COLOR", "Green", 255, App.Path & "\News.ini"
    WriteINI "COLOR", "Blue", 255, App.Path & "\News.ini"

    WriteINI "FONT", "Font", "Arial", App.Path & "\News.ini"
    WriteINI "FONT", "Size", "14", App.Path & "\News.ini"
End Sub

Private Sub FileCreateFontINI()
    Call WriteINI("FONT", "Font", "fixedsys", App.Path & "\Font.ini")
    Call WriteINI("FONT", "Size", 18, App.Path & "\Font.ini")
End Sub

Private Sub LoadColors()
    Dim R1 As Long
    Dim G1 As Long
    Dim B1 As Long

    On Error GoTo ErrorHandle

    ' chat box color
    R1 = CInt(ReadINI("CHATBOX", "R", App.Path & "\GUI\Colors.txt"))
    G1 = CInt(ReadINI("CHATBOX", "G", App.Path & "\GUI\Colors.txt"))
    B1 = CInt(ReadINI("CHATBOX", "B", App.Path & "\GUI\Colors.txt"))
    frmMirage.txtChat.BackColor = RGB(R1, G1, B1)

    ' chat box text color
    R1 = CInt(ReadINI("CHATTEXTBOX", "R", App.Path & "\GUI\Colors.txt"))
    G1 = CInt(ReadINI("CHATTEXTBOX", "G", App.Path & "\GUI\Colors.txt"))
    B1 = CInt(ReadINI("CHATTEXTBOX", "B", App.Path & "\GUI\Colors.txt"))
    frmMirage.txtMyTextBox.BackColor = RGB(R1, G1, B1)

    R1 = CInt(ReadINI("SPELLLIST", "R", App.Path & "\GUI\Colors.txt"))
    G1 = CInt(ReadINI("SPELLLIST", "G", App.Path & "\GUI\Colors.txt"))
    B1 = CInt(ReadINI("SPELLLIST", "B", App.Path & "\GUI\Colors.txt"))
    frmMirage.lstSpells.BackColor = RGB(R1, G1, B1)

    R1 = CInt(ReadINI("WHOLIST", "R", App.Path & "\GUI\Colors.txt"))
    G1 = CInt(ReadINI("WHOLIST", "G", App.Path & "\GUI\Colors.txt"))
    B1 = CInt(ReadINI("WHOLIST", "B", App.Path & "\GUI\Colors.txt"))
    frmMirage.lstOnline.BackColor = RGB(R1, G1, B1)

    R1 = CInt(ReadINI("NEWCHAR", "R", App.Path & "\GUI\Colors.txt"))
    G1 = CInt(ReadINI("NEWCHAR", "G", App.Path & "\GUI\Colors.txt"))
    B1 = CInt(ReadINI("NEWCHAR", "B", App.Path & "\GUI\Colors.txt"))
    frmNewChar.optMale.BackColor = RGB(R1, G1, B1)
    frmNewChar.optFemale.BackColor = RGB(R1, G1, B1)

    R1 = CInt(ReadINI("BACKGROUND", "R", App.Path & "\GUI\Colors.txt"))
    G1 = CInt(ReadINI("BACKGROUND", "G", App.Path & "\GUI\Colors.txt"))
    B1 = CInt(ReadINI("BACKGROUND", "B", App.Path & "\GUI\Colors.txt"))

    frmMirage.picInventory3.BackColor = RGB(R1, G1, B1)
    frmMirage.picInventory.BackColor = RGB(R1, G1, B1)
    frmMirage.itmDesc.BackColor = RGB(R1, G1, B1)
    frmMirage.picWhosOnline.BackColor = RGB(R1, G1, B1)
    frmMirage.picGuildAdmin.BackColor = RGB(R1, G1, B1)
    frmMirage.picGuildMember.BackColor = RGB(R1, G1, B1)
    frmMirage.picEquipment.BackColor = RGB(R1, G1, B1)
    frmMirage.picPlayerSpells.BackColor = RGB(R1, G1, B1)
    frmMirage.picOptions.BackColor = RGB(R1, G1, B1)

    frmMirage.chkBubbleBar.BackColor = RGB(R1, G1, B1)
    frmMirage.chkNpcBar.BackColor = RGB(R1, G1, B1)
    frmMirage.chkNpcName.BackColor = RGB(R1, G1, B1)
    frmMirage.chkPlayerBar.BackColor = RGB(R1, G1, B1)
    frmMirage.chkPlayerName.BackColor = RGB(R1, G1, B1)
    frmMirage.chkPlayerDamage.BackColor = RGB(R1, G1, B1)
    frmMirage.chkNpcDamage.BackColor = RGB(R1, G1, B1)
    frmMirage.chkMusic.BackColor = RGB(R1, G1, B1)
    frmMirage.chkSound.BackColor = RGB(R1, G1, B1)
    frmMirage.chkAutoScroll.BackColor = RGB(R1, G1, B1)

    Exit Sub

ErrorHandle:
    Call MsgBox("Error loading colors.txt")

End Sub

Private Sub FileCreateColorsTXT()
    WriteINI "CHATBOX", "R", 152, App.Path & "\GUI\Colors.txt"
    WriteINI "CHATBOX", "G", 146, App.Path & "\GUI\Colors.txt"
    WriteINI "CHATBOX", "B", 120, App.Path & "\GUI\Colors.txt"

    WriteINI "CHATTEXTBOX", "R", 152, App.Path & "\GUI\Colors.txt"
    WriteINI "CHATTEXTBOX", "G", 146, App.Path & "\GUI\Colors.txt"
    WriteINI "CHATTEXTBOX", "B", 120, App.Path & "\GUI\Colors.txt"

    WriteINI "BACKGROUND", "R", 152, App.Path & "\GUI\Colors.txt"
    WriteINI "BACKGROUND", "G", 146, App.Path & "\GUI\Colors.txt"
    WriteINI "BACKGROUND", "B", 120, App.Path & "\GUI\Colors.txt"

    WriteINI "SPELLLIST", "R", 152, App.Path & "\GUI\Colors.txt"
    WriteINI "SPELLLIST", "G", 146, App.Path & "\GUI\Colors.txt"
    WriteINI "SPELLLIST", "B", 120, App.Path & "\GUI\Colors.txt"

    WriteINI "WHOLIST", "R", 152, App.Path & "\GUI\Colors.txt"
    WriteINI "WHOLIST", "G", 146, App.Path & "\GUI\Colors.txt"
    WriteINI "WHOLIST", "B", 120, App.Path & "\GUI\Colors.txt"

    WriteINI "NEWCHAR", "R", 152, App.Path & "\GUI\Colors.txt"
    WriteINI "NEWCHAR", "G", 146, App.Path & "\GUI\Colors.txt"
    WriteINI "NEWCHAR", "B", 120, App.Path & "\GUI\Colors.txt"
End Sub

Private Sub LoadFont()
    On Error GoTo ErrorHandle

    Font = ReadINI("FONT", "Font", App.Path & "\Font.ini")
    fontsize = CByte(ReadINI("FONT", "Size", App.Path & "\Font.ini"))

    If Font = vbNullString Then
        Font = "fixedsys"
    End If

    If fontsize <= 0 Or fontsize > 32 Then
        fontsize = 18
    End If

    Call SetFont(Font, fontsize)

    Exit Sub

ErrorHandle:
    Call WriteINI("FONT", "Font", "fixedsys", App.Path & "\Font.ini")
    Call WriteINI("FONT", "Size", 18, App.Path & "\Font.ini")

    Call SetFont("fixedsys", 18)

End Sub
