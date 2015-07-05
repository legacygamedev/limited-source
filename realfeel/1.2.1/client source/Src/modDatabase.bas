Attribute VB_Name = "modDatabase"
Option Explicit

Public Function FileExist(ByVal FileName As String, Optional RAW As Boolean = False) As Boolean
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 07/12/2005  Shannara   Optimized function.
'****************************************************************
    
    If RAW = False Then
        If Dir(App.Path & "\" & FileName) = "" Then
            FileExist = False
            Exit Function
        Else
            FileExist = True
            Exit Function
        End If
    Else
        If Dir(FileName) = "" Then
            FileExist = False
            Exit Function
        Else
            FileExist = True
        End If
    End If
End Function

Public Sub AddLog(ByVal Text As String)
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 07/12/2005  Shannara   Added log constants.
'****************************************************************

Dim FileName As String
Dim f As Long
Dim LOG_DEBUG As String
LOG_DEBUG = "debug.txt"

    If Trim$(Command) = "-debug" Then
        If frmDebug.Visible = False Then
            frmDebug.Visible = True
        End If
        
        FileName = App.Path & LOG_PATH & LOG_DEBUG
    
        If Not FileExist(LOG_DEBUG, True) Then
            f = FreeFile
            Open FileName For Output As #f
            Close #f
        End If
    
        f = FreeFile
        Open FileName For Append As #f
            Print #f, Time & ": " & Text
        Close #f
    End If
End Sub

Public Sub SaveLocalMap(ByVal MapNum As Long)
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 07/12/2005  Shannara   Added map constants.
'****************************************************************

Dim FileName As String
Dim f As Long

    FileName = App.Path & MAP_PATH & "map" & MapNum & MAP_EXT
            
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , SaveMap
    Close #f
End Sub

Public Sub LoadMap(ByVal MapNum As Long)
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 07/12/2005  Shannara   Added map constants.
'****************************************************************

Dim FileName As String
Dim f As Long

    FileName = App.Path & MAP_PATH & "map" & MapNum & MAP_EXT
        
    f = FreeFile
    Open FileName For Binary As #f
        Get #f, , SaveMap
    Close #f
End Sub

Public Function GetMapRevision(ByVal MapNum As Long) As Long
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 07/12/2005  Shannara   Added map constants.
'****************************************************************

Dim FileName As String
Dim f As Long
Dim TmpMap As MapRec

    FileName = App.Path & MAP_PATH & "map" & MapNum & MAP_EXT
        
    f = FreeFile
    Open FileName For Binary As #f
        Get #f, , TmpMap
    Close #f
        
    GetMapRevision = TmpMap.Revision
End Function

Public Function GetVar(File As String, Header As String, Var As String) As String
Dim sSpaces As String   ' Max string length
Dim szReturn As String  ' Return default value if not found
  
    szReturn = ""
  
    sSpaces = Space(5000)
  
    Call GetPrivateProfileString(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
  
    GetVar = RTrim$(sSpaces)
    GetVar = Left(GetVar, Len(GetVar) - 1)
End Function

Public Sub PutVar(File As String, Header As String, Var As String, Value As String)
    Call WritePrivateProfileString(Header, Var, Value, File)
End Sub

Public Function VarExists(File As String, Header As String, Var As String) As Boolean
On Error GoTo HandleError
Dim sSpaces As String   ' Max string length
Dim szReturn As String  ' Return default value if not found
  
    szReturn = ""
  
    sSpaces = Space(5000)
  
    Call GetPrivateProfileString(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
  
    VarExists = True
    Exit Function
    
HandleError:
    VarExists = False
    Exit Function
    
End Function

Public Sub LoadMenu()
Dim IP As String
Dim Port As Integer
    frmMainMenu.Visible = True
'Check data to prevent error
If GetVar(App.Path & "\data.dat", "Address", "IP") = "" Then
    Call PutVar(App.Path & "\data.dat", "Address", "IP", CStr(frmDualSolace.Socket.LocalIP))
    IP = CStr(frmDualSolace.Socket.LocalIP)
Else
    IP = GetVar(App.Path & "\data.dat", "Address", "IP")
End If
If GetVar(App.Path & "\data.dat", "Address", "Port") = "" Then
    Call PutVar(App.Path & "\data.dat", "Address", "Port", frmDualSolace.Socket.LocalPort)
    Port = frmDualSolace.Socket.LocalPort
Else
    Port = GetVar(App.Path & "\data.dat", "Address", "Port")
End If
End Sub

Public Sub SetName(ByVal Name As String)
    frmDualSolace.Caption = Name
End Sub

Sub MovePicture(PB As PictureBox, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        PB.Left = PB.Left + X - SOffsetX
        PB.top = PB.top + Y - SOffsetY
    End If
End Sub

Sub MoveForm(frm As Form, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        frm.Left = frm.Left + X - SOffsetX
        frm.top = frm.top + Y - SOffsetY
    End If
End Sub

Public Sub PicFrameCheck()
Dim n As Long, cn As Long

' Check all PicFrame" & num & " variables
cn = 0
If CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Characters", "PicFrameInit")) <> 0 Then cn = cn + CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Characters", "PicFrameInit"))
If CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Delete Account", "PicFrameInit")) <> 0 Then cn = cn + CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Delete Account", "PicFrameInit"))
If CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Fix Item", "PicFrameInit")) <> 0 Then cn = cn + CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Fix Item", "PicFrameInit"))
If CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Game Settings", "PicFrameInit")) <> 0 Then cn = cn + CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Game Settings", "PicFrameInit"))
If CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "PicFrameInit")) <> 0 Then cn = cn + CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "PicFrameInit"))
If CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Menu Settings", "PicFrameInit")) <> 0 Then cn = cn + CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Menu Settings", "PicFrameInit"))
If CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Account", "PicFrameInit")) <> 0 Then cn = cn + CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Account", "PicFrameInit"))
If CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "PicFrameInit")) <> 0 Then cn = cn + CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "PicFrameInit"))
If CLng(GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "PicFrameInit")) <> 0 Then cn = cn + CLng(GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "PicFrameInit"))
If CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Information Window", "PicFrameInit")) <> 0 Then cn = cn + CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Information Window", "PicFrameInit"))
If CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "PicFrameInit")) <> 0 Then cn = cn + CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "PicFrameInit"))

If cn <> 0 Then
ReDim PicArray(1 To cn) As VB.PictureBox
If CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Characters", "PicFrameInit")) <> 0 Then
    For n = 1 To CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Characters", "PicFrameInit"))
        Call frmChars.SetArray(CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Characters", "PicFrameInit")))
        Call frmChars.MakePic(n)
    Next n
End If
If CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Delete Account", "PicFrameInit")) <> 0 Then
    For n = 1 To CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Delete Account", "PicFrameInit"))
        Call frmDeleteAccount.SetArray(CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Delete Account", "PicFrameInit")))
        Call frmDeleteAccount.MakePic(n)
    Next n
End If
If CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Fix Item", "PicFrameInit")) <> 0 Then
    For n = 1 To CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Fix Item", "PicFrameInit"))
        Call frmFixItem.SetArray(CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Fix Item", "PicFrameInit")))
        Call frmFixItem.MakePic(n)
    Next n
End If
If CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Game Settings", "PicFrameInit")) <> 0 Then
    For n = 1 To CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Game Settings", "PicFrameInit"))
        Call frmGameSettings.SetArray(CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Game Settings", "PicFrameInit")))
        Call frmGameSettings.MakePic(n)
    Next n
End If
If CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "PicFrameInit")) <> 0 Then
    For n = 1 To CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "PicFrameInit"))
        Call frmMainMenu.SetArray(CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "PicFrameInit")))
        Call frmMainMenu.MakePic(n)
    Next n
End If
If CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Menu Settings", "PicFrameInit")) <> 0 Then
    For n = 1 To CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Menu Settings", "PicFrameInit"))
        Call frmSettings.SetArray(CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Menu Settings", "PicFrameInit")))
        Call frmSettings.MakePic(n)
    Next n
End If
If CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Account", "PicFrameInit")) <> 0 Then
    For n = 1 To CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Account", "PicFrameInit"))
        Call frmNewAccount.SetArray(CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Account", "PicFrameInit")))
        Call frmNewAccount.MakePic(n)
    Next n
End If
If CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "PicFrameInit")) <> 0 Then
    For n = 1 To CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "PicFrameInit"))
        Call frmNewChar.SetArray(CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "PicFrameInit")))
        Call frmNewChar.MakePic(n)
    Next n
End If
If CLng(GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "PicFrameInit")) <> 0 Then
    For n = 1 To CLng(GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "PicFrameInit"))
        Call frmDualSolace.SetArray(CLng(GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "PicFrameInit")))
        Call frmDualSolace.MakePic(n)
    Next n
End If
If CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Information Window", "PicFrameInit")) <> 0 Then
    For n = 1 To CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Information Window", "PicFrameInit"))
        Call frmSendGetData.SetArray(CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Information Window", "PicFrameInit")))
        Call frmSendGetData.MakePic(n)
    Next n
End If
If CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "PicFrameInit")) <> 0 Then
    For n = 1 To CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "PicFrameInit"))
        Call frmTrade.SetArray(CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "PicFrameInit")))
        Call frmTrade.MakePic(n)
    Next n
End If
End If
End Sub

Public Sub LoadGUI()
' Load Character GUI
Call SetFormSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Characters", "Background"), frmChars)
With frmChars
.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Characters", "Background"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Characters", "Heading"), .picHeading)
.picHeading.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Characters", "Heading"))
.picHeading.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Characters", "HeadingX"))
.picHeading.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Characters", "HeadingY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Characters", "UseButton"), .picUseChar)
.picUseChar.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Characters", "UseButton"))
.picUseChar.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Characters", "UseButtonX"))
.picUseChar.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Characters", "UseButtonY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Characters", "NewButton"), .picNewChar)
.picNewChar.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Characters", "NewButton"))
.picNewChar.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Characters", "NewButtonX"))
.picNewChar.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Characters", "NewButtonY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Characters", "DeleteButton"), .picDeleteChar)
.picDeleteChar.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Characters", "DeleteButton"))
.picDeleteChar.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Characters", "DeleteButtonX"))
.picDeleteChar.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Characters", "DeleteButtonY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Characters", "CancelButton"), .picCancel)
.picCancel.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Characters", "CancelButton"))
.picCancel.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Characters", "CancelButtonX"))
.picCancel.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Characters", "CancelButtonY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Characters", "CancelButton"), .picCancel)
.picCancel.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Characters", "CancelButton"))
.picCancel.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Characters", "CancelButtonX"))
.picCancel.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Characters", "CancelButtonY"))

' Set Char List
.lstChars.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Characters", "ListCharX"))
.lstChars.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Characters", "ListCharY"))
.lstChars.ForeColor = GetVar(App.Path & GUI_PATH & "config.ini", "Characters", "ListCharForecolor")
.lstChars.BackColor = GetVar(App.Path & GUI_PATH & "config.ini", "Characters", "ListCharBackcolor")
End With
' End Load Character GUI

' Load Delete Account GUI
Call SetFormSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Delete Account", "Background"), frmDeleteAccount)
With frmDeleteAccount
.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Delete Account", "Background"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Delete Account", "Heading"), .picHeading)
.picHeading.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Delete Account", "Heading"))
.picHeading.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Delete Account", "HeadingX"))
.picHeading.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Delete Account", "HeadingY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Delete Account", "ConnectButton"), .picConnect)
.picConnect.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Delete Account", "ConnectButton"))
.picConnect.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Delete Account", "ConnectButtonX"))
.picConnect.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Delete Account", "ConnectButtonY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Delete Account", "CancelButton"), .picCancel)
.picCancel.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Delete Account", "CancelButton"))
.picCancel.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Delete Account", "CancelButtonX"))
.picCancel.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Delete Account", "CancelButtonY"))

' Set Name TextBox
.txtName.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Delete Account", "TextNameX"))
.txtName.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Delete Account", "TextNameY"))
.txtName.ForeColor = GetVar(App.Path & GUI_PATH & "config.ini", "Delete Account", "TextNameForecolor")
.txtName.BackColor = GetVar(App.Path & GUI_PATH & "config.ini", "Delete Account", "TextNameBackcolor")

' Set Password TextBox
.txtPassword.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Delete Account", "TextPasswordX"))
.txtPassword.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Delete Account", "TextPasswordY"))
.txtPassword.ForeColor = GetVar(App.Path & GUI_PATH & "config.ini", "Delete Account", "TextPasswordForecolor")
.txtPassword.BackColor = GetVar(App.Path & GUI_PATH & "config.ini", "Delete Account", "TextPasswordBackcolor")

End With
' End Load Delete Account GUI

' Load Fix Item GUI
Call SetFormSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Fix Item", "Background"), frmFixItem)
With frmFixItem
.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Fix Item", "Background"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Fix Item", "Heading"), .picHeading)
.picHeading.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Fix Item", "Heading"))
.picHeading.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Fix Item", "HeadingX"))
.picHeading.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Fix Item", "HeadingY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Fix Item", "RepairButton"), .picFix)
.picFix.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Fix Item", "RepairButton"))
.picFix.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Fix Item", "RepairButtonX"))
.picFix.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Fix Item", "RepairButtonY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Fix Item", "CancelButton"), .picCancel)
.picCancel.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Fix Item", "CancelButton"))
.picCancel.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Fix Item", "CancelButtonX"))
.picCancel.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Fix Item", "CancelButtonY"))

' Set Item ComboBox
.cmbItem.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Fix Item", "ComboBoxX"))
.cmbItem.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Fix Item", "ComboBoxY"))
.cmbItem.ForeColor = GetVar(App.Path & GUI_PATH & "config.ini", "Fix Item", "ComboBoxForecolor")
.cmbItem.BackColor = GetVar(App.Path & GUI_PATH & "config.ini", "Fix Item", "ComboBoxBackcolor")

End With
' End Load Fix Item GUI

' Load Game Settings GUI
Call SetFormSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Game Settings", "Background"), frmGameSettings)
With frmGameSettings
.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Game Settings", "Background"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Game Settings", "SubmitButton"), .picSubmit)
.picSubmit.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Game Settings", "SubmitButton"))
.picSubmit.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Game Settings", "SubmitButtonX"))
.picSubmit.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Game Settings", "SubmitButtonY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Game Settings", "CancelButton"), .picCancel)
.picCancel.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Game Settings", "CancelButton"))
.picCancel.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Game Settings", "CancelButtonX"))
.picCancel.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Game Settings", "CancelButtonY"))
End With
' End Load Fix Item GUI

' Load Main Menu GUI
Call SetFormSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "Background"), frmMainMenu)
With frmMainMenu
.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "Background"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "Heading"), .picHeading)
.picHeading.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "Heading"))
.picHeading.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "HeadingX"))
.picHeading.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "HeadingY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "NewButton"), .picNewAccount)
.picNewAccount.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "NewButton"))
.picNewAccount.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "NewButtonX"))
.picNewAccount.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "NewButtonY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "DeleteButton"), .picDelAccount)
.picDelAccount.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "DeleteButton"))
.picDelAccount.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "DeleteButtonX"))
.picDelAccount.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "DeleteButtonY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "CreditButton"), .picCredits)
.picCredits.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "CreditButton"))
.picCredits.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "CreditButtonX"))
.picCredits.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "CreditButtonY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "ConnectButton"), .picConnect)
.picConnect.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "ConnectButton"))
.picConnect.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "ConnectButtonX"))
.picConnect.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "ConnectButtonY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "RefreshButton"), .picRefresh)
.picRefresh.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "RefreshButton"))
.picRefresh.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "RefreshButtonX"))
.picRefresh.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "RefreshButtonY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "SettingsButton"), .picSettings)
.picSettings.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "SettingsButton"))
.picSettings.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "SettingsButtonX"))
.picSettings.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "SettingsButtonY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "CancelButton"), .picQuit)
.picQuit.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "CancelButton"))
.picQuit.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "CancelButtonX"))
.picQuit.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "CancelButtonY"))

' Set Name TextBox
.txtName.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "TextNameX"))
.txtName.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "TextNameY"))
.txtName.ForeColor = GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "TextNameForecolor")
.txtName.BackColor = GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "TextNameBackcolor")

' Set Password TextBox
.txtPassword.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "TextPasswordX"))
.txtPassword.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "TextPasswordY"))
.txtPassword.ForeColor = GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "TextPasswordForecolor")
.txtPassword.BackColor = GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "TextPasswordBackcolor")

' Set Info TextBox
.txtInfo.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "TextInfoX"))
.txtInfo.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "TextInfoY"))
.txtInfo.ForeColor = GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "TextInfoForecolor")
.txtInfo.BackColor = GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "TextInfoBackcolor")

' Set Status Label
.lblStatus.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "LabelStatusX"))
.lblStatus.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "LabelStatusY"))
.lblStatus.ForeColor = GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "LabelStatusForecolor")
.lblStatus.BackColor = GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "LabelStatusBackcolor")
If LCase$(GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "LabelStatusTransparent")) = "true" Then .lblStatus.BackStyle = 0
If LCase$(GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "LabelStatusTransparent")) <> "true" Then .lblStatus.BackStyle = 1
End With
' End Load Main Menu GUI

' Load Menu Settings GUI
Call SetFormSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Menu Settings", "Background"), frmSettings)
With frmSettings
.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Menu Settings", "Background"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Menu Settings", "Heading"), .picHeading)
.picHeading.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Menu Settings", "Heading"))
.picHeading.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Menu Settings", "HeadingX"))
.picHeading.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Menu Settings", "HeadingY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Menu Settings", "SaveButton"), .picSubmit)
.picSubmit.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Menu Settings", "SaveButton"))
.picSubmit.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Menu Settings", "SaveButtonX"))
.picSubmit.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Menu Settings", "SaveButtonY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Menu Settings", "CancelButton"), .picCancel)
.picCancel.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Menu Settings", "CancelButton"))
.picCancel.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Menu Settings", "CancelButtonX"))
.picCancel.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Menu Settings", "CancelButtonY"))

' Set IP TextBox
.txtIP.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Menu Settings", "TextIPX"))
.txtIP.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Menu Settings", "TextIPY"))
.txtIP.ForeColor = GetVar(App.Path & GUI_PATH & "config.ini", "Menu Settings", "TextIPForecolor")
.txtIP.BackColor = GetVar(App.Path & GUI_PATH & "config.ini", "Menu Settings", "TextIPBackcolor")

' Set Port TextBox
.txtPort.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Menu Settings", "TextPortX"))
.txtPort.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Menu Settings", "TextPortY"))
.txtPort.ForeColor = GetVar(App.Path & GUI_PATH & "config.ini", "Menu Settings", "TextPortForecolor")
.txtPort.BackColor = GetVar(App.Path & GUI_PATH & "config.ini", "Menu Settings", "TextPortBackcolor")

' Set Account Frame
.fraAccount.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Menu Settings", "FrameAccountX"))
.fraAccount.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Menu Settings", "FrameAccountY"))
.fraAccount.ForeColor = GetVar(App.Path & GUI_PATH & "config.ini", "Menu Settings", "FrameAccountForecolor")
.fraAccount.BackColor = GetVar(App.Path & GUI_PATH & "config.ini", "Menu Settings", "FrameAccountBackcolor")

' Set On Option
.optOn.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Menu Settings", "OptionOnX"))
.optOn.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Menu Settings", "OptionOnY"))
.optOn.ForeColor = GetVar(App.Path & GUI_PATH & "config.ini", "Menu Settings", "OptionOnForecolor")
.optOn.BackColor = GetVar(App.Path & GUI_PATH & "config.ini", "Menu Settings", "OptionOnBackcolor")

' Set Off Option
.optOff.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Menu Settings", "OptionOffX"))
.optOff.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Menu Settings", "OptionOffY"))
.optOff.ForeColor = GetVar(App.Path & GUI_PATH & "config.ini", "Menu Settings", "OptionOffForecolor")
.optOff.BackColor = GetVar(App.Path & GUI_PATH & "config.ini", "Menu Settings", "OptionOffBackcolor")
End With
' End Load Menu Settings GUI

' Load New Account GUI
Call SetFormSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "New Account", "Background"), frmNewAccount)
With frmNewAccount
.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "New Account", "Background"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "New Account", "Heading"), .picHeading)
.picHeading.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "New Account", "Heading"))
.picHeading.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Account", "HeadingX"))
.picHeading.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Account", "HeadingY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "New Account", "ConnectButton"), .picConnect)
.picConnect.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "New Account", "ConnectButton"))
.picConnect.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Account", "ConnectButtonX"))
.picConnect.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Account", "ConnectButtonY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "New Account", "CancelButton"), .picCancel)
.picCancel.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "New Account", "CancelButton"))
.picCancel.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Account", "CancelButtonX"))
.picCancel.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Account", "CancelButtonY"))

' Set Name TextBox
.txtName.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Account", "TextNameX"))
.txtName.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Account", "TextNameY"))
.txtName.ForeColor = GetVar(App.Path & GUI_PATH & "config.ini", "New Account", "TextNameForecolor")
.txtName.BackColor = GetVar(App.Path & GUI_PATH & "config.ini", "New Account", "TextNameBackcolor")

' Set Password TextBox
.txtPassword.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Account", "TextPasswordX"))
.txtPassword.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Account", "TextPasswordY"))
.txtPassword.ForeColor = GetVar(App.Path & GUI_PATH & "config.ini", "New Account", "TextPasswordForecolor")
.txtPassword.BackColor = GetVar(App.Path & GUI_PATH & "config.ini", "New Account", "TextPasswordBackcolor")
End With
' End Load New Account GUI

' Load New Character GUI
Call SetFormSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "Background"), frmNewChar)
With frmNewChar
.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "Background"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "Heading"), .picHeading)
.picHeading.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "Heading"))
.picHeading.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "HeadingX"))
.picHeading.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "HeadingY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "SaveButton"), .picAddChar)
.picAddChar.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "SaveButton"))
.picAddChar.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "SaveButtonX"))
.picAddChar.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "SaveButtonY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "CancelButton"), .picCancel)
.picCancel.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "CancelButton"))
.picCancel.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "CancelButtonX"))
.picCancel.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "CancelButtonY"))

' Set Name TextBox
.txtName.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "TextNameX"))
.txtName.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "TextNameY"))
.txtName.ForeColor = GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "TextNameForecolor")
.txtName.BackColor = GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "TextNameBackcolor")

' Set Male Option
.optMale.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "OptionMaleX"))
.optMale.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "OptionMaleY"))
.optMale.ForeColor = GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "OptionMaleForecolor")
.optMale.BackColor = GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "OptionMaleBackcolor")

' Set Female Option
.optFemale.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "OptionFemaleX"))
.optFemale.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "OptionFemaleY"))
.optFemale.ForeColor = GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "OptionFemaleForecolor")
.optFemale.BackColor = GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "OptionFemaleBackcolor")

' Set Sprite Picturebox
.picSprite.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "PictureSpriteX"))
.picSprite.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "PictureSpriteY"))

' Set STR Label
.lblStr.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "LabelSTRX"))
.lblStr.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "LabelSTRY"))
.lblStr.ForeColor = GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "LabelSTRForecolor")
.lblStr.BackColor = GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "LabelSTRBackcolor")
If LCase$(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "LabelSTRTransparent")) = "true" Then .lblStr.BackStyle = 0
If LCase$(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "LabelSTRTransparent")) <> "true" Then .lblStr.BackStyle = 1

' Set DEF Label
.lblDef.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "LabelDEFX"))
.lblDef.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "LabelDEFY"))
.lblDef.ForeColor = GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "LabelDEFForecolor")
.lblDef.BackColor = GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "LabelDEFBackcolor")
If LCase$(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "LabelDEFTransparent")) = "true" Then .lblDef.BackStyle = 0
If LCase$(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "LabelDEFTransparent")) <> "true" Then .lblDef.BackStyle = 1

' Set MAGI Label
.lblMagi.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "LabelMAGIX"))
.lblMagi.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "LabelMAGIY"))
.lblMagi.ForeColor = GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "LabelMAGIForecolor")
.lblMagi.BackColor = GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "LabelMAGIBackcolor")
If LCase$(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "LabelMAGITransparent")) = "true" Then .lblMagi.BackStyle = 0
If LCase$(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "LabelMAGITransparent")) <> "true" Then .lblMagi.BackStyle = 1

' Set SPEED Label
.lblSPEED.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "LabelSPEEDX"))
.lblSPEED.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "LabelSPEEDY"))
.lblSPEED.ForeColor = GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "LabelSPEEDForecolor")
.lblSPEED.BackColor = GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "LabelSPEEDBackcolor")
If LCase$(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "LabelSPEEDTransparent")) = "true" Then .lblSPEED.BackStyle = 0
If LCase$(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "LabelSPEEDTransparent")) <> "true" Then .lblSPEED.BackStyle = 1

' Set HP Label
.lblHP.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "LabelHPX"))
.lblHP.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "LabelHPY"))
.lblHP.ForeColor = GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "LabelHPForecolor")
.lblHP.BackColor = GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "LabelHPBackcolor")
If LCase$(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "LabelHPTransparent")) = "true" Then .lblHP.BackStyle = 0
If LCase$(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "LabelHPTransparent")) <> "true" Then .lblHP.BackStyle = 1

' Set MP Label
.lblMP.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "LabelMPX"))
.lblMP.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "LabelMPY"))
.lblMP.ForeColor = GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "LabelMPForecolor")
.lblMP.BackColor = GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "LabelMPBackcolor")
If LCase$(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "LabelMPTransparent")) = "true" Then .lblMP.BackStyle = 0
If LCase$(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "LabelMPTransparent")) <> "true" Then .lblMP.BackStyle = 1

' Set SP Label
.lblSP.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "LabelSPX"))
.lblSP.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "LabelSPY"))
.lblSP.ForeColor = GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "LabelSPForecolor")
.lblSP.BackColor = GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "LabelSPBackcolor")
If LCase$(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "LabelSPTransparent")) = "true" Then .lblSP.BackStyle = 0
If LCase$(GetVar(App.Path & GUI_PATH & "config.ini", "New Character", "LabelSPTransparent")) <> "true" Then .lblSP.BackStyle = 1
End With
' End Load New Character GUI

' Load RealFeel Game Screen GUI
Call SetFormSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "Background"), frmDualSolace)
With frmDualSolace
.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "Background"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "StatsFrame"), .picPlayer)
.picPlayer.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "StatsFrame"))
.picPlayer.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "StatsFrameX"))
.picPlayer.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "StatsFrameY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "EquipFrame"), .picEquipment)
.picEquipment.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "EquipFrame"))
.picEquipment.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "EquipFrameX"))
.picEquipment.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "EquipFrameY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "PanelFrame"), .picGUI)
.picGUI.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "PanelFrame"))
.picGUI.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "PanelFrameX"))
.picGUI.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "PanelFrameY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "UpgradeButton"), .picTrain)
.picTrain.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "UpgradeButton"))
.picTrain.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "UpgradeButtonX"))
.picTrain.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "UpgradeButtonY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "InventoryButton"), .picInventory)
.picInventory.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "InventoryButton"))
.picInventory.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "InventoryButtonX"))
.picInventory.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "InventoryButtonY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "SpellButton"), .picSpells)
.picSpells.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "SpellButton"))
.picSpells.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "SpellButtonX"))
.picSpells.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "SpellButtonY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "SideFrame"), .picSidebar)
.picSidebar.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "SideFrame"))
.picSidebar.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "SideFrameX"))
.picSidebar.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "SideFrameY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "AddFriendButton"), .picAdd)
.picAdd.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "AddFriendButton"))
.picAdd.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "AddFriendButtonX"))
.picAdd.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "AddFriendButtonY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "RemoveFriendButton"), .picRemove)
.picRemove.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "RemoveFriendButton"))
.picRemove.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "RemoveFriendButtonX"))
.picRemove.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "RemoveFriendButtonY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "CancelButton"), .picCancel)
.picCancel.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "CancelButton"))
.picCancel.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "CancelButtonX"))
.picCancel.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "CancelButtonY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Admin Panel", "Background"), .picAdminPanel)
.picAdminPanel.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Admin Panel", "Background"))
.picAdminPanel.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Admin Panel", "PlacementX"))
.picAdminPanel.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Admin Panel", "PlacementY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Bank", "Background"), .picBank)
.picBank.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Bank", "Background"))
.picBank.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Bank", "PlacementX"))
.picBank.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Bank", "PlacementY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Bank", "DepositButton"), .picDeposit)
.picDeposit.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Bank", "DepositButton"))
.picDeposit.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Bank", "DepositButtonX"))
.picDeposit.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Bank", "DepositButtonY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Bank", "WithdrawButton"), .picWithdraw)
.picWithdraw.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Bank", "WithdrawButton"))
.picWithdraw.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Bank", "WithdrawButtonX"))
.picWithdraw.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Bank", "WithdrawButtonY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Bank", "CancelButton"), .picBankExit)
.picBankExit.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Bank", "CancelButton"))
.picBankExit.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Bank", "CancelButtonX"))
.picBankExit.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Bank", "CancelButtonY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Inventory", "Background"), .picInv)
.picInv.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Inventory", "Background"))
.picInv.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Inventory", "PlacementX"))
.picInv.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Inventory", "PlacementY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Inventory", "UseButton"), .picUseItem)
.picUseItem.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Inventory", "UseButton"))
.picUseItem.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Inventory", "UseButtonX"))
.picUseItem.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Inventory", "UseButtonY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Inventory", "DropButton"), .picDropItem)
.picDropItem.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Inventory", "DropButton"))
.picDropItem.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Inventory", "DropButtonX"))
.picDropItem.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Inventory", "DropButtonY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Inventory", "CancelButton"), .picCancel)
.picCancel.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Inventory", "CancelButton"))
.picCancel.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Inventory", "CancelButtonX"))
.picCancel.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Inventory", "CancelButtonY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Train", "Background"), .picTrainMenu)
.picTrainMenu.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Train", "Background"))
.picTrainMenu.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Train", "PlacementX"))
.picTrainMenu.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Train", "PlacementY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Train", "Heading"), .picHeaderTrain)
.picHeaderTrain.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Train", "Heading"))
.picHeaderTrain.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Train", "HeadingX"))
.picHeaderTrain.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Train", "HeadingY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Train", "TrainButton"), .picTrainButton)
.picTrainButton.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Train", "TrainButton"))
.picTrainButton.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Train", "TrainButtonX"))
.picTrainButton.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Train", "TrainButtonY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Train", "CancelButton"), .picTrainCancel)
.picTrainCancel.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Train", "CancelButton"))
.picTrainCancel.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Train", "CancelButtonX"))
.picTrainCancel.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Train", "CancelButtonY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Spell", "Background"), .picPlayerSpells)
.picPlayerSpells.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Spell", "Background"))
.picPlayerSpells.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Spell", "PlacementX"))
.picPlayerSpells.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Spell", "PlacementY"))

' Set ChatEnter TextBox
.txtChatEnter.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "TextChatEnterX"))
.txtChatEnter.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "TextChatEnterY"))
.txtChatEnter.ForeColor = GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "TextChatEnterForecolor")
.txtChatEnter.BackColor = GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "TextChatEnterBackcolor")

' Set Chat TextBox
.txtChat.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "TextChatX"))
.txtChat.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "TextChatY"))
.txtChat.BackColor = GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "TextChatBackcolor")

' Set Players ListBox
.lstPlayers.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "ListPlayersX"))
.lstPlayers.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "ListPlayersY"))
.lstPlayers.ForeColor = GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "ListPlayersForecolor")
.lstPlayers.BackColor = GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "ListPlayersBackcolor")

' Set Friends ListBox
.lstFriends.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "ListFriendsX"))
.lstFriends.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "ListFriendsY"))
.lstFriends.ForeColor = GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "ListFriendsForecolor")
.lstFriends.BackColor = GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "ListFriendsBackcolor")

' Set Friends ListBox
.fraOptions.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "FrameOptionsX"))
.fraOptions.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "FrameOptionsY"))
.fraOptions.ForeColor = GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "FrameOptionsForecolor")
.fraOptions.BackColor = GetVar(App.Path & GUI_PATH & "config.ini", "RealFeel Game Screen", "FrameOptionsBackcolor")
End With
' End Load RealFeel Game Screen GUI

' Load Information Window GUI
Call SetFormSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Information Window", "Background"), frmSendGetData)
With frmSendGetData
.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Information Window", "Background"))
' Set Status Label
.lblStatus.ForeColor = GetVar(App.Path & GUI_PATH & "config.ini", "Information Window", "LabelStatusForecolor")
.lblStatus.BackColor = GetVar(App.Path & GUI_PATH & "config.ini", "Information Window", "LabelStatusBackcolor")
If LCase$(GetVar(App.Path & GUI_PATH & "config.ini", "Information Window", "LabelStatusTransparent")) = "true" Then .lblStatus.BackStyle = 0
If LCase$(GetVar(App.Path & GUI_PATH & "config.ini", "Information Window", "LabelStatusTransparent")) <> "true" Then .lblStatus.BackStyle = 1
End With
' End Load Information Window GUI

' Load Shop GUI
Call SetFormSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "Background"), frmTrade)
With frmTrade
.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "Background"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "TradeButton"), .picDeal)
.picDeal.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "TradeButton"))
.picDeal.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "TradeButtonX"))
.picDeal.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "TradeButtonY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "RepairButton"), .picFixItems)
.picFixItems.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "RepairButton"))
.picFixItems.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "RepairButtonX"))
.picFixItems.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "RepairButtonY"))

Call SetPicSize(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "CancelButton"), .picCancel)
.picCancel.Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "CancelButton"))
.picCancel.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "CancelButtonX"))
.picCancel.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "CancelButtonY"))

' Set Name Label
.lblName.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelNameX"))
.lblName.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelNameY"))
.lblName.ForeColor = GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelNameForecolor")
.lblName.BackColor = GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelNameBackcolor")
If LCase$(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelNameTransparent")) = "true" Then .lblName.BackStyle = 0
If LCase$(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelNameTransparent")) <> "true" Then .lblName.BackStyle = 1

' Set Description Label
.lblDescription.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelDescriptionX"))
.lblDescription.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelDescriptionY"))
.lblDescription.ForeColor = GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelDescriptionForecolor")
.lblDescription.BackColor = GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelDescriptionBackcolor")
If LCase$(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelDescriptionTransparent")) = "true" Then .lblDescription.BackStyle = 0
If LCase$(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelDescriptionTransparent")) <> "true" Then .lblDescription.BackStyle = 1

' Set Cost Label
.lblCost.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelCostX"))
.lblCost.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelCostY"))
.lblCost.ForeColor = GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelCostForecolor")
.lblCost.BackColor = GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelCostBackcolor")
If LCase$(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelCostTransparent")) = "true" Then .lblCost.BackStyle = 0
If LCase$(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelCostTransparent")) <> "true" Then .lblCost.BackStyle = 1

' Set Restock Label
.lblRestock.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelRestockX"))
.lblRestock.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelRestockY"))
.lblRestock.ForeColor = GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelRestockForecolor")
.lblRestock.BackColor = GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelRestockBackcolor")
If LCase$(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelRestockTransparent")) = "true" Then .lblRestock.BackStyle = 0
If LCase$(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelRestockTransparent")) <> "true" Then .lblRestock.BackStyle = 1

' Set HSTR Label
.lblHSTR.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelHSTRX"))
.lblHSTR.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelHSTRY"))
.lblHSTR.ForeColor = GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelHSTRForecolor")
.lblHSTR.BackColor = GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelHSTRBackcolor")
If LCase$(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelHSTRTransparent")) = "true" Then .lblHSTR.BackStyle = 0
If LCase$(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelHSTRTransparent")) <> "true" Then .lblHSTR.BackStyle = 1

' Set HDEF Label
.lblHDEF.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelHDEFX"))
.lblHDEF.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelHDEFY"))
.lblHDEF.ForeColor = GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelHDEFForecolor")
.lblHDEF.BackColor = GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelHDEFBackcolor")
If LCase$(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelHDEFTransparent")) = "true" Then .lblHDEF.BackStyle = 0
If LCase$(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelHDEFTransparent")) <> "true" Then .lblHDEF.BackStyle = 1

' Set HMAG Label
.lblHMAG.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelHMAGX"))
.lblHMAG.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelHMAGY"))
.lblHMAG.ForeColor = GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelHMAGForecolor")
.lblHMAG.BackColor = GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelHMAGBackcolor")
If LCase$(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelHMAGTransparent")) = "true" Then .lblHMAG.BackStyle = 0
If LCase$(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelHMAGTransparent")) <> "true" Then .lblHMAG.BackStyle = 1

' Set HSPD Label
.lblHSPD.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelHSPDX"))
.lblHSPD.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelHSPDY"))
.lblHSPD.ForeColor = GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelHSPDForecolor")
.lblHSPD.BackColor = GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelHSPDBackcolor")
If LCase$(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelHSPDTransparent")) = "true" Then .lblHSPD.BackStyle = 0
If LCase$(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "LabelHSPDTransparent")) <> "true" Then .lblHSPD.BackStyle = 1

' Set Trade ListBox
.lstTrade.Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "ListTradeX"))
.lstTrade.top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "ListTradeY"))
.lstTrade.ForeColor = GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "ListTradeForecolor")
.lstTrade.BackColor = GetVar(App.Path & GUI_PATH & "config.ini", "Shop", "ListTradeBackcolor")
End With
' End Load Shop GUI
End Sub

Public Sub DataCheck()
Dim FilePath As String, f As Integer
Dim n As Long
f = FreeFile

' Set global paths
DLL_PATH = GetVar(App.Path & "\data.dat", "FilePaths", "DLLs")
GFX_PATH = GetVar(App.Path & "\data.dat", "FilePaths", "GFX")
LOG_PATH = GetVar(App.Path & "\data.dat", "FilePaths", "LOGS")
MAP_PATH = GetVar(App.Path & "\data.dat", "FilePaths", "MAPS")
MUSIC_PATH = GetVar(App.Path & "\data.dat", "FilePaths", "MUSIC")
SOUND_PATH = GetVar(App.Path & "\data.dat", "FilePaths", "SOUND")
GUI_PATH = GetVar(App.Path & "\data.dat", "FilePaths", "GUI")

FilePath = App.Path & GUI_PATH & "config.ini"
If Not FileExist(GUI_PATH & "config.ini") Then
    Open FilePath For Output As #f
        Print #f, ";This is the configuration for the client!"
        Print #f, ";Note: The Background for each window determines the size of the window."
        Print #f, ";Note2: All picture sizes are adjusted to pictures loaded."
        Print #f, ";Note3: PicFrameInit sets the number of PicFrames on the window."
        Print #f, ";Note4: For each PicFrame instance, a PicFrameTarget, PicFrameX, and PicFrameY must be added."
        Print #f, ";       PicFrame1Target, PicFrame1X, PicFrame1Y, PicFrame2Target..."
        Print #f, ";Note5: Keep in mind that any new PicFrames are placed on top of all other images."
        Print #f, ";Note6: Forecolors and Backcolors are determined from HTML color codes. The structure follows this guideline:"
        Print #f, ";       &H00_YOUR_COLOR_CODE_"
        
        Print #f, "" 'add spacing
        Print #f, "" 'add spacing
        
        Print #f, "[Characters]"
        Print #f, "Background="
        Print #f, "Heading="
        Print #f, "HeadingX="
        Print #f, "HeadingY="
        Print #f, "UseButton="
        Print #f, "UseButtonX="
        Print #f, "UseButtonY="
        Print #f, "NewButton="
        Print #f, "NewButtonX="
        Print #f, "NewButtonY="
        Print #f, "DeleteButton="
        Print #f, "DeleteButtonX="
        Print #f, "DeleteButtonY="
        Print #f, "CancelButton="
        Print #f, "CancelButtonX="
        Print #f, "CancelButtonY="
        Print #f, "ListCharForecolor="
        Print #f, "ListCharBackcolor="
        Print #f, "ListCharX="
        Print #f, "ListCharY="
        Print #f, "PicFrameInit=0"
        
        Print #f, "" 'add spacing
        Print #f, "" 'add spacing
        
        Print #f, "[Delete Account]"
        Print #f, "Background="
        Print #f, "Heading="
        Print #f, "HeadingX="
        Print #f, "HeadingY="
        Print #f, "ConnectButton="
        Print #f, "ConnectButtonX="
        Print #f, "ConnectButtonY="
        Print #f, "CancelButton="
        Print #f, "CancelButtonX="
        Print #f, "CancelButtonY="
        Print #f, "TextNameForecolor="
        Print #f, "TextNameBackcolor="
        Print #f, "TextNameX="
        Print #f, "TextNameY="
        Print #f, "TextPasswordForecolor="
        Print #f, "TextPasswordBackcolor="
        Print #f, "TextPasswordX="
        Print #f, "TextPasswordY="
        Print #f, "PicFrameInit=0"
        
        Print #f, "" 'add spacing
        Print #f, "" 'add spacing
        
        Print #f, "[Fix Item]"
        Print #f, "Background="
        Print #f, "Heading="
        Print #f, "HeadingX="
        Print #f, "HeadingY="
        Print #f, "RepairButton="
        Print #f, "RepairButtonX="
        Print #f, "RepairButtonY="
        Print #f, "CancelButton="
        Print #f, "CancelButtonX="
        Print #f, "CancelButtonY="
        Print #f, "ComboBoxForecolor="
        Print #f, "ComboBoxBackcolor="
        Print #f, "ComboBoxX="
        Print #f, "ComboBoxY="
        Print #f, "PicFrameInit=0"
        
        Print #f, "" 'add spacing
        Print #f, "" 'add spacing
        
        Print #f, "[Game Settings]"
        Print #f, "Background="
        Print #f, "SubmitButton="
        Print #f, "SubmitButtonX="
        Print #f, "SubmitButtonY="
        Print #f, "CancelButton="
        Print #f, "CancelButtonX="
        Print #f, "CancelButtonY="
        Print #f, "PicFrameInit=0"
        
        Print #f, "" 'add spacing
        Print #f, "" 'add spacing
        
        Print #f, "[Main Menu]"
        Print #f, "Background="
        Print #f, "Heading="
        Print #f, "HeadingX="
        Print #f, "HeadingY="
        Print #f, "NewButton="
        Print #f, "NewButtonX="
        Print #f, "NewButtonY="
        Print #f, "DeleteButton="
        Print #f, "DeleteButtonX="
        Print #f, "DeleteButtonY="
        Print #f, "CreditButton="
        Print #f, "CreditButtonX="
        Print #f, "CreditButtonY="
        Print #f, "ConnectButton="
        Print #f, "ConnectButtonX="
        Print #f, "ConnectButtonY="
        Print #f, "RefreshButton="
        Print #f, "RefreshButtonX="
        Print #f, "RefreshButtonY="
        Print #f, "SettingsButton="
        Print #f, "SettingsButtonX="
        Print #f, "SettingsButtonY="
        Print #f, "CancelButton="
        Print #f, "CancelButtonX="
        Print #f, "CancelButtonY="
        Print #f, "TextNameForecolor="
        Print #f, "TextNameBackcolor="
        Print #f, "TextNameX="
        Print #f, "TextNameY="
        Print #f, "TextPasswordForecolor="
        Print #f, "TextPasswordBackcolor="
        Print #f, "TextPasswordX="
        Print #f, "TextPasswordY="
        Print #f, "TextInfoForecolor="
        Print #f, "TextInfoBackcolor="
        Print #f, "TextInfoX="
        Print #f, "TextInfoY="
        Print #f, "LabelStatusForecolor="
        Print #f, "LabelStatusBackcolor="
        Print #f, "LabelStatusX="
        Print #f, "LabelStatusY="
        Print #f, "LabelStatusTransparent="
        Print #f, "PicFrameInit=0"
        
        Print #f, "" 'add spacing
        Print #f, "" 'add spacing
        
        Print #f, "[Menu Settings]"
        Print #f, "Background="
        Print #f, "Heading="
        Print #f, "HeadingX="
        Print #f, "HeadingY="
        Print #f, "SaveButton="
        Print #f, "SaveButtonX="
        Print #f, "SaveButtonY="
        Print #f, "CancelButton="
        Print #f, "CancelButtonX="
        Print #f, "CancelButtonY="
        Print #f, "TextIPForecolor="
        Print #f, "TextIPBackcolor="
        Print #f, "TextIPX="
        Print #f, "TextIPY="
        Print #f, "TextPortForecolor="
        Print #f, "TextPortBackcolor="
        Print #f, "TextPortX="
        Print #f, "TextPortY="
        Print #f, "FrameAccountForecolor="
        Print #f, "FrameAccountBackcolor="
        Print #f, "FrameAccountX="
        Print #f, "FrameAccountY="
        Print #f, "OptionOnForecolor="
        Print #f, "OptionOnBackcolor="
        Print #f, "OptionOnX="
        Print #f, "OptionOnY="
        Print #f, "OptionOffForecolor="
        Print #f, "OptionOffBackcolor="
        Print #f, "OptionOffX="
        Print #f, "OptionOffY="
        Print #f, "PicFrameInit=0"
        
        Print #f, "" 'add spacing
        Print #f, "" 'add spacing
        
        Print #f, "[New Account]"
        Print #f, "Background="
        Print #f, "Heading="
        Print #f, "HeadingX="
        Print #f, "HeadingY="
        Print #f, "ConnectButton="
        Print #f, "ConnectButtonX="
        Print #f, "ConnectButtonY="
        Print #f, "CancelButton="
        Print #f, "CancelButtonX="
        Print #f, "CancelButtonY="
        Print #f, "TextNameForecolor="
        Print #f, "TextNameBackcolor="
        Print #f, "TextNameX="
        Print #f, "TextNameY="
        Print #f, "TextPasswordForecolor="
        Print #f, "TextPasswordBackcolor="
        Print #f, "TextPasswordX="
        Print #f, "TextPasswordY="
        Print #f, "PicFrameInit=0"
        
        Print #f, "" 'add spacing
        Print #f, "" 'add spacing
        
        Print #f, "[New Character]"
        Print #f, "Background="
        Print #f, "Heading="
        Print #f, "HeadingX="
        Print #f, "HeadingY="
        Print #f, "SaveButton="
        Print #f, "SaveButtonX="
        Print #f, "SaveButtonY="
        Print #f, "CancelButton="
        Print #f, "CancelButtonX="
        Print #f, "CancelButtonY="
        Print #f, "TextNameForecolor="
        Print #f, "TextNameBackcolor="
        Print #f, "TextNameX="
        Print #f, "TextNameY="
        Print #f, "OptionMaleForecolor="
        Print #f, "OptionMaleBackcolor="
        Print #f, "OptionMaleX="
        Print #f, "OptionMaleY="
        Print #f, "OptionFemaleForecolor="
        Print #f, "OptionFemaleBackcolor="
        Print #f, "OptionFemaleX="
        Print #f, "OptionFemaleY="
        Print #f, "PictureSpriteX="
        Print #f, "PictureSpriteY="
        Print #f, "LabelSTRForecolor="
        Print #f, "LabelSTRBackcolor="
        Print #f, "LabelSTRX="
        Print #f, "LabelSTRY="
        Print #f, "LabelSTRTransparent="
        Print #f, "LabelDEFForecolor="
        Print #f, "LabelDEFBackcolor="
        Print #f, "LabelDEFX="
        Print #f, "LabelDEFY="
        Print #f, "LabelDEFTransparent="
        Print #f, "LabelMAGIForecolor="
        Print #f, "LabelMAGIBackcolor="
        Print #f, "LabelMAGIX="
        Print #f, "LabelMAGIY="
        Print #f, "LabelMAGITransparent="
        Print #f, "LabelSPEEDForecolor="
        Print #f, "LabelSPEEDBackcolor="
        Print #f, "LabelSPEEDX="
        Print #f, "LabelSPEEDY="
        Print #f, "LabelSPEEDTransparent="
        Print #f, "LabelHPForecolor="
        Print #f, "LabelHPBackcolor="
        Print #f, "LabelHPX="
        Print #f, "LabelHPY="
        Print #f, "LabelHPTransparent="
        Print #f, "LabelMPForecolor="
        Print #f, "LabelMPBackcolor="
        Print #f, "LabelMPX="
        Print #f, "LabelMPY="
        Print #f, "LabelMPTransparent="
        Print #f, "LabelSPForecolor="
        Print #f, "LabelSPBackcolor="
        Print #f, "LabelSPX="
        Print #f, "LabelSPY="
        Print #f, "LabelSPTransparent="
        Print #f, "PicFrameInit=0"
        
        Print #f, "" 'add spacing
        Print #f, "" 'add spacing
        
        Print #f, "[RealFeel Game Screen]"
        Print #f, "Background="
        Print #f, "StatsFrameme="
        Print #f, "StatsFramemeX="
        Print #f, "StatsFramemeY="
        Print #f, "EquipFrameme="
        Print #f, "EquipFramemeX="
        Print #f, "EquipFramemeY="
        Print #f, "PanelFrameme="
        Print #f, "PanelFramemeX="
        Print #f, "PanelFramemeY="
        Print #f, "UpgradeButton="
        Print #f, "UpgradeButtonX="
        Print #f, "UpgradeButtonY="
        Print #f, "InventoryButton="
        Print #f, "InventoryButtonX="
        Print #f, "InventoryButtonY="
        Print #f, "SpellButton="
        Print #f, "SpellButtonX="
        Print #f, "SpellButtonY="
        Print #f, "SideFrameme="
        Print #f, "SideFramemeX="
        Print #f, "SideFramemeY="
        Print #f, "AddFriendButton="
        Print #f, "AddFriendButtonX="
        Print #f, "AddFriendButtonY="
        Print #f, "RemoveFriendButton="
        Print #f, "RemoveFriendButtonX="
        Print #f, "RemoveFriendButtonY="
        Print #f, "CancelButton="
        Print #f, "CancelButtonX="
        Print #f, "CancelButtonY="
        Print #f, "TextChatEnterForecolor="
        Print #f, "TextChatEnterBackcolor="
        Print #f, "TextChatEnterX="
        Print #f, "TextChatEnterY="
        Print #f, "TextChatBackcolor="
        Print #f, "TextChatX="
        Print #f, "TextChatY="
        Print #f, "ListPlayersForecolor="
        Print #f, "ListPlayersBackcolor="
        Print #f, "ListPlayersX="
        Print #f, "ListPlayersY="
        Print #f, "ListFriendsForecolor="
        Print #f, "ListFriendsBackcolor="
        Print #f, "ListFriendsX="
        Print #f, "ListFriendsY="
        Print #f, "FrameOptionsForecolor="
        Print #f, "FrameOptionsBackcolor="
        Print #f, "FrameOptionsX="
        Print #f, "FrameOptionsY="
        Print #f, "PicFrameInit=0"
        
        Print #f, "" 'add spacing
        
        Print #f, "     [Admin Panel]"
        Print #f, "     Background="
        Print #f, "     PlacementX="
        Print #f, "     PlacementY="
        
        Print #f, "" 'add spacing
        
        Print #f, "     [Bank]"
        Print #f, "     Background="
        Print #f, "     PlacementX="
        Print #f, "     PlacementY="
        Print #f, "     DepositButton="
        Print #f, "     DepositButtonX="
        Print #f, "     DepositButtonY="
        Print #f, "     WithdrawButton="
        Print #f, "     WithdrawButtonX="
        Print #f, "     WithdrawButtonY="
        Print #f, "     CancelButton="
        Print #f, "     CancelButtonX="
        Print #f, "     CancelButtonY="
        
        Print #f, "" 'add spacing
        
        Print #f, "     [Inventory]"
        Print #f, "     Background="
        Print #f, "     PlacementX="
        Print #f, "     PlacementY="
        Print #f, "     UseButton="
        Print #f, "     UseButtonX="
        Print #f, "     UseButtonY="
        Print #f, "     DropButton="
        Print #f, "     DropButtonX="
        Print #f, "     DropButtonY="
        Print #f, "     CancelButton="
        Print #f, "     CancelButtonX="
        Print #f, "     CancelButtonY="
        
        Print #f, "" 'add spacing
        
        Print #f, "     [Train]"
        Print #f, "     Background="
        Print #f, "     PlacementX="
        Print #f, "     PlacementY="
        Print #f, "     Heading="
        Print #f, "     HeadingX="
        Print #f, "     HeadingY="
        Print #f, "     TrainButton="
        Print #f, "     TrainButtonX="
        Print #f, "     TrainButtonY="
        Print #f, "     CancelButton="
        Print #f, "     CancelButtonX="
        Print #f, "     CancelButtonY="
        
        Print #f, "" 'add spacing
        
        Print #f, "     [Spell]"
        Print #f, "     Background="
        Print #f, "     PlacementX="
        Print #f, "     PlacementY="

        
        Print #f, "" 'add spacing
        Print #f, "" 'add spacing
        
        Print #f, "[Information Window]"
        Print #f, "Background="
        Print #f, "LabelStatusForecolor="
        Print #f, "LabelStatusBackcolor="
        Print #f, "LabelStatusX="
        Print #f, "LabelStatusY="
        Print #f, "LabelStatusTransparent="
        Print #f, "PicFrameInit=0"
        
        Print #f, ""
        Print #f, ""
        
        Print #f, "[Shop]"
        Print #f, "Background="
        Print #f, "TradeButton="
        Print #f, "TradeButtonX="
        Print #f, "TradeButtonY="
        Print #f, "RepairButton="
        Print #f, "RepairButtonX="
        Print #f, "RepairButtonY="
        Print #f, "CancelButton="
        Print #f, "CancelButtonX="
        Print #f, "CancelButtonY="
        Print #f, "LabelNameForecolor="
        Print #f, "LabelNameBackcolor="
        Print #f, "LabelNameX="
        Print #f, "LabelNameY="
        Print #f, "LabelNameTransparent="
        Print #f, "LabelDescriptionForecolor="
        Print #f, "LabelDescriptionBackcolor="
        Print #f, "LabelDescriptionX="
        Print #f, "LabelDescriptionY="
        Print #f, "LabelDescriptionTransparent="
        Print #f, "LabelCostForecolor="
        Print #f, "LabelCostBackcolor="
        Print #f, "LabelCostX="
        Print #f, "LabelCostY="
        Print #f, "LabelCostTransparent="
        Print #f, "LabelStockForecolor="
        Print #f, "LabelStockBackcolor="
        Print #f, "LabelStockX="
        Print #f, "LabelStockY="
        Print #f, "LabelStockTransparent="
        Print #f, "LabelRestockForecolor="
        Print #f, "LabelRestockBackcolor="
        Print #f, "LabelRestockX="
        Print #f, "LabelRestockY="
        Print #f, "LabelRestockTransparent="
        Print #f, "LabelHSTRForecolor="
        Print #f, "LabelHSTRBackcolor="
        Print #f, "LabelHSTRX="
        Print #f, "LabelHSTRY="
        Print #f, "LabelHSTRTransparent="
        Print #f, "LabelHDEFForecolor="
        Print #f, "LabelHDEFBackcolor="
        Print #f, "LabelHDEFX="
        Print #f, "LabelHDEFY="
        Print #f, "LabelHDEFTransparent="
        Print #f, "LabelHMAGForecolor="
        Print #f, "LabelHMAGBackcolor="
        Print #f, "LabelHMAGX="
        Print #f, "LabelHMAGY="
        Print #f, "LabelHMAGTransparent="
        Print #f, "LabelHSPDForecolor="
        Print #f, "LabelHSPDBackcolor="
        Print #f, "LabelHSPDX="
        Print #f, "LabelHSPDY="
        Print #f, "LabelHSPDTransparent="
        Print #f, "ListTradeForecolor="
        Print #f, "ListTradeBackcolor="
        Print #f, "ListTradeX="
        Print #f, "ListTradeY="
        Print #f, "PicFrameInit=0"
    Close #f
    
    MsgBox "The config.ini file was not found! A file has been created but not completed! Please fill in all data before proceeding!"
    End
End If

Call LoadGUI
Call PicFrameCheck
End Sub

Public Sub AddText(ByVal Msg As String, ByVal Color As Integer)
Dim s As String
  
    s = vbNewLine & Msg
    frmDualSolace.txtChat.SelStart = Len(frmDualSolace.txtChat.Text)
    frmDualSolace.txtChat.SelColor = QBColor(Color)
    frmDualSolace.txtChat.SelText = s
    frmDualSolace.txtChat.SelStart = Len(frmDualSolace.txtChat.Text) - 1
End Sub

Public Sub TextAdd(ByVal Txt As TextBox, Msg As String, NewLine As Boolean)
    If NewLine Then
        Txt.Text = Txt.Text + Msg + vbCrLf
    Else
        Txt.Text = Txt.Text + Msg
    End If
    
    Txt.SelStart = Len(Txt.Text) - 1
End Sub

Function Parse(ByVal num As Long, ByVal Data As String)
Dim i As Long
Dim n As Long
Dim sChar As Long
Dim eChar As Long

    n = 0
    sChar = 1
    
    For i = 1 To Len(Data)
        If Mid(Data, i, 1) = SEP_CHAR Then
            If n = num Then
                eChar = i
                Parse = Mid(Data, sChar, eChar - sChar)
                Exit For
            End If
            
            sChar = i + 1
            n = n + 1
        End If
    Next i
End Function

Public Sub MakeRegBatch(ByVal FileName As String, ByVal FileExt As String, ByVal FilePath As String)
Dim Path As String
Dim f As Integer

Path = FilePath & FileName & ".bat"
f = FreeFile

Debug.Print "PATH NAME: " & Path
Open Path For Output As #f
    Print #f, "@ECHO OFF"
    Print #f, "ECHO :: Registering " & FileName & "." & FileExt & " for the RealFeel Engine!"
    Print #f, "COPY " & FilePath & FileName & "." & FileExt & " C:\Windows\System32 /Y"
    Print #f, "REGSVR32 " & FilePath & FileName & "." & FileExt & " / s"
    Print #f, "ECHO Registered Successfully!"
Close #f
End Sub

Public Sub RunRegBatch(ByVal FileName As String, ByVal FilePath As String, Optional KillIt As Boolean = False)
    Call Shell(FilePath & FileName & ".bat")
    If KillIt = True Then Call Kill(FilePath & FileName & ".bat")
End Sub

Public Function Exists(ByVal Path As String)
On Error Resume Next
Exists = False
If Dir(Path) <> "" Then Exists = True
End Function

Public Function CheckReg()
Dim FSys As Object, Folder As Object, FolderFiles As Object, File As Object, FileName As String

    ' Preset the return value
    CheckReg = True

    ' Preset DLL path
    DLL_PATH = GetVar(App.Path & "\data.dat", "FilePaths", "DLLs")

    ' Create the file
    Set FSys = CreateObject("Scripting.FileSystemObject")

    'Set the folder objects
    Set Folder = FSys.GetFolder(App.Path & DLL_PATH)
    Set FolderFiles = Folder.Files

    For Each File In FolderFiles
        FileName = Mid(File, Len(App.Path & DLL_PATH) + 1, (Len(File) - Len(App.Path & DLL_PATH)))
        If Not Exists("C:\Windows\System32\" & FileName) Then
            If UCase$(Right$(FileName, 3)) = "OCX" Or UCase$(Right$(FileName, 3)) = "DLL" Then
                frmRegFiles.lstRegFiles.AddItem FileName
                CheckReg = False
            End If
        End If
    Next File
    
    'Destroy the folder objects
    Set File = Nothing
    Set FolderFiles = Nothing
    Set Folder = Nothing
    Set FSys = Nothing
End Function
