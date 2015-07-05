Attribute VB_Name = "modAES"
Option Explicit

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Global Authenticated As Boolean
Global TempEnc As String
Global DCount As Integer
Global RandomKey As String

Public Const KeytxtAlias = "bmso;jt54#@TK#$mg43gj34og;3j4g40"
Public Const KeySendAuth = "aS?dj3mfsdf923rWFfj3#rrj3oj24@#4"
Public Const KeyReplyStr = "hdkgjg$#@GKwegj34tW$TK$3tj34gjj4"
Public Const KeyAuthRepl = "sajlkdjaSDKAsdhas;ldkaj$#K3lo4j3"
Public Const RandStrEncr = "GKSDfljsdkvsdnl4yu3pwrjn324@#$@,"
Sub Pause(Interval)
Dim Current
    
Current = Timer
Do While Timer - Current < Val(Interval)
    DoEvents
Loop
End Sub
Public Function Encrypt(Password As String, Text As String) As String
On Error GoTo errorhandler
GoSub begin

errorhandler:
Exit Function

begin:
    Dim oTest As AES, sTemp As String, bytIn() As Byte
    Dim bytOut() As Byte, bytPassword() As Byte, bytClear() As Byte
    Dim lCount As Long, lLength As Long
    
    If Text = "" Or Password = "" Then Exit Function
    
    Set oTest = New AES
    
    bytIn = Text
    bytPassword = Password

    bytOut = oTest.EncryptData(bytIn, bytPassword)

    sTemp = ""
    For lCount = 0 To UBound(bytOut)
        sTemp = sTemp & Right("0" & Hex(bytOut(lCount)), 2)
    Next
    Encrypt = sTemp
End Function

Public Function Decrypt(Password As String, EncryptedString As String) As String
On Error GoTo errorhandler
GoSub begin

errorhandler:
Decrypt = "[DM] Error"
Exit Function

begin:
    If EncryptedString = "" Or Password = "" Then Exit Function
    
    Dim oTest As AES, sTemp As String, bytIn() As Byte, bytOut() As Byte, bytPassword() As Byte, bytClear() As Byte, lCount As Long, lLength As Long, DC As String
    Set oTest = New AES
    
    bytIn = EncryptedString
    bytPassword = Password
    sTemp = EncryptedString
    
    lLength = Len(sTemp)
    ReDim bytOut((lLength \ 2) - 1)

    For lCount = 1 To lLength Step 2
        bytOut(lCount \ 2) = CByte("&H" & Mid(sTemp, lCount, 2))
    Next

    bytClear = oTest.DecryptData(bytOut, bytPassword)
    DC = bytClear
    If DC = vbNullString Then Decrypt = "[DM] Error" Else Decrypt = bytClear
End Function
