Attribute VB_Name = "modSystem"
Option Explicit

' Used for listing sounds
Public Const MAX_PATH = 260
Private Const ERROR_NO_MORE_FILES = 18
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

' Used for listing sounds
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetTickCount Lib "KERNEL32" () As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)

' Used for listing sounds
Private Declare Function FindFirstFile Lib "KERNEL32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "KERNEL32" (ByVal hFindFile As Long) As Long
Private Declare Function FindNextFile Lib "KERNEL32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
        
' Playing sound
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

' originally in BitmapUtils, but made public for use in CompressPackets.
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal ByteLen As Long)

Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer

    intZeroPos = InStr(strString, Chr$(0))

    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

Function FileExists(ByVal FileName As String) As Boolean
    On Error Resume Next

    FileExists = (GetAttr(App.Path & "\" & FileName) And vbDirectory) = 0
End Function

Sub ListMusic(ByVal sStartDir As String)
    Dim lpFindFileData As WIN32_FIND_DATA
    Dim lFileHdl  As Long
    Dim sTemp As String
    Dim lRet As Long
    Dim strPath As String

    On Error Resume Next

    If Right$(sStartDir, 1) <> "\" Then
        sStartDir = sStartDir & "\"
    End If

    frmMapProperties.lstMusic.Clear

    frmMapProperties.lstMusic.addItem "None", 0

    sStartDir = sStartDir & "*.*"

    lFileHdl = FindFirstFile(sStartDir, lpFindFileData)

    If lFileHdl <> -1 Then
        Do Until lRet = ERROR_NO_MORE_FILES
            strPath = Left$(sStartDir, Len(sStartDir) - 4) & "\"

            If (lpFindFileData.dwFileAttributes And FILE_ATTRIBUTE_NORMAL) = vbNormal Then
                sTemp = StrConv(StripTerminator(lpFindFileData.cFileName), vbProperCase)
                frmMapProperties.lstMusic.addItem sTemp
            End If

            lRet = FindNextFile(lFileHdl, lpFindFileData)

            If lRet = 0 Then
                Exit Do
            End If
        Loop
    End If

    lRet = FindClose(lFileHdl)
End Sub

' This sub-routine still needs to be optimized. [Mellowz]
Sub ListBGS(ByVal sStartDir As String)
    Exit Sub
    Dim lpFindFileData As WIN32_FIND_DATA, lFileHdl  As Long
    Dim sTemp As String, sTemp2 As String, lRet As Long, iLastIndex  As Integer
    Dim strPath As String

    On Error Resume Next

    If Right$(sStartDir, 1) <> "\" Then
        sStartDir = sStartDir & "\"
    End If

    sStartDir = sStartDir & "*.*"

    lFileHdl = FindFirstFile(sStartDir, lpFindFileData)

    If lFileHdl <> -1 Then
        Do Until lRet = ERROR_NO_MORE_FILES
            strPath = Left$(sStartDir, Len(sStartDir) - 4) & "\"
            If (lpFindFileData.dwFileAttributes And FILE_ATTRIBUTE_NORMAL) = vbNormal Then
                sTemp = StrConv(StripTerminator(lpFindFileData.cFileName), vbProperCase)
            End If
            lRet = FindNextFile(lFileHdl, lpFindFileData)
            If lRet = 0 Then
                Exit Do
            End If
        Loop
    End If
    lRet = FindClose(lFileHdl)
End Sub

' This sub-routine still needs to be optimized. [Mellowz]
Sub ListSounds(ByVal sStartDir As String, ByVal Form As Long)
    Dim lpFindFileData As WIN32_FIND_DATA
    Dim lFileHdl  As Long
    Dim sTemp As String
    Dim lRet As Long
    Dim strPath As String

    On Error Resume Next

    If Right$(sStartDir, 1) <> "\" Then
        sStartDir = sStartDir & "\"
    End If
    If Form = 1 Then
        frmSound.lstSound.Clear
    ElseIf Form = 2 Then
        frmNotice.lstSound.Clear
    End If

    sStartDir = sStartDir & "*.*"

    lFileHdl = FindFirstFile(sStartDir, lpFindFileData)

    If lFileHdl <> -1 Then
        Do Until lRet = ERROR_NO_MORE_FILES
            strPath = Left$(sStartDir, Len(sStartDir) - 4) & "\"
            If (lpFindFileData.dwFileAttributes And FILE_ATTRIBUTE_NORMAL) = vbNormal Then
                sTemp = StrConv(StripTerminator(lpFindFileData.cFileName), vbProperCase)
                If Form = 1 Then
                    frmSound.lstSound.addItem sTemp
                ElseIf Form = 2 Then
                    frmNotice.lstSound.addItem sTemp
                End If
            End If
            lRet = FindNextFile(lFileHdl, lpFindFileData)
            If lRet = 0 Then
                Exit Do
            End If
        Loop
    End If
    lRet = FindClose(lFileHdl)
End Sub

Sub MovePicture(PB As PictureBox, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim GlobalX As Integer
    Dim GlobalY As Integer

    GlobalX = PB.Left
    GlobalY = PB.Top

    If Button = 1 Then
        PB.Left = GlobalX + x - SOffsetX
        PB.Top = GlobalY + y - SOffsetY
    End If
End Sub

Public Function IsInArray(CtrlArray As Variant, Index As Integer) As Boolean
    On Error Resume Next
    Dim x As String
    x = CtrlArray(Index).name
    If Err.Number = 0 Then IsInArray = True
End Function

