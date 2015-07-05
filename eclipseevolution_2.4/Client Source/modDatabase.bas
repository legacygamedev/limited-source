Attribute VB_Name = "modDatabase"
Option Explicit
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

Private Declare Function FindFirstFile Lib "KERNEL32" Alias "FindFirstFileA" _
        (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "KERNEL32" (ByVal hFindFile As Long) As Long
Private Declare Function FindNextFile Lib "KERNEL32" Alias "FindNextFileA" _
        (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long

'Kernal API Declarations
'   originally in BitmapUtils, but made public for use in CompressPackets.
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal ByteLen As Long)

Public SOffsetX As Integer
Public SOffsetY As Integer


Sub AddLog(ByVal Text As String)

  Dim FileName As String
  Dim f As Long

    If Trim$(Command) = "-debug" Then
        If frmDebug.Visible = False Then
            frmDebug.Visible = True
        End If

        FileName = App.Path & "\debug.txt"

        If Not FileExist("debug.txt") Then
            f = FreeFile
            Open FileName For Output As #f
            Close #f
        End If

        f = FreeFile
        Open FileName For Append As #f
        Print #f, Time & ": " & Text
        Close #f
    End If

    Exit Sub

End Sub

Function FileExist(FileName As String) As Boolean

    On Error GoTo ErrorHandler
    ' get the attributes and ensure that it isn't a directory
    FileExist = (GetAttr(App.Path & "\" & FileName) And vbDirectory) = 0
    ErrorHandler:
    ' if an error occurs, this function returns False

End Function

Function GetMapRevision(ByVal MapNum As Long) As Long

    GetMapRevision = Map(MapNum).Revision

End Function


'Returns true if the tile is a roof tile and the player is under that section of roof
Function IsTileRoof(ByVal x As Integer, ByVal y As Integer) As Boolean

    'you should have seen the original function. Logic statement notfun :( -Pickle
  Dim IsRoof As Boolean

    If Map(GetPlayerMap(MyIndex)).Tile(x, y).Type = TILE_TYPE_ROOF Or Map(GetPlayerMap(MyIndex)).Tile(x, y).Type = TILE_TYPE_ROOFBLOCK Then 'If the tile is a roof or a roofblock
        If Map(GetPlayerMap(MyIndex)).Tile(x, y).String1 = Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).String1 Then 'If the roof ID is the same
            IsTileRoof = True
            Exit Function
        End If

    End If

    IsTileRoof = False

End Function

Function ListBGS(ByVal sStartDir As String)

    Exit Function
  Dim lpFindFileData As WIN32_FIND_DATA
  Dim  lFileHdl  As Long
  Dim sTemp As String
  Dim  sTemp2 As String
  Dim  lRet As Long
  Dim  iLastIndex  As Integer
  Dim strPath As String

    On Error Resume Next

    If right$(sStartDir, 1) <> "\" Then sStartDir = sStartDir & "\"

    sStartDir = sStartDir & "*.*"

    lFileHdl = FindFirstFile(sStartDir, lpFindFileData)

    If lFileHdl <> -1 Then

        Do Until lRet = ERROR_NO_MORE_FILES
            strPath = Left$(sStartDir, Len(sStartDir) - 4) & "\"

            If (lpFindFileData.dwFileAttributes And FILE_ATTRIBUTE_NORMAL) = vbNormal Then
                sTemp = StrConv(StripTerminator(lpFindFileData.cFileName), vbProperCase)
            End If

            lRet = FindNextFile(lFileHdl, lpFindFileData)
            If lRet = 0 Then Exit Do
        Loop

    End If
    lRet = FindClose(lFileHdl)

End Function

Function ListMusic(ByVal sStartDir As String)

  Dim lpFindFileData As WIN32_FIND_DATA
  Dim  lFileHdl  As Long
  Dim sTemp As String
  Dim  lRet As Long
  Dim strPath As String

    On Error Resume Next

    If right$(sStartDir, 1) <> "\" Then sStartDir = sStartDir & "\"
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
            If lRet = 0 Then Exit Do
        Loop

    End If
    lRet = FindClose(lFileHdl)

End Function

Function ListSounds(ByVal sStartDir As String, ByVal Form As Long)

  Dim lpFindFileData As WIN32_FIND_DATA
  Dim  lFileHdl  As Long
  Dim sTemp As String
  Dim  lRet As Long
  Dim strPath As String

    On Error Resume Next

    If right$(sStartDir, 1) <> "\" Then sStartDir = sStartDir & "\"

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
            If lRet = 0 Then Exit Do
        Loop

    End If
    lRet = FindClose(lFileHdl)

End Function

Sub LoadMap(ByVal MapNum As Long)

  Dim FileName As String
  Dim f As Long

    FileName = App.Path & "\maps\map" & MapNum & ".dat"

    If FileExist("maps\map" & MapNum & ".dat") = False Then Exit Sub
    f = FreeFile
    Open FileName For Binary As #f
    Get #f, , Map(MapNum)
    Close #f

End Sub

Sub MoveForm(f As Form, Button As Integer, Shift As Integer, x As Single, y As Single)

  Dim GlobalX As Integer
  Dim GlobalY As Integer

    GlobalX = f.Left
    GlobalY = f.Top

    If Button = 1 Then
        f.Left = GlobalX + x - SOffsetX
        f.Top = GlobalY + y - SOffsetY
    End If

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

Sub SaveLocalMap(ByVal MapNum As Long)

  Dim FileName As String
  Dim f As Long

    FileName = App.Path & "\maps\map" & MapNum & ".dat"

    f = FreeFile
    Open FileName For Binary As #f
    Put #f, , Map(MapNum)
    Close #f

End Sub

Function StripTerminator(ByVal strString As String) As String

  Dim intZeroPos As Integer

    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
     Else
        StripTerminator = strString
    End If

End Function

