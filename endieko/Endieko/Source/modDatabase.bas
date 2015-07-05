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

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" _
    (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" _
    (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
    
Public TmpMap As MapRec

Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

Function FileExist(ByVal FileName As String) As Boolean
    If Dir(App.Path & "\" & FileName) = "" Then
        FileExist = False
    Else
        FileExist = True
    End If
End Function

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
End Sub

Sub SaveLocalMap(ByVal MapNum As Long)
Dim FileName As String
Dim f As Long

    FileName = App.Path & "\maps\map" & MapNum & ".dat"
            
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , SaveMap
    Close #f
End Sub

Sub LoadMap(ByVal MapNum As Long)
Dim FileName As String
Dim f As Long

    FileName = App.Path & "\maps\map" & MapNum & ".dat"
        
    f = FreeFile
    Open FileName For Binary As #f
        Get #f, , SaveMap
    Close #f
End Sub

Function GetMapRevision(ByVal MapNum As Long) As Long
Dim FileName As String
Dim f As Long


    FileName = App.Path & "\maps\map" & MapNum & ".dat"
        
    f = FreeFile
    Open FileName For Binary As #f
        Get #f, , TmpMap
    Close #f
        
    GetMapRevision = TmpMap.Revision
End Function

Function ListMusic(ByVal sStartDir As String)
    Dim lpFindFileData As WIN32_FIND_DATA, lFileHdl  As Long
    Dim sTemp As String, sTemp2 As String, lRet As Long, iLastIndex  As Integer
    Dim strPath As String
    
    On Error Resume Next
    
    If Right$(sStartDir, 1) <> "\" Then sStartDir = sStartDir & "\"
    frmMapProperties.lstMusic.Clear
    
    frmMapProperties.lstMusic.AddItem "None", 0
    
    sStartDir = sStartDir & "*.*"
    
    lFileHdl = FindFirstFile(sStartDir, lpFindFileData)
    
    If lFileHdl <> -1 Then
        Do Until lRet = ERROR_NO_MORE_FILES
                strPath = Left$(sStartDir, Len(sStartDir) - 4) & "\"
                    If (lpFindFileData.dwFileAttributes And FILE_ATTRIBUTE_NORMAL) = vbNormal Then
                        sTemp = StrConv(StripTerminator(lpFindFileData.cFileName), vbProperCase)
                        If Right$(sTemp, 4) = ".mid" Then
                            frmMapProperties.lstMusic.AddItem sTemp
                        End If
                    End If
                lRet = FindNextFile(lFileHdl, lpFindFileData)
            If lRet = 0 Then Exit Do
        Loop
    End If
    lRet = FindClose(lFileHdl)
End Function

Function ListSounds(ByVal sStartDir As String, ByVal Form As Long)
    Dim lpFindFileData As WIN32_FIND_DATA, lFileHdl  As Long
    Dim sTemp As String, sTemp2 As String, lRet As Long, iLastIndex  As Integer
    Dim strPath As String
    
    On Error Resume Next
    
    If Right$(sStartDir, 1) <> "\" Then sStartDir = sStartDir & "\"
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
                        If Right$(sTemp, 4) = ".wav" Then
                            If Form = 1 Then
                                frmSound.lstSound.AddItem sTemp
                            ElseIf Form = 2 Then
                                frmNotice.lstSound.AddItem sTemp
                            End If
                        End If
                    End If
                lRet = FindNextFile(lFileHdl, lpFindFileData)
            If lRet = 0 Then Exit Do
        Loop
    End If
    lRet = FindClose(lFileHdl)
End Function

Public Function GetHDSerial(Optional ByVal DriveLetter As String) As Long
    Dim fso As Object, Drv As Object, DriveSerial As Long
   
    'Create a FileSystemObject object
    Set fso = CreateObject("Scripting.FileSystemObject")
   
    'Assign the current drive letter if not specified
    If DriveLetter <> "" Then
        Set Drv = fso.GetDrive(DriveLetter)
    Else
        Set Drv = fso.GetDrive(fso.GetDriveName(App.Path))
    End If

    With Drv
        If .IsReady Then
             DriveSerial = Abs(.SerialNumber)
        Else    '"Drive Not Ready!"
             DriveSerial = -1
        End If
    End With
   
    'Clean up
    Set Drv = Nothing
    Set fso = Nothing
   
    GetHDSerial = DriveSerial
End Function

Public Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single
    Randomize Timer
    RandomNumber = Int((UpperBound - LowerBound + 1) * Rnd + LowerBound)
End Function

Public Sub GradObj(Obj As Object, Optional Color1 As Long = -1, Optional Color2 As Long = vbWhite)
Dim Y As Long

    If Color1 = -1 Then
        Color1 = vbWhite
        Color2 = vbGreen
    End If

    
    Dim R As Integer
    Dim G As Integer
    Dim B As Integer
    Dim R2 As Integer
    Dim G2 As Integer
    Dim B2 As Integer
    Dim Col As Integer
    Dim ValueRed As Integer
    Dim ValueGreen As Integer
    Dim ValueBlue As Integer
    Obj.Cls
    On Error Resume Next
    Obj.AutoRedraw = True
    Obj.DrawStyle = vbInsideSolid
    Obj.DrawMode = vbCopyPen
    Obj.ScaleMode = vbPixels
    Obj.DrawWidth = 10
    Obj.ScaleHeight = 58
    
    Col = (Color1 And 255)
    R = Col And 255
    Col = Int(Color1 / 256)
    G = Col And 255
    Col = Int(Color1 / 65536)
    B = Col And 255
    Col = (Color2 And 255)
    R2 = Col And 255
    Col = Int(Color2 / 256)
    G2 = Col And 255
    Col = Int(Color2 / 65536)
    B2 = Col And 255
    ValueRed = Abs(R - R2) / Obj.ScaleHeight
    ValueGreen = Abs(G - G2) / Obj.ScaleHeight
    ValueBlue = Abs(B - B2) / Obj.ScaleHeight
    If R2 < R Then ValueRed = -ValueRed
    If G2 < G Then ValueGreen = -ValueGreen
    If B2 < B Then ValueBlue = -ValueBlue


    For Y = 0 To Obj.ScaleHeight
        R2 = R + ValueRed * Y
        G2 = G + ValueGreen * Y
        B2 = B + ValueBlue * Y
        Obj.Line (0, Y)-(Obj.ScaleWidth, Y), RGB(R2, G2, B2)
    Next Y

    Obj.AutoRedraw = False
End Sub
