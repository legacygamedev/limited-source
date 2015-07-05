Attribute VB_Name = "modDatabase"
'   Copyright (c) 2006 Joshua Bendig
'   This file is part of Asgard.
'
'    Asgard is free software; you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation; either version 2 of the License, or
'    (at your option) any later version.
'
'    Asgard is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Asgard; if not, write to the Free Software
'    Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301  USA

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
    
Public SOffsetX As Integer
Public SOffsetY As Integer

Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

Function FileExist(ByVal filename As String) As Boolean
    If Dir(App.Path & "\" & filename) = "" Then
        FileExist = False
    Else
        FileExist = True
    End If
End Function

Sub AddLog(ByVal Text As String)
Dim filename As String
Dim f As Long

    If Trim(Command) = "-debug" Then
        If frmDebug.Visible = False Then
            frmDebug.Visible = True
        End If
        
        filename = App.Path & "\debug.txt"
    
        If Not FileExist("debug.txt") Then
            f = FreeFile
            Open filename For Output As #f
            Close #f
        End If
    
        f = FreeFile
        Open filename For Append As #f
            Print #f, Time & ": " & Text
        Close #f
    End If
End Sub

Sub SaveLocalMap(ByVal MapNum As Long)
Dim filename As String
Dim f As Long

    filename = App.Path & "\maps\map" & MapNum & ".dat"
                            
    f = FreeFile
    Open filename For Binary As #f
        Put #f, , Map(MapNum)
    Close #f
End Sub

Sub LoadMap(ByVal MapNum As Long)
Dim filename As String
Dim f As Long

    filename = App.Path & "\maps\map" & MapNum & ".dat"
        
    If FileExist("maps\map" & MapNum & ".dat") = False Then Exit Sub
    f = FreeFile
    Open filename For Binary As #f
        Get #f, , Map(MapNum)
    Close #f
End Sub

Function GetMapRevision(ByVal MapNum As Long) As Long
    GetMapRevision = Map(MapNum).Revision
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
                        If Right$(sTemp, 4) = ".mp3" Then
                            frmMapProperties.lstMusic.AddItem sTemp
                        End If
                        If Right$(sTemp, 4) = ".wma" Then
                            frmMapProperties.lstMusic.AddItem sTemp
                        End If
                        If Right$(sTemp, 4) = ".ogg" Then
                            frmMapProperties.lstMusic.AddItem sTemp
                        End If
                        If Right$(sTemp, 4) = ".wav" Then
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
