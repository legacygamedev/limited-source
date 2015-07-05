Attribute VB_Name = "modDatabase"
'   This file is part of the Cerberus Engine 2nd Edition.
'
'    The Cerberus Engine 2nd Edition is free software; you can redistribute it
'    and/or modify it under the terms of the GNU General Public License as
'    published by the Free Software Foundation; either version 2 of the License,
'    or (at your option) any later version.
'
'    Cerberus 2nd Edition is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Cerberus 2nd Edition; if not, write to the Free Software
'    Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301  USA

Option Explicit

Public Function GetVar(File As String, Header As String, Var As String) As String
Dim sSpaces As String   ' Max string length
Dim szReturn As String  ' Return default value if not found
  
    szReturn = ""
  
    sSpaces = Space(5000)
  
    Call GetPrivateProfileString(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
  
    GetVar = RTrim(sSpaces)
    GetVar = Left(GetVar, Len(GetVar) - 1)
End Function

Public Function FileExist(ByVal FileName As String, Optional RAW As Boolean = False) As Boolean
    
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

Public Sub PutVar(File As String, Header As String, Var As String, Value As String)
    Call WritePrivateProfileString(Header, Var, Value, File)
End Sub

Public Sub SaveLocalMap(ByVal MapNum As Long)
Dim FileName As String
Dim f As Long

    FileName = App.Path & MAP_PATH & "map" & MapNum & MAP_EXT
            
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , SaveMap
    Close #f
End Sub

Public Sub LoadMap(ByVal MapNum As Long)
Dim FileName As String
Dim f As Long

    FileName = App.Path & MAP_PATH & "map" & MapNum & MAP_EXT
        
    f = FreeFile
    Open FileName For Binary As #f
        Get #f, , SaveMap
    Close #f
End Sub

Public Function GetMapRevision(ByVal MapNum As Long) As Long
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
