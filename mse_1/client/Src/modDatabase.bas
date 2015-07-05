Attribute VB_Name = "modDatabase"
Option Explicit

Public Function GetVar(File As String, Header As String, Var As String) As String
Dim sSpaces As String   ' Max string length
Dim szReturn As String  ' Return default value if not found
  
    If FileExist(File, True) = False Then
        GetVar = vbNullString
    End If
    
    szReturn = vbNullString
  
    sSpaces = Space(5000)
  
    Call GetPrivateProfileString(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
  
    GetVar = RTrim(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

Public Sub PutVar(File As String, Header As String, Var As String, Value As String)
    Call WritePrivateProfileString(Header, Var, Value, File)
End Sub

Public Function FileExist(ByVal FileName As String, Optional RAW As Boolean = False) As Boolean
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 07/12/2005  Shannara   Optimized function.
'****************************************************************
    
    If RAW = False Then
        If LenB(Dir$(App.Path & "\" & FileName)) = 0 Then
            FileExist = False
            Exit Function
        Else
            FileExist = True
            Exit Function
        End If
    Else
        If LenB(Dir$(FileName)) = 0 Then
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
