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

    If Trim(Command) = "-debug" Then
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
