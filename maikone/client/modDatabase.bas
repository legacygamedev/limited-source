Attribute VB_Name = "modDatabase"
Option Explicit

Public Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)
Public Declare Function GetPrivateProfileString& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal ReturnedString$, ByVal RSSize&, ByVal FileName$)

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

    If Trim(Command) = "-debug" Then
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
Dim TmpMap As MapRec

    FileName = App.Path & "\maps\map" & MapNum & ".dat"
        
    f = FreeFile
    Open FileName For Binary As #f
        Get #f, , TmpMap
    Close #f
        
    GetMapRevision = TmpMap.Revision
End Function

Sub PutVar(File As String, Header As String, Var As String, Value As String)
    If Trim(Value) = "0" Or Trim(Value) = "" Then
        If ExistVar(File, Header, Var) Then
            Call DelVar(File, Header, Var)
        End If
    Else
        Call WritePrivateProfileString(Header, Var, Value, File)
    End If
End Sub

Function ExistVar(File As String, Header As String, Var As String) As Boolean
Dim sSpaces As String
Dim szReturn As String
  
    szReturn = "somethingwierdheresothatitcouldntbeguessed"
  
    sSpaces = Space(5000)
  
    Call GetPrivateProfileString(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
  
    If RTrim(sSpaces) = "somethingwierdheresothatitcouldntbeguessed" Then
        ExistVar = False
    Else
        ExistVar = True
    End If
End Function

Public Sub DelVar(sFileName As String, sSection As String, sKey As String)

   If Len(Trim(sKey)) <> 0 Then
      WritePrivateProfileString sSection, sKey, _
         vbNullString, sFileName
   Else
      WritePrivateProfileString _
         sSection, sKey, vbNullString, sFileName
   End If
End Sub
Function GetVar(File As String, Header As String, Var As String) As String
Dim sSpaces As String   ' Max string length
Dim szReturn As String  ' Return default value if not found
 
    szReturn = ""
 
    sSpaces = Space(5000)
 
    Call GetPrivateProfileString(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
 
    GetVar = RTrim(sSpaces)
    GetVar = Left(GetVar, Len(GetVar) - 1)
End Function
