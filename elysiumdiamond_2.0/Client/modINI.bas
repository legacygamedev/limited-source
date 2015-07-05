Attribute VB_Name = "ModINI"
Option Explicit
Public Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)
Public Declare Function GetPrivateProfileString& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal ReturnedString$, ByVal RSSize&, ByVal FileName$)

Public Sub WriteINI(INISection As String, INIKey As String, INIValue As String, INIFile As String)
    Call WritePrivateProfileString(INISection, INIKey, INIValue, INIFile)
End Sub

Public Function ReadINI(INISection As String, INIKey As String, INIFile As String) As String
    Dim StringBuffer As String
    Dim StringBufferSize As Long
    
    StringBuffer = Space$(255)
    StringBufferSize = Len(StringBuffer)
    
    StringBufferSize = GetPrivateProfileString(INISection, INIKey, "", StringBuffer, StringBufferSize, INIFile)
    
    If StringBufferSize > 0 Then
        ReadINI = Left$(StringBuffer, StringBufferSize)
    Else
        ReadINI = ""
    End If
End Function

Function GetVar(File As String, Header As String, Var As String) As String
Dim sSpaces As String   ' Max string length
Dim szReturn As String  ' Return default value if not found
  
    szReturn = ""
  
    sSpaces = Space(5000)
  
    Call GetPrivateProfileString(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
  
    GetVar = RTrim(sSpaces)
    GetVar = Left(GetVar, Len(GetVar) - 1)
End Function

' Taken from the server
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

Sub PutVar(File As String, Header As String, Var As String, Value As String)
    If Trim(Value) = "0" Or Trim(Value) = "" Then
        If ExistVar(File, Header, Var) Then
            Call DelVar(File, Header, Var)
        End If
    Else
        Call WritePrivateProfileString(Header, Var, Value, File)
    End If
End Sub

Public Sub DelVar(sFileName As String, sSection As String, sKey As String)

   If Len(Trim(sKey)) <> 0 Then
      WritePrivateProfileString sSection, sKey, _
         vbNullString, sFileName
   Else
      WritePrivateProfileString _
         sSection, sKey, vbNullString, sFileName
   End If
End Sub
