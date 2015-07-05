Attribute VB_Name = "modMain"
Option Explicit

Public Const CONFIG_FILE As String = "\config.aph"
Public KEYWORD As String

' (needs to match key in client)
Public Const DEFAULT_KEY As String = "admin"

Public Config As ConfigRec

Type ConfigRec
    Password As String * 10
    IP As String * 15
    Port As Integer
End Type

Public Function FileExist(ByVal FileName As String) As Boolean
    FileExist = LenB(Dir$(FileName)) > 0
End Function

Public Function Encryption(CodeKey As String, DataIn As String) As String
On Error Resume Next
Dim lonDataPtr As Long
Dim strDataOut As String
Dim intXOrValue1 As Integer
Dim intXOrValue2 As Integer

    For lonDataPtr = 1 To Len(DataIn)
    
        intXOrValue1 = Asc(Mid$(DataIn, lonDataPtr, 1))
        intXOrValue2 = Asc(Mid$(CodeKey, ((lonDataPtr Mod Len(CodeKey)) + 1), 1))
        
        strDataOut = strDataOut + Chr$(intXOrValue1 Xor intXOrValue2)
        
    Next

    Encryption = strDataOut
   
End Function
