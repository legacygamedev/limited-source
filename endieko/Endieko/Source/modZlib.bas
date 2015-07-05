Attribute VB_Name = "modZlib"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function ZCompress Lib "zlib.dll" Alias "compress" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Function ZUncompress Lib "zlib.dll" Alias "uncompress" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long

Public Function Compress(Data, Optional Key)
   Dim lKey As Long  'original size
   Dim sTmp As String  'string buffer
   Dim bData() As Byte  'data buffer
   Dim bRet() As Byte  'output buffer
   Dim lCSz As Long  'compressed size
   
   If TypeName(Data) = "Byte()" Then 'if given byte array data
      bData = Data  'copy to data buffer
   ElseIf TypeName(Data) = "String" Then 'if given string data
      If Len(Data) > 0 Then 'if there is data
         sTmp = Data 'copy to string buffer
         ReDim bData(Len(sTmp) - 1) 'allocate data buffer
         CopyMemory bData(0), ByVal sTmp, Len(sTmp) 'copy to data buffer
         sTmp = vbNullString 'deallocate string buffer
      End If
   End If
   If StrPtr(bData) <> 0 Then 'if data buffer contains data
      lKey = UBound(bData) + 1 'get data size
      lCSz = lKey + (lKey * 0.01) + 12 'estimate compressed size
      ReDim bRet(lCSz - 1) 'allocate output buffer
      Call ZCompress(bRet(0), lCSz, bData(0), lKey) 'compress data (lCSz returns actual size)
      ReDim Preserve bRet(lCSz - 1) 'resize output buffer to actual size
      Erase bData 'deallocate data buffer
      If IsMissing(Key) Then 'if Key variable not supplied
         ReDim bData(lCSz + 3) 'allocate data buffer
         CopyMemory bData(0), lKey, 4 'copy key to buffer
         CopyMemory bData(4), bRet(0), lCSz 'copy data to buffer
         Erase bRet 'deallocate output buffer
         bRet = bData 'copy to output buffer
         Erase bData 'deallocate data buffer
      Else 'Key variable is supplied
         Key = lKey 'set Key variable
      End If
      If TypeName(Data) = "Byte()" Then 'if given byte array data
         Compress = bRet 'return output buffer
      ElseIf TypeName(Data) = "String" Then 'if given string data
         sTmp = Space(UBound(bRet) + 1) 'allocate string buffer
         CopyMemory ByVal sTmp, bRet(0), UBound(bRet) + 1 'copy to string buffer
         Compress = sTmp 'return string buffer
         sTmp = vbNullString 'deallocate string buffer
      End If
      Erase bRet 'deallocate output buffer
   End If
End Function

Public Function Uncompress(Data, Optional ByVal Key)
   Dim lKey As Long  'original size
   Dim sTmp As String  'string buffer
   Dim bData() As Byte  'data buffer
   Dim bRet() As Byte  'output buffer
   Dim lCSz As Long  'compressed size
  
   If TypeName(Data) = "Byte()" Then 'if given byte array data
      bData = Data 'copy to data buffer
   ElseIf TypeName(Data) = "String" Then 'if given string data
      If Len(Data) > 0 Then 'if there is data
         sTmp = Data 'copy to string buffer
         ReDim bData(Len(sTmp) - 1) 'allocate data buffer
         CopyMemory bData(0), ByVal sTmp, Len(sTmp) 'copy to data buffer
         sTmp = vbNullString 'deallocate string buffer
      End If
   End If
   If StrPtr(bData) <> 0 Then 'if there is data
      If IsMissing(Key) Then 'if Key variable not supplied
         lCSz = UBound(bData) - 3 'get actual data size
         CopyMemory lKey, bData(0), 4 'copy key value to key
         ReDim bRet(lCSz - 1) 'allocate output buffer
         CopyMemory bRet(0), bData(4), lCSz 'copy data to output buffer
         Erase bData 'deallocate data buffer
         bData = bRet 'copy to data buffer
         Erase bRet 'deallocate output buffer
      Else 'Key variable is supplied
         lCSz = UBound(bData) + 1 'get data size
         lKey = Key 'get Key
      End If
      ReDim bRet(lKey - 1) 'allocate output buffer
      Call ZUncompress(bRet(0), lKey, bData(0), lCSz) 'decompress to output buffer
      Erase bData 'deallocate data buffer
      If TypeName(Data) = "Byte()" Then 'if given byte array data
         Uncompress = bRet 'return output buffer
      ElseIf TypeName(Data) = "String" Then 'if given string data
         sTmp = Space(lKey) 'allocate string buffer
         CopyMemory ByVal sTmp, bRet(0), lKey 'copy to string buffer
         Uncompress = sTmp 'return string buffer
         sTmp = vbNullString 'deallocate string buffer
      End If
      Erase bRet 'deallocate return buffer
   End If
End Function
