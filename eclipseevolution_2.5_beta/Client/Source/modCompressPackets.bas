Attribute VB_Name = "modCompressPackets"
Option Explicit

' The subs in this module are used for compressing and
'   decompressing packet data. It uses zLib, which is
'   also used in the BitmapUtils class.

Private Const Z_OK As Integer = 0

'int compress (Bytef *dest, uLongf *destLen, const Bytef *source, uLong sourceLen);

Private Declare _
  Function zlibCompress _
  Lib "zlib.dll" _
  Alias "compress" _
  ( _
    ByRef dest As Any, _
    ByRef destLen As Long, _
    ByRef Source As Any, _
    ByVal sourceLen As Long _
  ) As Long
  
  'int uncompress (Bytef *dest, uLongf *destLen, const Bytef *source, uLong sourceLen);

Private Declare _
  Function zlibUncompress _
  Lib "zlib.dll" _
  Alias "uncompress" _
  ( _
    ByRef dest As Any, _
    ByRef destLen As Long, _
    ByRef Source As Any, _
    ByVal sourceLen As Long _
  ) As Long


'********************
' compressPacket
'--------------------
'returns a byte array
'of the compressed packet
'thanks Matthew Hall
'************************
Public Function compressPacket(ByVal packet As String)
    Dim byteData() As Byte
    Dim tempData() As Byte
    Dim dataLen As Long
    
    byteData = StrConv(packet, vbFromUnicode)
    dataLen = 0&
    
    On Error Resume Next
    nDataLen = UBound(Data) + 1&
    On Error GoTo 0
    
    ' Make a buffer for the compressed data
    Dim cDataLen&
    cDataLen = CLng((dataLen + 12&) * 1.1)
    ReDim tempData(cDestLen - 1&)
    
    If zlibCompress(tempData, cDataLen, byteData, dataLen) = Z_OK Then
        compressPacket = byteData
    Else
        compressPacket = -1
    End If
    
End Function
