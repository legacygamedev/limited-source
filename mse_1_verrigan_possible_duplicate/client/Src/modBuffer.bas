Attribute VB_Name = "modBuffer"
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 09/13/2005  Verrigan   Created module to handle adding and
'*                        reading data to/from a 'buffer'.
'****************************************************************
Option Explicit
'One thing to note, always initialize your byte array to "".
'i.e. dBytes = ""
'This sets the UBound of the byte array to -1, allowing us
'to have a size of 0. If the array is not initialized, or
'you have emptied the array, aLen will cause a RT error.
Public Function aLen(ByRef dBytes() As Byte) As Integer
  aLen = (UBound(dBytes) - LBound(dBytes)) + 1
End Function
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 09/13/2005  Verrigan   PrefixBuffer() function created so you
'*                        can add data to the beginning of a
'*                        buffer.
'****************************************************************
Public Function PrefixBuffer(ByRef Buffer() As Byte, ByVal StartAddr As Long, ByVal ByteLen As Long) As Byte()
  Dim tBytes() As Byte
  
  ReDim tBytes(UBound(Buffer) + ByteLen)
  
  Call CopyMemory(tBytes(0), ByVal StartAddr, ByteLen)
  If aLen(Buffer) > 0 Then
    Call CopyMemory(tBytes(ByteLen), Buffer(0), aLen(Buffer))
  End If
  
  PrefixBuffer = tBytes
End Function
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 09/13/2005  Verrigan   AddToBuffer() function created so you
'*                        can concatenate byte arrays. Data is
'*                        added to the end of the buffer.
'****************************************************************
Public Function AddToBuffer(ByRef Buffer() As Byte, ByRef Additional() As Byte, Optional ByVal ByteLen As Long) As Byte()
  Dim tBytes() As Byte
  Dim tLen As Long
  
  If ByteLen = 0 Then
    ByteLen = aLen(Additional)
  End If
  tLen = aLen(Buffer) + ByteLen
  
  ReDim tBytes(tLen - 1)
  
  If Not UBound(Buffer) < 0 Then
    Call CopyMemory(tBytes(0), Buffer(0), aLen(Buffer))
  End If
  
  Call CopyMemory(tBytes(aLen(Buffer)), Additional(0), aLen(Additional))
  
  AddToBuffer = tBytes
End Function
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 09/13/2005  Verrigan   AddByteToBuffer() function created to
'*                        add a byte to the end of a buffer.
'****************************************************************
Public Function AddByteToBuffer(ByRef Buffer() As Byte, ByVal vData As Byte) As Byte()
  Dim tBytes() As Byte

  ReDim tBytes(UBound(Buffer) + 1)
  
  If aLen(Buffer) > 0 Then
    Call CopyMemory(tBytes(0), Buffer(0), aLen(Buffer))
  End If
  
  Call CopyMemory(tBytes(aLen(Buffer)), vData, 1)
  
  AddByteToBuffer = tBytes
End Function
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 09/13/2005  Verrigan   AddIntegerToBuffer() function created
'*                        to add an integer to the end of a
'*                        buffer.
'****************************************************************
Public Function AddIntegerToBuffer(ByRef Buffer() As Byte, vData As Integer) As Byte()
  Dim tBytes() As Byte

  ReDim tBytes(UBound(Buffer) + 2)
  
  If aLen(Buffer) > 0 Then
    Call CopyMemory(tBytes(0), Buffer(0), aLen(Buffer))
  End If
  
  Call CopyMemory(tBytes(aLen(Buffer)), vData, 2)
  
  AddIntegerToBuffer = tBytes
End Function
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 09/13/2005  Verrigan   AddLongToBuffer() function created to
'*                        add a long to the end of a buffer.
'****************************************************************
Public Function AddLongToBuffer(ByRef Buffer() As Byte, vData As Long) As Byte()
  Dim tBytes() As Byte

  ReDim tBytes(UBound(Buffer) + 2)
  
  If aLen(Buffer) > 0 Then
    Call CopyMemory(tBytes(0), Buffer(0), aLen(Buffer))
  End If
  
  Call CopyMemory(tBytes(aLen(Buffer)), vData, 4)
  
  AddLongToBuffer = tBytes
End Function
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 09/13/2005  Verrigan   AddStringToBuffer() function created to
'*                        add a string to the end of a buffer.
'*                        It adds the string length (as an
'*                        integer) and then the string.
'****************************************************************
Public Function AddStringToBuffer(ByRef Buffer() As Byte, vData As String) As Byte()
  Dim sLen As Integer
  Dim sBytes() As Byte
  Dim tBytes() As Byte
  
  ReDim tBytes(UBound(Buffer) + 2 + Len(vData))
  
  sLen = CInt(Len(vData))
  sBytes = StrConv(vData, vbFromUnicode)
  If aLen(Buffer) > 0 Then
    Call CopyMemory(tBytes(0), Buffer(0), aLen(Buffer))
  End If
  
  Call CopyMemory(tBytes(aLen(Buffer)), sLen, 2)
  Call CopyMemory(tBytes(aLen(Buffer) + 2), sBytes(0), aLen(sBytes))
  
  AddStringToBuffer = tBytes
End Function
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 09/13/2005  Verrigan   GetFromBuffer() function created to
'*                        grab a specified number of bytes from
'*                        the beginning of the buffer. All of the
'*                        Get*FromBuffer() functions have the
'*                        option to remove those bytes before it
'*                        returns the value to the caller.
'****************************************************************
Public Function GetFromBuffer(ByRef Buffer() As Byte, ByVal ByteLen As Long, Optional ClearBytes As Boolean) As Byte()
  Dim tBytes() As Byte

  ReDim tBytes(ByteLen - 1)
  
  Call CopyMemory(tBytes(0), Buffer(0), ByteLen)
  
  GetFromBuffer = tBytes
  
  If ClearBytes Then
    Buffer = RemoveFromBuffer(Buffer, ByteLen)
  End If
End Function
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 09/13/2005  Verrigan   GetByteFromBuffer() function created to
'*                        grab 1 byte from the beginning of the
'*                        buffer.
'****************************************************************
Public Function GetByteFromBuffer(ByRef Buffer() As Byte, Optional ClearBytes As Boolean) As Byte
  Dim RetVal As Byte
  
  Call CopyMemory(RetVal, Buffer(0), 1)
  
  GetByteFromBuffer = RetVal
  
  If ClearBytes Then
    Buffer = RemoveFromBuffer(Buffer, 1)
  End If
End Function
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 09/13/2005  Verrigan   GetIntegerFromBuffer() function created
'*                        to grab 1 integer from the beginning of
'*                        the buffer.
'****************************************************************
Public Function GetIntegerFromBuffer(ByRef Buffer() As Byte, Optional ClearBytes As Boolean) As Integer
  Dim RetVal As Integer
  
  Call CopyMemory(RetVal, Buffer(0), 2)
  
  GetIntegerFromBuffer = RetVal
  
  If ClearBytes Then
    Buffer = RemoveFromBuffer(Buffer, 2)
  End If
End Function
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 09/13/2005  Verrigan   GetLongFromBuffer() function created to
'*                        grab 1 long from the beginning of the
'*                        buffer.
'****************************************************************
Public Function GetLongFromBuffer(ByRef Buffer() As Byte, Optional ClearBytes As Boolean) As Long
  Dim RetVal As Long
  
  Call CopyMemory(RetVal, Buffer(0), 4)
  
  GetLongFromBuffer = RetVal
  
  If ClearBytes Then
    Buffer = RemoveFromBuffer(Buffer, 4)
  End If
End Function
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 09/13/2005  Verrigan   GetStringFromBuffer() function created
'*                        to grab a string from the beginning of
'*                        the buffer. It grabs the first 2 bytes
'*                        to determine the size of the string,
'*                        and then the string.
'****************************************************************
Public Function GetStringFromBuffer(ByRef Buffer() As Byte, Optional ClearBytes As Boolean) As String
  Dim sLen As Integer
  Dim tBytes() As Byte
  
  sLen = GetIntegerFromBuffer(Buffer)
  
  ReDim tBytes(sLen - 1)
  
  Call CopyMemory(tBytes(0), Buffer(2), sLen)
  
  GetStringFromBuffer = StrConv(tBytes, vbUnicode)
  
  If ClearBytes Then
    Buffer = RemoveFromBuffer(Buffer, sLen + 2)
  End If
End Function
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 09/13/2005  Verrigan   RemoveFromBuffer() function created to
'*                        delete a specified number of bytes from
'*                        the beginning of the buffer.
'****************************************************************
Public Function RemoveFromBuffer(ByRef Buffer() As Byte, ByVal ByteLen As Long) As Byte()
  Dim tBytes() As Byte
  
  If ByteLen > UBound(Buffer) Then
    RemoveFromBuffer = ""
    Exit Function
  End If
  
  ReDim tBytes(UBound(Buffer) - ByteLen)
  
  Call CopyMemory(tBytes(0), Buffer(LBound(Buffer) + ByteLen), aLen(Buffer) - ByteLen)
  
  RemoveFromBuffer = tBytes
End Function
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 09/13/2005  Verrigan   FillBuffer() function created to copy
'*                        a specified number of bytes from a
'*                        specified memory location to a buffer.
'****************************************************************
Public Function FillBuffer(ByVal StartAddr As Long, ByVal ByteLen As Long) As Byte()
  Dim tBytes() As Byte
  
  ReDim tBytes(ByteLen - 1)
  
  Call CopyMemory(tBytes(0), ByVal StartAddr, ByteLen)
  
  FillBuffer = tBytes
End Function

