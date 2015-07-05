VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This class will be used for building packets
Private Buffer() As Byte        ' Our byte array to hold all data

Private BufferSize As Long      ' Used to keep track of how big our buffer is
Private WriteHead As Long       ' This is used for holding our position in the byte array when writing variables
Private ReadHead As Long        ' This is used for holding our position in the byte array when reading variables

Private Sub Class_Initialize()
    Flush
End Sub

Public Sub PreAllocate(ByVal nLength As Long)
    WriteHead = 0
    ReadHead = 0
    BufferSize = nLength - 1
    ReDim Buffer(0 To BufferSize)
End Sub

Public Sub Allocate(ByVal nLength As Long)
    BufferSize = BufferSize + nLength
    ReDim Preserve Buffer(0 To BufferSize)
End Sub

Public Sub Flush()
    WriteHead = 0
    ReadHead = 0
    BufferSize = 0
    ReDim Buffer(0)
End Sub

Public Sub WriteByte(ByVal nByte As Byte)

    If WriteHead > BufferSize Then Allocate 1
 
    Buffer(WriteHead) = nByte
    WriteHead = WriteHead + 1
End Sub

Public Sub WriteBytes(ByRef nByte() As Byte)
Dim nLength As Long
    
    nLength = (UBound(nByte) - LBound(nByte)) + 1
    
    If WriteHead + nLength - 1 > BufferSize Then Allocate nLength
    
    CopyMemory Buffer(WriteHead), nByte(0), nLength
    WriteHead = WriteHead + nLength
End Sub

Public Sub WriteInteger(ByVal nInteger As Integer)

    If WriteHead + 1 > BufferSize Then Allocate 2
        
    CopyMemory Buffer(WriteHead), nInteger, 2
    WriteHead = WriteHead + 2
End Sub

Public Sub WriteLong(ByVal nLong As Long)

    If WriteHead + 3 > BufferSize Then Allocate 4
    
    CopyMemory Buffer(WriteHead), nLong, 4
    WriteHead = WriteHead + 4
End Sub

Public Sub WriteString(ByRef nString As String)
Dim sBytes() As Byte
Dim sLength As Long
    
    sLength = Len(nString)
    sBytes = StrConv(nString, vbFromUnicode)
    
    WriteLong sLength
    
    If sLength <= 0 Then Exit Sub
    
    If WriteHead + sLength - 1 > BufferSize Then Allocate sLength
      
    CopyMemory Buffer(WriteHead), sBytes(0), sLength
    WriteHead = WriteHead + sLength
End Sub

Public Function ReadByte(Optional MoveReadHead As Boolean = True) As Byte

    If ReadHead > BufferSize Then Exit Function
    
    ReadByte = Buffer(ReadHead)
    If MoveReadHead Then ReadHead = ReadHead + 1
End Function

Public Function ReadBytes(ByVal nLength As Long, Optional MoveReadHead As Boolean = True) As Byte()
Dim Data() As Byte

    If ReadHead + nLength - 1 > BufferSize Then Exit Function
    
    ReDim Data(nLength)
        
    CopyMemory Data(0), Buffer(ReadHead), nLength
    If MoveReadHead Then ReadHead = ReadHead + nLength
    
    ReadBytes = Data
End Function

Public Function ReadInteger(Optional MoveReadHead As Boolean = True) As Integer

    If ReadHead + 1 > BufferSize Then Exit Function
    
    CopyMemory ReadInteger, Buffer(ReadHead), 2
    If MoveReadHead Then ReadHead = ReadHead + 2
End Function

Public Function ReadLong(Optional MoveReadHead As Boolean = True) As Long

    If ReadHead + 3 > BufferSize Then Exit Function
    
    CopyMemory ReadLong, Buffer(ReadHead), 4
    If MoveReadHead Then ReadHead = ReadHead + 4
End Function

Public Function ReadString(Optional MoveReadHead As Boolean = True) As String
Dim sLength As Long
Dim sBytes() As Byte
  
    sLength = ReadLong
    If sLength <= 0 Then Exit Function
    
    ReDim sBytes(sLength - 1)
    
    CopyMemory sBytes(0), Buffer(ReadHead), sLength
    
    ReadString = StrConv(sBytes, vbUnicode)
    If MoveReadHead Then ReadHead = ReadHead + sLength
End Function

Public Function Count() As Long
    Count = (UBound(Buffer) - LBound(Buffer)) + 1
End Function

Public Function Length() As Long
    Length = Count - ReadHead
End Function

Public Function ToArray() As Byte()
    ToArray = Buffer()
End Function

Public Function ToString() As String
    ToString = StrConv(Buffer, vbUnicode)
End Function

Public Function CompressBuffer() As Long
Dim Result As Long
Dim Size As Long
Dim tempSize As Long
Dim tempBuffer() As Byte
Dim compressed() As Byte

    CompressBuffer = -1
    
    tempSize = Count
    If tempSize > 0 Then
        Size = tempSize
        Size = Size + (Size * 0.01) + 12
        ReDim tempBuffer(0 To Size - 1) As Byte
        
        Result = Compress(tempBuffer(0), Size, Buffer(0), tempSize)
        
        If Result = 0 Then
            ReDim compressed(0 To Size - 1)
            CopyMemory compressed(0), tempBuffer(0), Size
        
            Flush
            WriteLong tempSize      ' original size
            WriteLong Size          ' compressed size
            WriteBytes compressed() ' compressed data
            CompressBuffer = 0
        End If
    End If
End Function

Public Function DecompressBuffer() As Long
Dim Result As Long
Dim OrigSize As Long
Dim tempBuffer() As Byte
Dim compressed() As Byte

    DecompressBuffer = -1
    
    If Count > 0 Then
        OrigSize = ReadLong + 1
        ReDim tempBuffer(0 To OrigSize - 1) As Byte

        compressed = ReadBytes(ReadLong)
        
        Result = uncompress(tempBuffer(0), OrigSize, compressed(0), UBound(compressed) + 1)
        
        If Result = 0 Then
            ReDim compressed(0 To OrigSize - 1)
            CopyMemory compressed(0), tempBuffer(0), OrigSize
        
            Flush
            WriteBytes compressed()
            DecompressBuffer = 0
        End If
    End If
End Function


