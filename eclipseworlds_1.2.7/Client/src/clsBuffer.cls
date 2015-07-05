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

Private buffer() As Byte
Private BufferSize As Long
Private WriteHead As Long
Private ReadHead As Long

Private Sub Class_Initialize()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Flush
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Class_Initialize", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub PreAllocate(ByVal nLength As Long)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    WriteHead = 0
    ReadHead = 0
    BufferSize = nLength - 1
    ReDim buffer(0 To BufferSize)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "PreAllocate", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub Allocate(ByVal nLength As Long)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If BufferSize = 0 And nLength > 1 Then nLength = nLength - 1
    BufferSize = BufferSize + nLength
    ReDim Preserve buffer(0 To BufferSize)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Allocate", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub Flush()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    WriteHead = 0
    ReadHead = 0
    BufferSize = 0
    ReDim buffer(0)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Flush", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub Trim()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' If the readhead is past the buffersize, this means everything has been read in the array, flush it
    If ReadHead >= count Then
        Flush
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Trim", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub WriteByte(ByVal nByte As Byte)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If WriteHead > BufferSize Then Allocate 1
 
    buffer(WriteHead) = nByte
    WriteHead = WriteHead + 1
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "WriteByte", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub WriteBytes(ByRef nByte() As Byte)
    Dim nLength As Long
        
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    nLength = (UBound(nByte) - LBound(nByte)) + 1
    
    If WriteHead + nLength - 1 > BufferSize Then Allocate nLength
    
    CopyMemory buffer(WriteHead), nByte(0), nLength
    WriteHead = WriteHead + nLength
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "WriteBytes", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub WriteInteger(ByVal nInteger As Integer)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If WriteHead + 1 > BufferSize Then Allocate 2
        
    CopyMemory buffer(WriteHead), nInteger, 2
    WriteHead = WriteHead + 2
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "WriteInteger", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub WriteLong(ByVal nLong As Long)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If WriteHead + 3 > BufferSize Then Allocate 4
    
    CopyMemory buffer(WriteHead), nLong, 4
    WriteHead = WriteHead + 4
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "WriteLong", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub WriteString(ByRef nString As String)
    Dim sBytes() As Byte
        Dim sLength As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    sLength = Len(nString)
    sBytes = StrConv(nString, vbFromUnicode)
    
    WriteLong sLength
    
    If sLength <= 0 Then Exit Sub
    
    If WriteHead + sLength - 1 > BufferSize Then Allocate sLength
      
    CopyMemory buffer(WriteHead), sBytes(0), sLength
    WriteHead = WriteHead + sLength
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "WriteString", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Function ReadByte(Optional MoveReadHead As Boolean = True) As Byte
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If ReadHead > BufferSize Then Exit Function
    
    ReadByte = buffer(ReadHead)
    If MoveReadHead Then ReadHead = ReadHead + 1
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "ReadByte", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Public Function ReadBytes(ByVal nLength As Long, Optional MoveReadHead As Boolean = True) As Byte()
    Dim data() As Byte
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If nLength = 0 Then Exit Function
    If ReadHead + nLength - 1 > BufferSize Then Exit Function
        
    ReDim data(nLength - 1)
        
    CopyMemory data(0), buffer(ReadHead), nLength
    If MoveReadHead Then ReadHead = ReadHead + nLength
    
    ReadBytes = data
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "ReadBytes", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Public Function ReadInteger(Optional MoveReadHead As Boolean = True) As Integer
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If ReadHead + 1 > BufferSize Then Exit Function
    
    CopyMemory ReadInteger, buffer(ReadHead), 2
    If MoveReadHead Then ReadHead = ReadHead + 2
    Exit Function

' Error handler
ErrorHandler:
    HandleError "ReadInteger", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Public Function ReadLong(Optional MoveReadHead As Boolean = True) As Long
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If ReadHead + 3 > BufferSize Then Exit Function
    
    CopyMemory ReadLong, buffer(ReadHead), 4
    If MoveReadHead Then ReadHead = ReadHead + 4
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "ReadLong", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Public Function ReadString(Optional MoveReadHead As Boolean = True) As String
    Dim sLength As Long
    Dim sBytes() As Byte
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    sLength = ReadLong(False)

    If sLength <= 0 Then
        If MoveReadHead Then ReadHead = ReadHead + 4
        Exit Function
    End If
    
    ReDim sBytes(sLength - 1)
    
    CopyMemory sBytes(0), buffer(ReadHead + 4), sLength
    
    ReadString = StrConv(sBytes, vbUnicode)
    If MoveReadHead Then ReadHead = ReadHead + sLength + 4
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "ReadString", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Public Function count() As Long
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    count = (UBound(buffer) - LBound(buffer)) + 1
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "Count", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Public Function Length() As Long
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Length = count - ReadHead
    Exit Function
        
' Error handler
ErrorHandler:
    HandleError "Length", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Public Function ToArray() As Byte()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ToArray = buffer()
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "ToArray", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Public Function ToString() As String
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ToString = StrConv(buffer, vbUnicode)
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "ToString", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function
