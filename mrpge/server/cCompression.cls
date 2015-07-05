VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCompression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'//*****************************************************************************
'// Project: Compress Example
'//
'// Object Name: cCompress.cls
'// Description: ZLIB Library Compression Wrapper
'// Programmer:  Gary Noble
'// Date:        October 05, 2002
'// References:
'// Comments:    Uses ZLIB.DLL
'//*****************************************************************************
'// Modifications:
'//   GN - Created at 10:0 PM
'//
'//*****************************************************************************

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function compress Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Function compress2 Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long, ByVal level As Long) As Long
Private Declare Function uncompress Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long


Const m_def_CompressedSize = 0
Const m_def_OriginalSize = 0

Dim m_CompressedSize As Long
Dim m_OriginalSize As Long

Enum enumCompressErrors
    COMPRESS_OK = 0
    COMPRESS_STREAM_END = 1
    COMPRESS_NEED_DICT = 2
    COMPRESS_ERRNO = -1
    COMPRESS_STREAM_ERROR = -2
    COMPRESS_DATA_ERROR = -3
    COMPRESS_MEM_ERROR = -4
    COMPRESS_BUF_ERROR = -5
    COMPRESS_VERSION_ERROR = -6
End Enum

Enum CompressionLevels
    COMPRESS_NONE = 0
    COMPRESS_BEST_SPEED = 1
    '//-- note that levels 2-8 exist, too
    COMPRESS_BEST = 9
    COMPRESS_DEFAULT = -1
End Enum


'//*****************************************************************************
'// Procedure:   CompressByteArray
'// Description: Compresses A Byte Array
'// Parameters:  TheData - IO:  (Variant)
'//
'// Returns:     Byte, CompressionLevel As CompressionLevels) As Long
'// Comments:    None.
'// Programmer:  Gary Noble
'// Date:        October 05, 2002
'// Comments:    None.
'//*****************************************************************************
'// Modifications:
'//   GN - Created On October 05, 2002 at 10:2 PM
'//
'//*****************************************************************************
Public Function CompressByteArray(TheData() As Byte, CompressionLevel As CompressionLevels) As Long
Dim result          As Long
Dim BufferSize      As Long
Dim TempBuffer()    As Byte

OriginalSize = UBound(TheData) + 1

'//-- Allocate memory for byte array
BufferSize = UBound(TheData) + 1
BufferSize = BufferSize + (BufferSize * 0.01) + 12
ReDim TempBuffer(BufferSize)

'//-- Compress byte array (data)
result = compress2(TempBuffer(0), BufferSize, TheData(0), UBound(TheData) + 1, CompressionLevel)

'//-- Truncate to compressed size
ReDim Preserve TheData(BufferSize - 1)
CopyMemory TheData(0), TempBuffer(0), BufferSize

'//-- Cleanup
Erase TempBuffer

'//-- Set property
CompressedSize = UBound(TheData) + 1

'//-- Return Error/Success
CompressByteArray = result

End Function


'//*****************************************************************************
'// Procedure:   CompressString
'// Description: Compresses Any Passed String
'// Parameters:  TheString        - O:   (String)
'//              CompressionLevel - O:   (Integer)
'//
'// Returns:     Long
'// Comments:    None.
'// Programmer:  Gary Noble
'// Date:        October 05, 2002
'// Comments:    None.
'//*****************************************************************************
'// Modifications:
'//   GN - Created On October 05, 2002 at 10:3 PM
'//
'//*****************************************************************************
Public Function CompressString(TheString As String, CompressionLevel As CompressionLevels) As Long
Dim orgSize     As Long
Dim ret         As Long

OriginalSize = Len(TheString)

'//-- Allocate string space for the buffers
Dim CmpSize As Long
Dim TBuff As String

orgSize = Len(TheString)
TBuff = String(orgSize + (orgSize * 0.01) + 12, 0)
CmpSize = Len(TBuff)

'//-- Compress string (temporary string buffer) data
ret = compress2(ByVal TBuff, CmpSize, ByVal TheString, Len(TheString), CompressionLevel)

'//-- Crop the string and set it to the actual string.
TheString = Left$(TBuff, CmpSize)

'//-- Set compressed size of string.
CompressedSize = CmpSize

'//-- Cleanup
TBuff = ""

'//-- Return Error/Success
CompressString = ret

End Function



'//*****************************************************************************
'// Procedure:   DecompressByteArray
'// Description: Decompress A Byte Array
'// Parameters:  TheData - IO:  (Variant)
'//
'// Returns:     Byte, OrigSize As Long) As Long
'// Comments:    None.
'// Programmer:  Gary Noble
'// Date:        October 05, 2002
'// Comments:    None.
'//*****************************************************************************
'// Modifications:
'//   GN - Created On October 05, 2002 at 10:5 PM
'//
'//*****************************************************************************
Public Function DecompressByteArray(TheData() As Byte, OrigSize As Long) As Long
Dim result As Long

'//-- Allocate memory for buffers
Dim BufferSize As Long
Dim TempBuffer() As Byte

OriginalSize = OrigSize
CompressedSize = UBound(TheData) + 1

BufferSize = OrigSize
BufferSize = BufferSize + (BufferSize * 0.01) + 12
ReDim TempBuffer(BufferSize)

'//-- Decompress Data
result = uncompress(TempBuffer(0), BufferSize, TheData(0), UBound(TheData) + 1)

'//-- Truncate buffer to compressed size
ReDim Preserve TheData(BufferSize - 1)
CopyMemory TheData(0), TempBuffer(0), BufferSize

'//-- Return Error/Success
DecompressByteArray = result

End Function


'//*****************************************************************************
'// Procedure:   DecompressString
'// Description: Descompresses A Given String
'// Parameters:  TheString - O:   (String)
'//              OrigSize  - O:   (Long)
'//
'// Returns:     Long
'// Comments:    None.
'// Programmer:  Gary Noble
'// Date:        October 05, 2002
'// Comments:    None.
'//*****************************************************************************
'// Modifications:
'//   GN - Created On October 05, 2002 at 10:6 PM
'//
'//   MH - Moved 'OriginalSize = OrigSize' line to above the TBuff = String...
'//        line, so the correct buffer size supplied in OrigSize is always used.
'//        Modification on April 29, 2003 by Matt Hall 2003,
'//        http://darksideispn.no-ip.com
'//
'//*****************************************************************************
Public Function DecompressString(TheString As String, OrigSize As Long) As Long
Dim result As Long

'//-- Allocate string space
Dim CmpSize As Long
Dim TBuff As String

'//-- Set OriginalSize property
OriginalSize = OrigSize

TBuff = String(OriginalSize + (OriginalSize * 0.01) + 12, 0)
CmpSize = Len(TBuff)

'//-- Decompress
result = uncompress(ByVal TBuff, CmpSize, ByVal TheString, Len(TheString))

'//-- Make string the size of the uncompressed string
TheString = Left$(TBuff, CmpSize)

CompressedSize = CmpSize

'//-- Return Error/Success
DecompressString = result

End Function


'//*****************************************************************************
'// Procedure:   CompressFile
'// Description: Compress A File
'// Parameters:  FilePathIn       - O:   (String)
'//              FilePathOut      - O:   (String)
'//              CompressionLevel - O:   (Integer)
'//
'// Returns:     Long
'// Comments:    None.
'// Programmer:  Gary Noble
'// Date:        October 05, 2002
'// Comments:    None.
'//*****************************************************************************
'// Modifications:
'//   GN - Created On October 05, 2002 at 10:7 PM
'//
'//*****************************************************************************
Public Function CompressFile(FilePathIn As String, FilePathOut As String, CompressionLevel As CompressionLevels) As Long
Dim TheBytes()  As Byte
Dim result      As Long
Dim lngFileLen  As Long

lngFileLen = FileLen(FilePathIn)

'//-- allocate byte array
ReDim TheBytes(lngFileLen - 1)

'//-- read byte array from file
Open FilePathIn For Binary Access Read As #1
    Get #1, , TheBytes()
Close #1

'//-- compress byte array
result = CompressByteArray(TheBytes(), CompressionLevel)

'//-- kill any file in place
On Error Resume Next

Kill FilePathOut
On Error GoTo 0

'//-- Write it out
Open FilePathOut For Binary Access Write As #1
    Put #1, , lngFileLen 'must store the length of the original file
    Put #1, , TheBytes()
Close #1

'//-- Clean up
Erase TheBytes
CompressFile = result

End Function


'//*****************************************************************************
'// Procedure:   DecompressFile
'// Description: Decompress A File
'// Parameters:  FilePathIn  - O:   (String)
'//              FilePathOut - O:   (String)
'//
'// Returns:     Long
'// Comments:    None.
'// Programmer:  Gary Noble
'// Date:        October 05, 2002
'// Comments:    None.
'//*****************************************************************************
'// Modifications:
'//   GN - Created On October 05, 2002 at 10:8 PM
'//
'//*****************************************************************************
Public Function DecompressFile(FilePathIn As String, FilePathOut As String) As Long
Dim TheBytes()  As Byte
Dim result      As Long
Dim lngFileLen  As Long

'//-- Allocate byte array
ReDim TheBytes(FileLen(FilePathIn) - 1)

'//-- Read byte array from file
Open FilePathIn For Binary Access Read As #1
    Get #1, , lngFileLen    '//-- the original (uncompressed) file's length
    Get #1, , TheBytes()
Close #1

'//-- Decompress
result = DecompressByteArray(TheBytes(), lngFileLen)

'//-- kill any file already there
On Error Resume Next
Kill FilePathOut
On Error GoTo 0

'//-- Write it out
Open FilePathOut For Binary Access Write As #1
    Put #1, , TheBytes()
Close #1

'//-- Clean Up
Erase TheBytes
DecompressFile = result

End Function

Public Property Get CompressedSize() As Long
    CompressedSize = m_CompressedSize
End Property

Private Property Let CompressedSize(ByVal New_CompressedSize As Long)
    m_CompressedSize = New_CompressedSize
    
End Property

Public Property Get OriginalSize() As Long
    OriginalSize = m_OriginalSize
    
End Property

Private Property Let OriginalSize(ByVal New_OriginalSize As Long)
    m_OriginalSize = New_OriginalSize
    
End Property

Private Sub Class_Initialize()
    m_CompressedSize = m_def_CompressedSize
    m_OriginalSize = m_def_OriginalSize
    
End Sub

