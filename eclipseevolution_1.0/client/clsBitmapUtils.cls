VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BitmapUtils"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 1001
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'clsBitmapUtils.cls, by Ron Zuber
'
'This entire module is Copyright © 2005-2006 by Ron Zuber, except for certain
'parts described below. All Rights Reserved.
'
'TERMS OF USE:
'The source for this module was created by Ron Zuber (Verrigan) and was
'released for use in all visual basic applications. You are granted the
'right to copy functions from this source into your own modules, provided
'that this header is placed in any module that contains any of this source.
'
'DESCRIPTION:
'This class module was created to load a bitmap into memory, and blt it
'directly to any object that has a device context (DC). (DirectX Surface,
'Picture Box, Form, etc.)
'
'DISCLAIMER:
'You agree to use this class module at your "own-risk".
'Ron Zuber shall not be held responsible if your computer freezes, blows
'up, or shuts down unexpectedly due to the use of this source.
'Ron Zuber shall not be held responsible if you fail to save your project
'prior to using this source, and you lose all changes to your source
'because of the use of this source. (i.e. VB Crashes)
'
'NOT COVERED UNDER TERMS:
'To demonstrate the use of this code, en(de)cryption and (de)compression
'methods have been added. Base code for the en(de)cryption and
'(de)compression methods was found at http://www.planetsourcecode.com.
'
'The original author of the (de)cryption code is unknown, as the reference
'I used to create this class module used code from someone else, and did
'not reference the original author. If the original author is found, this
'class module will be updated accordingly.
'
'The original author of the (de)compression code was Alex Ionescu, and is
'Copyright © 2003 Alex Ionescu. All Rights Reserved.
'
'END HEADER
Option Explicit

'Crypto API Declarations
Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, ByVal pbData As String, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptDeriveKey Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hBaseData As Long, ByVal dwFlags As Long, ByRef phKey As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Private Declare Function CryptEncrypt Lib "advapi32.dll" (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, ByRef pdwDataLen As Long, ByVal dwBufLen As Long) As Long
Private Declare Function CryptDestroyKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptDecrypt Lib "advapi32.dll" (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, ByRef pdwDataLen As Long) As Long

'Compression API Declarations
Private Declare Function BufferCompress Lib "ntdll.dll" Alias "RtlCompressBuffer" (ByVal CompressionFormatAndEngine As Integer, ByVal UnCompressedBuffer As Long, ByVal UnCompressedBufferSize As Long, ByVal CompressedBuffer As Long, ByVal CompressedBufferSize As Long, ByVal UncompressedChunkSize As Long, FinalCompressedSize As Long, ByVal WorkSpace As Long) As Long
Private Declare Function BufferDecompress Lib "ntdll.dll" Alias "RtlDecompressBuffer" (ByVal CompressionFormat As Integer, ByVal UnCompressedBufferPtr As Long, ByVal UnCompressedBufferSize As Long, ByVal CompressedBuffer As Long, ByVal CompressedBufferSize As Long, CompressedSize As Long) As Long
Private Declare Function GetCompressionWorkSpaceSize Lib "ntdll.dll" Alias "RtlGetCompressionWorkSpaceSize" (ByVal CompressionFormatAndEngine As Integer, CompressBufferWorkSpaceSize As Long, CompressFragmentWorkSpaceSize As Long) As Long
Private Declare Function VMemAlloc Lib "ntdll.dll" Alias "NtAllocateVirtualMemory" (ByVal ProcessHandle As Long, BaseAddress As Long, ByVal ZeroBits As Long, RegionSize As Long, ByVal AllocationType As Long, ByVal Protect As Long) As Long
Private Declare Function VMemFree Lib "ntdll.dll" Alias "NtFreeVirtualMemory" (ByVal ProcessHandle As Long, BaseAddress As Long, RegionSize As Long, ByVal FreeType As Long) As Long

'ZLib Compression API Declarations
Private Declare Function ZCompress Lib "zlib.dll" Alias "compress" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Function ZUncompress Lib "zlib.dll" Alias "uncompress" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long

'Kernal API Declarations
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal ByteLen As Long)

'GDI API Declarations
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hDC As Long, ByVal pBitmapInfo As Long, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function SetDIBits Lib "gdi32.dll" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, ByVal lpBI As Long, ByVal wUsage As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'Module Constants
Private Const MODULE_NAME             As String = "Bitmap Utility Class"
Private Const ERR_NUMBER_SUCCESS      As Long = 0&

'Error numbers specific to the module.
Private Enum MODULE_ERRORS
  ERR_BASE = vbObjectError + 1001 'The starting error.
  ERR_FILE_EXISTS
  ERR_INVALID_LENGTH
  ERR_CREATE_DC
  ERR_CREATE_BITMAP
  ERR_CRYPT_ACQUIRECONTEXT
  ERR_CRYPT_CREATEHASH
  ERR_CRYPT_HASHDATA
  ERR_CRYPT_DERIVEKEY
  ERR_CRYPT_ENCRYPT
  ERR_CRYPT_DECRYPT
  ERR_CRYPT_EMPTYKEY
  ERR_ZLIB_MISSING
  ERR_UNKNOWN_ERR
End Enum

'Crypto Constants (Unknown Source)
Private Const ALG_CLASS_DATA_ENCRYPT  As Long = 24576&
Private Const ALG_CLASS_HASH          As Long = 32768
Private Const ALG_SID_MD5             As Long = 3&
Private Const ALG_SID_RC4             As Long = 1&
Private Const ALG_TYPE_ANY            As Long = 0&
Private Const ALG_TYPE_STREAM         As Long = 2048&
Private Const CALG_MD5                As Long = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_MD5)
Private Const CALG_RC4                As Long = ((ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_STREAM) Or ALG_SID_RC4)
Private Const CRYPT_NEWKEYSET         As Long = 8&
Private Const ENCRYPT_ALGORITHM       As Long = CALG_RC4
Private Const PROV_RSA_FULL           As Long = 1&
Private Const KEY_CONTAINER           As String = "TestString"
Private Const SERVICE_PROVIDER        As String = "Microsoft Base Cryptographic Provider v1.0"

'Compression Buffer Constants (Alex Ionescu)
Private Const MEM_COMMIT              As Long = &H1000
Private Const MEM_DECOMMIT            As Long = &H4000
Private Const PAGE_EXECUTE_READWRITE  As Long = &H40
Private Const cFormat = &H2 Or &H100 'NT only supports LZNT1 &H100 = max compression.

'Image Constants
Private Const TYPE_BITMAP As Integer = 19778
Private Const DIB_RGB_COLORS = 0 'No color table defined.
Private Const CBM_INIT = &H4
Private Const MIN_HEADER_LEN = 54

'Image Offsets
Private Const BM_OS_OFFSET = 10
Private Const BM_INFO_OFFSET = 14
Private Const BM_WIDTH_OFFSET = 18
Private Const BM_HEIGHT_OFFSET = 22
Private Const BM_BITCOUNT_OFFSET = 28

'Local properties
Private mvarByteData() As Byte
'Sub LoadBitmap() - Left in for compatibility. Calls LoadByteData
Public Sub LoadBitmap(ByRef FileName As String)
  Call LoadByteData(FileName)
End Sub
'Sub SaveBitmap() - Left in for compatibility. Calls SaveByteData
Public Sub SaveBitmap(ByRef FileName As String)
  Call SaveByteData(FileName)
End Sub
'Sub CheckDataLen() - Checks the length of ByteDataLen against a position.
'If the length is less than the position, then an error is read.
Private Sub CheckDataLen(mPos As Long)
  If ByteDataLen < mPos Then
    Call Err.Raise(ERR_INVALID_LENGTH, MODULE_NAME, "Cannot read byte location.")
    Exit Sub
  End If
End Sub
'Property Get PixelOffset() - Returns the starting location of the pixels
'for the bitmap in the byte data.
Private Property Get PixelOffset()
On Error Resume Next
  Dim RetVal As Long
  
  'Check to make sure we have enough bytes stored in memory.
  Call CheckDataLen(MIN_HEADER_LEN)
  If Err.Number <> ERR_NUMBER_SUCCESS Then
    Call Err.Raise(Err.Number, Err.Source, Err.Description)
    Exit Property
  End If
  
  'Get the offset value of the pixels from the byte data.
  Call CopyMemory(RetVal, mvarByteData(LBound(mvarByteData) + BM_OS_OFFSET), 4)
  
  PixelOffset = RetVal
End Property
'Blt() - Creates a new Device Context (DC) and sets the bits on the new DC based
'        on the byte data. Then blt's to the destination DC.
Public Sub Blt(ByVal hDestDC As Long)
Attribute Blt.VB_Description = "Draw the byte data from memory to a device context."
Attribute Blt.VB_HelpID = 1012
On Error Resume Next
  Dim hBM As Long, hBitmap As Long
  Dim DC As Long
  
  'Check to make sure we have enough bytes stored in memory.
  Call CheckDataLen(MIN_HEADER_LEN)
  If Err.Number <> ERR_NUMBER_SUCCESS Then
    Call Err.Raise(Err.Number, Err.Source, Err.Description)
    Exit Sub
  End If
  
  'Create a compatible DC based on the passed destination.
  DC = CreateCompatibleDC(hDestDC)
  
  'Check to see if we were successful.
  If DC = 0 Then
    Call Err.Raise(ERR_CREATE_DC, MODULE_NAME, "Error creating device context.")
    Exit Sub
  End If
  
  'Create the device independant bitmap in the device context.
  hBM = CreateDIBSection(DC, VarPtr(mvarByteData(LBound(mvarByteData) + BM_INFO_OFFSET)), DIB_RGB_COLORS, ByVal 0&, ByVal 0&, ByVal 0&)
  
  'Check to see if we were successful.
  If hBM = 0 Then
    Call DeleteDC(DC)
    Call Err.Raise(ERR_CREATE_BITMAP, MODULE_NAME, "Unable to create bitmap.")
    Exit Sub
  End If
  
  'Select the newly created bitmap into the device context.
  Call SelectObject(DC, hBM)
  
  'Set the bits (pixels).
  Call SetDIBits(DC, hBM, 0, ImageHeight, mvarByteData(LBound(mvarByteData) + PixelOffset), VarPtr(mvarByteData(LBound(mvarByteData) + BM_INFO_OFFSET)), DIB_RGB_COLORS)
  
  'Blt the image to the destination device context.
  Call BitBlt(hDestDC, 0, 0, ImageWidth, ImageHeight, DC, 0, 0, vbSrcCopy)
  
  'The bitmap is copied.. Delete the bitmap handle and temporary DC.
  Call DeleteObject(hBM)
  Call DeleteDC(DC)
End Sub
'Compression code originally written by Alex Ionescu.
'Alex's code has been modified to work with this class module.
Public Sub CompressByteData()
Attribute CompressByteData.VB_Description = "Compress the byte data."
Attribute CompressByteData.VB_HelpID = 1008
On Error Resume Next
  Dim WorkSpace As Long, WorkSpaceSize As Long
  Dim LenUC As Long, LenC As Long
  Dim uBytes() As Byte, cBytes() As Byte
  
  LenUC = ByteDataLen
  
  LenC = LenUC
  
  If LenC < 32 Then LenC = 32
  
  Call GetCompressionWorkSpaceSize(cFormat, WorkSpaceSize, 0)
  Call VMemAlloc(-1, WorkSpace, 0, WorkSpaceSize, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
  
  ReDim uBytes(LenUC - 1)
  Call CopyMemory(uBytes(LBound(uBytes)), mvarByteData(LBound(mvarByteData)), ByteDataLen)
  
  'Empty our byte array to save memory.
  EraseByteData
  
  ReDim cBytes(LenC - 1)
  
  Call BufferCompress(cFormat, VarPtr(uBytes(LBound(uBytes))), LenUC, VarPtr(cBytes(LBound(cBytes))), LenC, &H1000, LenC, WorkSpace)
  
  ReDim Preserve cBytes(LenC - 1)
  
  Erase uBytes 'Erase the uncompressed bytes.
  
  ReDim mvarByteData(LenC + 4 - 1)
  Call CopyMemory(mvarByteData(LBound(mvarByteData)), LenUC, 4)
  Call CopyMemory(mvarByteData(LBound(mvarByteData) + 4), cBytes(LBound(cBytes)), LenC)

  Erase cBytes
  Call VMemFree(-1, WorkSpace, 0, MEM_DECOMMIT)
  WorkSpace = 0
End Sub
'CompressByteData_ZLib() - Compression code for using ZLib.
'Requires zlib.dll.
Public Sub CompressByteData_ZLib()
On Error Resume Next
  Dim LenUC As Long, LenC As Long
  Dim uBytes() As Byte, cBytes() As Byte
  
  LenUC = ByteDataLen
  LenC = LenUC
  
  If LenC < 32 Then LenC = 32
  
  ReDim uBytes(LenUC - 1)
  Call CopyMemory(uBytes(LBound(uBytes)), mvarByteData(LBound(mvarByteData)), ByteDataLen)
  
  'Empty our byte array to save memory.
  EraseByteData
  
  ReDim cBytes(LBound(uBytes) To UBound(uBytes))
  Call ZCompress(cBytes(LBound(cBytes)), LenC, uBytes(LBound(uBytes)), LenUC)
  If Err.Number <> ERR_NUMBER_SUCCESS Then
    Select Case Err.Number
    Case 53
      Call Err.Raise(ERR_ZLIB_MISSING, MODULE_NAME, "Missing file zlib.dll. Unable to compress data.")
      Err.Clear
    Case Else
      Call Err.Raise(ERR_UNKNOWN_ERR, MODULE_NAME, "An unknown error occured. (" & Err.Number & ") " & Err.Description)
      Err.Clear
    End Select
    Exit Sub
  End If
  
  ReDim Preserve cBytes(LenC - 1)
  
  Erase uBytes 'Erase the uncompressed bytes.
  
  ReDim mvarByteData(LenC + 4 - 1)
  Call CopyMemory(mvarByteData(LBound(mvarByteData)), LenUC, 4)
  Call CopyMemory(mvarByteData(LBound(mvarByteData) + 4), cBytes(LBound(cBytes)), LenC)
  
  Erase cBytes
End Sub
'Decompression code originally written by Alex Ionescu.
'Alex's code has been modified to work with this class module.
Public Sub DecompressByteData()
Attribute DecompressByteData.VB_Description = "Decompress the byte data."
Attribute DecompressByteData.VB_HelpID = 1009
On Error Resume Next
  Dim cBytes() As Byte, uBytes() As Byte
  Dim LenUC As Long, LenC As Long
  
  If ByteDataLen < 4 Then Exit Sub
  
  LenC = ByteDataLen - 4
  
  Call CopyMemory(LenUC, mvarByteData(LBound(mvarByteData)), 4)
  ReDim cBytes(LenC - 1)
  ReDim uBytes(LenUC - 1)
  
  Call CopyMemory(cBytes(LBound(cBytes)), mvarByteData(LBound(mvarByteData) + 4), LenC)
  
  'Empty our byte array to save memory.
  EraseByteData
  
  Call BufferDecompress(cFormat, VarPtr(uBytes(LBound(uBytes))), LenUC, VarPtr(cBytes(LBound(cBytes))), LenC, LenC)
  
  Erase cBytes
  
  ReDim mvarByteData(LenUC - 1)
  Call CopyMemory(mvarByteData(LBound(mvarByteData)), uBytes(LBound(uBytes)), LenUC)
  
  Erase uBytes
End Sub
'DecompressByteData_ZLib() - Decompression code for using ZLib.
'Requires zlib.dll.
Public Sub DecompressByteData_ZLib()
On Error Resume Next
  Dim cBytes() As Byte, uBytes() As Byte
  Dim LenUC As Long, LenC As Long
  
  If ByteDataLen < 4 Then Exit Sub
  
  LenC = ByteDataLen - 4
  
  Call CopyMemory(LenUC, mvarByteData(LBound(mvarByteData)), 4)
  ReDim cBytes(LenC - 1)
  ReDim uBytes(LenUC - 1)
  
  Call CopyMemory(cBytes(LBound(cBytes)), mvarByteData(LBound(mvarByteData) + 4), LenC)
  
  'Empty our byte array to save memory.
  EraseByteData
  
  Call ZUncompress(uBytes(LBound(uBytes)), LenUC, cBytes(LBound(cBytes)), LenC)
  If Err.Number <> ERR_NUMBER_SUCCESS Then
    Select Case Err.Number
    Case 53
      Call Err.Raise(ERR_ZLIB_MISSING, MODULE_NAME, "Missing file zlib.dll. Unable to decompress data.")
      Err.Clear
    Case Else
      Call Err.Raise(ERR_UNKNOWN_ERR, MODULE_NAME, "An unknown error occured. (" & Err.Number & ") " & Err.Description)
      Err.Clear
    End Select
    Exit Sub
  End If
  
  Erase cBytes
  
  ReDim mvarByteData(LenUC - 1)
  Call CopyMemory(mvarByteData(LBound(mvarByteData)), uBytes(LBound(uBytes)), LenUC)
  
  Erase uBytes
End Sub
'Encryption/Decryption code written by an unknown author, and modified
'to work with this class module.
Private Sub EncryptOrDecrypt(ByRef strBytes As String, ByRef Key As String, Encrypt As Boolean)
On Error Resume Next
  Dim CryptProv As Long
  Dim Hash      As Long
  Dim lngKey    As Long
  Dim Length    As Long
  
  Err.Clear
  
  'Acquire the crypt context
  If (CryptAcquireContext(CryptProv, KEY_CONTAINER, SERVICE_PROVIDER, PROV_RSA_FULL, CRYPT_NEWKEYSET) = 0) Then
    If (CryptAcquireContext(CryptProv, KEY_CONTAINER, SERVICE_PROVIDER, PROV_RSA_FULL, 0&) = 0) Then
      Call Err.Raise(ERR_CRYPT_ACQUIRECONTEXT, MODULE_NAME, "Error during CryptAcquireContext for a new key container." & vbCrLf & "A container with this name probably already exists.")
    End If
  End If
  'Create the hash
  If Err.Number = ERR_NUMBER_SUCCESS Then
    If (CryptCreateHash(CryptProv, CALG_MD5, 0&, 0&, Hash) = 0) Then
      Call Err.Raise(ERR_CRYPT_CREATEHASH, MODULE_NAME, "Could not create a Hash Object (CryptCreateHash API)")
    End If
  End If
  
  'Calculate the hash value
  If Err.Number = ERR_NUMBER_SUCCESS Then
    If (CryptHashData(Hash, Key, Len(Key), 0&) = 0) Then
      Call Err.Raise(ERR_CRYPT_HASHDATA, MODULE_NAME, "Could not calculate a Hash Value (CryptHashData API)")
    End If
  End If
  
  'Get the key
  If Err.Number = ERR_NUMBER_SUCCESS Then
    If (CryptDeriveKey(CryptProv, ENCRYPT_ALGORITHM, Hash, 0&, lngKey) = 0) Then
      Call Err.Raise(ERR_CRYPT_DERIVEKEY, MODULE_NAME, "Could not create a session key (CryptDeriveKey API)")
    End If
  End If
 
  If Err.Number = ERR_NUMBER_SUCCESS Then
    Length = Len(strBytes)

    If (Encrypt) Then 'Encrypt
      If (CryptEncrypt(lngKey, 0&, 1&, 0&, strBytes, Length, Length) = 0) Then
        Call Err.Raise(ERR_CRYPT_ENCRYPT, MODULE_NAME, "Error during CryptEncrypt.")
      End If
    Else ' Decrypt
      If (CryptDecrypt(lngKey, 0&, 1&, 0&, strBytes, Length) = 0) Then
        Call Err.Raise(ERR_CRYPT_DECRYPT, MODULE_NAME, "Error during CryptDecrypt.")
      End If
    End If
  End If
  
  If Err.Number <> ERR_NUMBER_SUCCESS Then
    If Err.Number = 53 Then
      GoTo Exit_EncryptOrDecrypt
RaiseError:
      Call Err.Raise(Err.Number, MODULE_NAME, Err.Description)
    End If
  End If
 
Exit_EncryptOrDecrypt:
On Error Resume Next
  If lngKey <> 0& Then Call CryptDestroyKey(lngKey)
  If Hash <> 0& Then Call CryptDestroyHash(Hash)
  If CryptProv <> 0& Then Call CryptReleaseContext(CryptProv, 0)
  If Err.Number <> ERR_NUMBER_SUCCESS Then Resume RaiseError
End Sub
'Sub EncryptByteData() - Passes a key and the byte data (as unicode)
'to the EncryptOrDecrypt Sub to encrypt the byte data.
Public Sub EncryptByteData(ByRef Key As String)
Attribute EncryptByteData.VB_Description = "Encrypt the byte data based on the key provided."
Attribute EncryptByteData.VB_HelpID = 1010
On Error Resume Next
  Dim strBytes As String
  
  'Take out the leading and trailing spaces..
  Key = Trim$(Key)
  
  'Check to make sure the key is not empty.
  If Len(Key) = 0 Then
    Call Err.Raise(ERR_CRYPT_EMPTYKEY, MODULE_NAME, "Key cannot be empty.")
    Exit Sub
  End If
  
  'Change our byte data to unicode.
  strBytes = StrConv(mvarByteData, vbUnicode)
  
  If Err.Number <> ERR_NUMBER_SUCCESS Then
    Call Err.Raise(Err.Number, MODULE_NAME, Err.Description)
    Exit Sub
  End If
  
  'Encrypt the data.
  Call EncryptOrDecrypt(strBytes, Key, True)
  
  'Move the data to our local byte data array.
  mvarByteData = strBytes
End Sub
'Sub DecryptByteData() - Passes a key and the byte data (as unicode)
'to the EncryptOrDecrypt Sub to decrypt the byte data. Then converts
'the data back to normal (from unicode).
Public Sub DecryptByteData(ByRef Key As String)
Attribute DecryptByteData.VB_Description = "Decrypt the byte data based on the key provided."
Attribute DecryptByteData.VB_HelpID = 1011
On Error Resume Next
  Dim strBytes As String
  
  'Take out the leading and trailing spaces..
  Key = Trim$(Key)
  
  'Check to make sure the key is not empty.
  If Len(Key) = 0 Then
    Call Err.Raise(ERR_CRYPT_EMPTYKEY, MODULE_NAME, "Key cannot be empty.")
    Exit Sub
  End If
  
  'Move the data to a string.
  strBytes = mvarByteData
  
  If Err.Number <> ERR_NUMBER_SUCCESS Then
    Call Err.Raise(Err.Number, MODULE_NAME, Err.Description)
    Exit Sub
  End If
  
  'Decrypt the data.
  Call EncryptOrDecrypt(strBytes, Key, False)
  
  'Convert the data back to normal. (From Unicode)
  mvarByteData = StrConv(strBytes, vbFromUnicode)
End Sub
'Sub SaveByteData() - Saves the byte data in memory to a specified
'file.
Public Sub SaveByteData(ByRef FileName As String)
On Error Resume Next
  Dim fn As Integer
  
  'Check if the file already exists.
  If FileExists(FileName) Then
    Call Err.Raise(ERR_FILE_EXISTS, MODULE_NAME, "File already exists.")
    Exit Sub
  End If
  
  'Get an available file handle, open the file, write the data to
  'it, and close the handle.
  fn = FreeFile
  Open FileName For Binary As #fn
    Put #fn, , mvarByteData
  Close #fn
  If Err.Number <> ERR_NUMBER_SUCCESS Then
    Call Err.Raise(Err.Number, Err.Source, Err.Description)
    Exit Sub
  End If
End Sub
'Sub LoadByteData() - Loads the byte data into memory from a specified
'file.
Public Sub LoadByteData(ByRef FileName As String)
On Error Resume Next
  Dim fn As Integer
  Dim fSize As Long
  
  'Get the file size.
  fSize = FileLen(FileName)
  If Err.Number <> ERR_NUMBER_SUCCESS Then
    Call Err.Raise(Err.Number, MODULE_NAME, Err.Description)
    Exit Sub
  End If
  
  'Redim our byte array to hold the data.
  ReDim mvarByteData(0 To fSize - 1)
  
  'Get an available file handle, open the file, read the data from
  'it, and close the handle.
  fn = FreeFile
  Open FileName For Binary As #fn
    Get #fn, , mvarByteData
  Close #fn
  If Err.Number <> ERR_NUMBER_SUCCESS Then
    Call Err.Raise(Err.Number, MODULE_NAME, Err.Description)
    Exit Sub
  End If
End Sub
'Sub SetByteData() - Copies the byte data from the memory pointer
'to our byte array based on the DataLen passed.
Public Sub SetByteData(ByVal MemPtr As Long, ByVal DataLen As Long)
Attribute SetByteData.VB_Description = "Copy the data from the memptr into byte data.\r\nUse this method when compressing or encrypting the byte data manually."
Attribute SetByteData.VB_HelpID = 1006
  If DataLen <= 0 Then
    Call Err.Raise(ERR_INVALID_LENGTH, MODULE_NAME, "Invalid length passed to SetByteData(). Data Length must be a positive value greater than 0.")
    Exit Sub
  End If
  ReDim mvarByteData(DataLen - 1)
  
  Call CopyMemory(mvarByteData(LBound(mvarByteData)), ByVal MemPtr, DataLen)
End Sub
'Function FileExists() - Check to see if a file exists.
Private Function FileExists(ByRef FileName As String) As Boolean
  FileExists = False
  
  If Not Dir(FileName) = "" Then FileExists = True
End Function
'Property Get ByteData() - Returns the starting location of our byte
'array.
Public Property Get ByteData() As Long
Attribute ByteData.VB_Description = "Pointer to memory location of the byte data."
Attribute ByteData.VB_HelpID = 1013
On Error Resume Next
  ByteData = VarPtr(mvarByteData(LBound(mvarByteData)))
  If Err.Number <> ERR_NUMBER_SUCCESS Then ByteData = 0
End Property
'Propert Get ByteDataLen() - Returns the length of bytes in our array.
Public Property Get ByteDataLen() As Long
Attribute ByteDataLen.VB_Description = "The length of the byte data."
Attribute ByteDataLen.VB_HelpID = 1014
On Error Resume Next
  ByteDataLen = (UBound(mvarByteData) - LBound(mvarByteData)) + 1
  If Err.Number <> ERR_NUMBER_SUCCESS Then
    ByteDataLen = 0
    Err.Clear
  End If
End Property
'Property Get ImageHeight() - Returns the height of our image in pixels.
Public Property Get ImageHeight() As Long
Attribute ImageHeight.VB_HelpID = 1016
On Error Resume Next
  Dim RetVal As Long
  
  'Check to make sure our header exists.
  Call CheckDataLen(MIN_HEADER_LEN)
  If Err.Number <> ERR_NUMBER_SUCCESS Then
    Call Err.Raise(Err.Number, Err.Source, Err.Description)
    Exit Property
  End If
  
  Call CopyMemory(RetVal, mvarByteData(LBound(mvarByteData) + BM_HEIGHT_OFFSET), 4)
  ImageHeight = RetVal
End Property
'Property Get ImageWidth() - Returns the width of our image in pixels.
Public Property Get ImageWidth() As Long
Attribute ImageWidth.VB_HelpID = 1015
On Error Resume Next
  Dim RetVal As Long
  
  'Check to make sure our header exists.
  Call CheckDataLen(MIN_HEADER_LEN)
  If Err.Number <> ERR_NUMBER_SUCCESS Then
    Call Err.Raise(Err.Number, Err.Source, Err.Description)
    Exit Property
  End If
  
  Call CopyMemory(RetVal, mvarByteData(LBound(mvarByteData) + BM_WIDTH_OFFSET), 4)
  ImageWidth = RetVal
End Property
'Sub EraseByteData() - Free up the memory used by our byte array.
Public Sub EraseByteData()
Attribute EraseByteData.VB_Description = "Free the memory used by the byte data."
Attribute EraseByteData.VB_HelpID = 1013
  Erase mvarByteData
End Sub
'Terminate the class. :)
Private Sub Class_Terminate()
  EraseByteData
End Sub
