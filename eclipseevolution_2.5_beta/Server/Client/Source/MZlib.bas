Attribute VB_Name = "MZlib"
Option Explicit

' Used to compress and decompress packets. Cheers to Phil Fresle for the wrapper which
'   helped greatly in the making of this.
'     Packet Structure:
'     COMPRESSED-LENGTH:DATA+
'               ^ SEP_CHAR  ^ END_CHAR
'     COMPRESSED flag - 1 for compressed, 0 for not
'     LENGTH - length of decompressed data

'TCP DATA CONSTANTS
Public Const DATA_UNCOMPRESSED As Integer = 0
Public Const DATA_COMPRESSED As Integer = 1

Private Declare Function compress Lib "zLib.dll" _
        (ByVal compr As String, comprLen As Any, _
        ByVal buf As String, ByVal buflen As Long) As Long
        
Private Declare Function uncompress Lib "zLib.dll" _
        (ByVal uncompr As String, uncomprLen As Any, _
        ByVal compr As String, ByVal lcompr As Long) As Long


'*******************************************
' CompressPacket - returns compressed packet
'*******************************************
Public Function CompressPacket(ByVal Data As String) As String
    'Error handlin'
    If Len(Data) <= 0 Then Err.Raise vbObjectError + 2, "compressPacket", "No data to compress"

    Dim cData As String 'Compressed data
    Dim cDataLen As Long 'Length of compressed string
    Dim dataLen As Long 'Length of original data
    
    dataLen = Len(Data)
    cDataLen = (dataLen * 1.01) + 13
    cData = Space(cDataLen)
    
    If compress(cData, cDataLen, Data, dataLen) <> 0 Then
        'Error handlin'
        Err.Raise vbObjectError + 1, "compressPacket", "Error while compressing packet"
    Else
        CompressPacket = cData
    End If
End Function

'**********************************************
' DeompressPacket - returns decompressed packet
'**********************************************
Public Function DecompressPacket(ByVal Data As String) As String
    Dim cData As String 'Raw compressed data
    Dim retrn As Long
    
    'Error handlin'
    If Len(Data) <= 0 Then Err.Raise vbObjectError + 2, "compressPacket", "No data to compress"
    
    Dim dataLen As Long 'Length of uncompressed string
    Dim cLen As Long 'Length of compressed string
    Dim decompressedStr As String
    
    dataLen = Val(Mid$(Data, InStr(1, Data, SEP_CHAR) + 1))
    cData = Mid$(Data, InStr(1, Data, ":") + 1)
    cLen = Len(cData)
    decompressedStr = Space(dataLen)
    
    retrn = uncompress(decompressedStr, dataLen, cData, cLen)
    
    If retrn <> 0 Then
        If retrn = -3 Then
            'The data was not compressed, so return the regular data
            DecompressPacket = Data
        Else
            'Error handlin'
            Err.Raise vbObjectError + 1, "compressPacket", "Error while compressing packet"
        End If
    Else
        DecompressPacket = decompressedStr
    End If
End Function

'**************************************************
' PacketIsCompressed - checks for compession flag
'**************************************************
Public Function PacketIsCompressed(ByVal Packet As String) As Boolean
    'It's only compressed if the first value is 1, so if we get an error, just return false.
    PacketIsCompressed = False
    
    On Error Resume Next
    PacketIsCompressed = (Mid$(Packet, 1, InStr(1, Packet, SEP_CHAR, vbBinaryCompare)) = DATA_COMPRESSED)
End Function


