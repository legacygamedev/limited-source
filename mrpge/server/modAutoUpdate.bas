Attribute VB_Name = "modAutoUpdate"
'REDUNDANT

Option Explicit
Dim strFilename(MAX_PLAYERS) As String
Dim lngFilePos(MAX_PLAYERS) As Long
Dim blnTransfering(MAX_PLAYERS) As Boolean
Dim lngFileLength(MAX_PLAYERS) As Long
Dim packetCounter(MAX_PLAYERS) As Long
Private Const CHUNK_SIZE = 3096
Dim hFile As Long
Public UPDATER_END_CHAR As String
Dim num_of_Packets As Double
Dim filenameLength As Long
Dim packetsize As Long

' packet layouts
' NEXT packet_num
' EEOF
' FEOF
' BBOF max_packets length_of_filename file_name file_size
' DATA packet_num packet_size packet

' packet_num has length of 4
' max_packets has a length of 4
' packet_size has a length of 4
' file_size has length 4 (size in KB)
' length_of_filename length = 3
    
Sub update_socketclose(ByVal index As Integer)
    strFilename(index) = ""
    lngFilePos(index) = 0
    blnTransfering(index) = False
End Sub

Sub update_sendData(index As Integer, strdata As String)
Dim packet As String
    packet = strdata '& Chr(0)
    'frmServer.UpdaterSocket(index).SendData (packet)
End Sub


Sub Update_AcceptConnection(index As Integer, ByVal requestID As Long)
Dim i As Long

    If (index = 0) Then
        i = Update_findOpenSlot
        
        If i <> 0 Then
            ' Whoho, we can connect them
            'frmServer.UpdaterSocket(i).Close
            'frmServer.UpdaterSocket(i).Accept requestID
        End If
    End If
End Sub
Sub Update_IncomingData(index As Integer, ByVal bytesTotal As Long)
'On Error Resume Next
Dim splitArr() As String
Dim Buffer As String
Dim packet As String
Dim top As String * 3
Dim Start As Integer
Dim i As Long
    If index > 0 Then
        'frmServer.UpdaterSocket(index).GetData Buffer, vbString, bytesTotal
        'splitArr = Split(Buffer, Chr(0))
        'For i = 0 To UBound(splitArr)
        'Buffer = splitArr(i)
            If Mid$(Buffer, 1, 4) = "NEXT" Then
                lngFilePos(index) = CLng(Mid$(Buffer, 5, 4)) * CHUNK_SIZE
                Call SendNextChunk(index, CLng(Mid$(Buffer, 5, 4)))
            ElseIf Buffer = "NUMFILES" Then
                Call update_sendData(index, "NUMFILES25")
            ElseIf Buffer = "GETFILE" Then
                Call sendFiles(index)
            Else
                'Debug.Print Buffer
            End If
        'Next i
    End If
End Sub

Private Function Update_findOpenSlot()
    Dim i As Integer

    Update_findOpenSlot = 0
    
    For i = 1 To MAX_PLAYERS
        If Not Update_IsConnected(i) Then
            Update_findOpenSlot = i
            Exit Function
        End If
    Next i
End Function

Private Function Update_IsConnected(index As Integer)
    'If frmServer.UpdaterSocket(index).State = sckConnected Then
        Update_IsConnected = True
    'Else
        Update_IsConnected = False
    'End If
End Function

Private Sub sendFiles(index As Integer, Optional ByVal fName As String = "\client\client.exe")
strFilename(index) = fName
    blnTransfering(index) = True
    If fName = "EOF" Then
        Call update_sendData(index, "BEOF")
        'frmServer.UpdaterSocket(index).SendData ("EOF")
        DoEvents
    ElseIf fName = "EEOF" Then
        Call update_sendData(index, "EEOF")
        'frmServer.UpdaterSocket(index).SendData ("EEOF")
        DoEvents
    Else
        'BBOF max_packets length_of_filename file_name file_size
        Close #hFile
        hFile = FreeFile
        Open App.Path & strFilename(index) For Binary As #hFile
            lngFileLength(index) = LOF(hFile)
        Close #hFile
        DoEvents
        num_of_Packets = (lngFileLength(index)) / CHUNK_SIZE
        
        If num_of_Packets > Int(num_of_Packets) Then num_of_Packets = Int(num_of_Packets) + 1
        
        filenameLength = addLeadingZeros(Len(Right(strFilename(index), Len(strFilename(index)) - 7)), 3)
        
        Call update_sendData(index, "BBOF" & addLeadingZeros(num_of_Packets, 4) & addLeadingZeros(filenameLength, 3) & Right(strFilename(index), Len(strFilename(index)) - 7) & addLeadingZeros(lngFileLength(index), 4))
        'Call update_sendData(index, "BBOF" & Right(strFilename(index), Len(strFilename(index)) - 7))
        'frmServer.UpdaterSocket(index).SendData ("BOF" & Right(strFilename(index), Len(strFilename(index)) - 7))
        DoEvents
    End If
End Sub

Private Sub SendNextChunk(index As Integer, Optional ByVal packet As Long)
'Debug.Print packet
On Error GoTo error:
    Dim lngChunksize As Long
    Dim strdata As String
    Dim strFile As String
    strFilename(index) = strFilename(index)
    'make sure prev file is closed
    Close #hFile
    DoEvents
    
    hFile = FreeFile
    strFile = App.Path & strFilename(index)
    Open strFile For Binary As #hFile
    'Debug.Print "file: " & strFile
    lngFilePos(index) = packet * CHUNK_SIZE + 1
        If (lngFilePos(index) = 0) Then lngFilePos(index) = 1
        
        Seek hFile, lngFilePos(index)
        
        lngChunksize = LOF(hFile) + 1 - lngFilePos(index)
        If (lngChunksize > CHUNK_SIZE) Then lngChunksize = CHUNK_SIZE
        'Debug.Print lngChunksize
        If (lngChunksize <= 0) Then
            'strdata = "EOF"
            Call update_sendData(index, "BEOF")
            'frmServer.UpdaterSocket(index).SendData (strdata)
            DoEvents
            
            lngFilePos(index) = 0
            Call sendFiles(index, getNextFile(index))
            'frmServer.txtUpdater.Text = "EOF"
            DoEvents
            'Close #hFile
            'Debug.Print "closeOK"
            'DoEvents
            'Exit Sub
        Else
            ' DATA packet_num packet_size packet
            strdata = String$(lngChunksize, 0)
            Get #hFile, , strdata
            lngFilePos(index) = lngFilePos(index) + lngChunksize
            packetsize = addLeadingZeros(Len(strdata), 4)
            Call update_sendData(index, "DATA" & addLeadingZeros(packet, 4) & addLeadingZeros(packetsize, 4) & strdata)
            'Call frmServer.UpdaterSocket(index).SendData(strdata)
            DoEvents
            'frmServer.txtUpdater.Text = "NEXT"
            'Close #hFile
            'Debug.Print "closeOK"
            'DoEvents
        End If
    Close #hFile
    Exit Sub
error:
    If Err.number = 67 Then
        Close #hFile
        'Debug.Print "error"
        'Resume
        Exit Sub
    End If
    'MsgBox Err.Description, , Err.Number
End Sub
Private Function getNextFile(index As Integer)
    Select Case strFilename(index)
        Case Is = "\client\client.exe"
            strFilename(index) = "\client\fmod.dll"
            'strFilename(index) = "EEOF"
        Case Is = "\client\fmod.dll"
            strFilename(index) = "Tiles0.bmp"
        Case Is = "\client\data\bmp\Tiles0.bmp"
            strFilename(index) = "Tiles1.bmp"
        Case Is = "\client\data\bmp\Tiles1.bmp"
            strFilename(index) = "Tiles2.bmp"
        Case Is = "\client\data\bmp\Tiles2.bmp"
            strFilename(index) = "Tiles3.bmp"
        Case Is = "\client\data\bmp\Tiles3.bmp"
            strFilename(index) = "Tiles4.bmp"
        Case Is = "\client\data\bmp\Tiles4.bmp"
            strFilename(index) = "Tiles5.bmp"
        Case Is = "\client\data\bmp\Tiles5.bmp"
            strFilename(index) = "Tiles6.bmp"
        Case Is = "\client\data\bmp\Tiles6.bmp"
            strFilename(index) = "Tiles7.bmp"
        Case Is = "\client\data\bmp\Tiles7.bmp"
            strFilename(index) = "Tiles8.bmp"
        Case Is = "\client\data\bmp\Tiles8.bmp"
            strFilename(index) = "Tiles9.bmp"
        Case Is = "\client\data\bmp\Tiles9.bmp"
            strFilename(index) = "Tiles10.bmp"
        Case Is = "\client\data\bmp\Tiles10.bmp"
            strFilename(index) = "Tiles11.bmp"
        Case Is = "\client\data\bmp\Tiles11.bmp"
            strFilename(index) = "Tiles12.bmp"
        Case Is = "\client\data\bmp\Tiles12.bmp"
            strFilename(index) = "Tiles13.bmp"
        Case Is = "\client\data\bmp\Tiles13.bmp"
            strFilename(index) = "Tiles14.bmp"
        Case Is = "\client\data\bmp\Tiles14.bmp"
            strFilename(index) = "Tiles15.bmp"
        Case Is = "\client\data\bmp\Tiles15.bmp"
            strFilename(index) = "Tiles16.bmp"
        Case Is = "\client\data\bmp\Tiles16.bmp"
            strFilename(index) = "Tiles17.bmp"
        Case Is = "\client\data\bmp\Tiles17.bmp"
            strFilename(index) = "Tiles18.bmp"
        Case Is = "\client\data\bmp\Tiles18.bmp"
            strFilename(index) = "Tiles19.bmp"
        Case Is = "\client\data\bmp\Tiles19.bmp"
            strFilename(index) = "Items.bmp"
        Case Is = "\client\data\bmp\Items.bmp"
            strFilename(index) = "Sprites.bmp"
        Case Is = "\client\data\bmp\Sprites.bmp"
            strFilename(index) = "night.bmp"
        Case Else
            strFilename(index) = "EEOF"
    End Select
    If strFilename(index) <> "EEOF" And strFilename(index) <> "\client\fmod.dll" Then
        strFilename(index) = "\client\data\bmp\" & strFilename(index)
    End If
    getNextFile = strFilename(index)
    DoEvents
End Function


Function addLeadingZeros(ByVal num As String, ByVal total_digits As Long) As String
    If Len(num) < total_digits Then
        While Len(num) < total_digits
            num = "0" & num
        Wend
    End If
    addLeadingZeros = num
End Function


