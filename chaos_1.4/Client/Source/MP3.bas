Attribute VB_Name = "MP3"

Option Explicit


'constants
Private Const FLAG_HEADER_UNSYNC = &H80
Private Const FLAG_HEADER_EXTENDED = &H40
Private Const FLAG_HEADER_EXPERIMENTAL = &H20
Private Const FLAG_HEADER_FOOTER = &H10

Private Const FLAG_EXTENDED_UPDATE = &H40
Private Const FLAG_EXTENDED_CRC32 = &H20
Private Const FLAG_EXTENDED_RESTRICT = &H10

Private Const FLAG_RESTRICT_64FRAMES128KB = &H40        'even if no size restrict is set, there is a retriction to 128 frames
Private Const FLAG_RESTRICT_32FRAMES40KB = &H80         'and a total of 1 MB of tag size when restrict flag is set
Private Const FLAG_RESTRICT_32FRAMES4KB = &HC0          'without the restriction flag the tag size limit is 256 MB
Private Const FLAG_RESTRICT_ISO88591UTF8ONLY = &H20
Private Const FLAG_RESTRICT_1024CHARACTERS = &H8
Private Const FLAG_RESTRICT_128CHARACTERS = &H10
Private Const FLAG_RESTRICT_30CHARACTERS = &H18
Private Const FLAG_RESTRICT_PNGJPEGONLY = &H4
Private Const FLAG_RESTRICT_256PIXELS = &H1
Private Const FLAG_RESTRICT_64PIXELS = &H2
Private Const FLAG_RESTRICT_64PIXELSONLY = &H3

Private Const FLAG_FORMAT_GROUP = &H40
Private Const FLAG_FORMAT_COMPRESSED = &H8
Private Const FLAG_FORMAT_ENCRYPT = &H4
Private Const FLAG_FORMAT_UNSYNC = &H2
Private Const FLAG_FORMAT_DLI = &H1

Private Const FLAG_STATUS_TAGALTER = &H40
Private Const FLAG_STATUS_FILEALTER = &H20
Private Const FLAG_STATUS_READONLY = &H10


'status messages
Public Enum MP3_CLASS_STATUS
    mp3classerror = -1              'an error has occurred
    mp3classunknown = 0             'unknown status
    mp3classnotready = 1            'MCI not ready
    mp3classplaying = 2             'MCI playing
    mp3classpaused = 3              'MCI paused
    mp3classstopped = 4             'MCI stopped
    mp3classseeking = 5             'MCI seeking
    mp3classrecording = 6           'MCI recording
End Enum

'id3v2 restricted info
Public Enum ID3V2_RESTRICT_SIZE
    id3restrictsizenone = 0
    id3restrict1MB = 1              'tag size 1MB with 128 frames max
    id3restrict128KB = 2            'tag size 128KB with 64 frames max
    id3restrict40KB = 3             'tag size 40KB with 32 frames max
    id3restrict4KB = 4              'tag size 4KB with 32 frames max
End Enum

Public Enum ID3V2_RESTRICT_CHARS
    id3restrictcharsnone = 0
    id3restrict1024 = 1             'limit frame data characters to 1024
    id3restrict128 = 2              'limit frame data characters to 128
    id3restrict30 = 3               'limit frame data characters to 30
End Enum

Public Enum ID3V2_RESTRICT_IMAGE
    id3restrictimagenone = 0
    id3restrict256 = 1              'image size up to 256 x 256 pixels
    id3restrict64 = 2               'image size up to 64 x 64 pixels
    id3restrict64only = 3           'image size only 64 x 64 pixels except if other size is demanded
End Enum


'ID3v1 tag format

'helps loading and saving data
Private Type ID3v1_TAG              '(128 bytes)
    Tag As String * 3               'always TAG
    Title As String * 30            'title, 30 characters
    Artist As String * 30           'artist, 30 characters
    Album As String * 30            'album, 30 characters
    Year As String * 4              'year, 4 characters
    Comment As String * 30          'comment, 30 characters (or 28 if track# included)
    Genre As Byte                   'genre, 255 for none defined
End Type

'storaging data for use
Private Type ID3v1_STORE
    Title As String                 'title
    Artist As String                'artist
    Album As String                 'album
    Year As Integer                 'year
    Comment As String               'comment
    Genre As Byte                   'genre (use GetGenre for text)
    Track As Byte                   'track
End Type


'see www.id3.org for complete description on ID3v2 tag format

'helps loading and saving data
Private Type ID3v2_HEADERFOOTER     '(10 bytes)
    Tag As String * 3               'always ID3 in header and 3DI in footer; footer is a copy of header
    Version As Byte                 'ID3v2 version
    Minor As Byte                   'ID3v2 minor version
    Flag As Byte                    'bit 1: unsynchronisation
                                    'bit 2: extended header
                                    'bit 3: experimental
                                    'bit 4: footer
                                    'bit 5-8: unused flags
    Size As Long                    'size of the tag: 4 synchsafe bytes, first bit of a byte always zero
End Type

'helps loading and saving data
Private Type ID3v2_EXTENDEDHEADER   '(6 bytes)
    Size As Long                    'size of the extended header: 4 synchsafe bytes, first bit of a byte always zero
    Flags As Byte                   'number of flag bytes
    Flag As Byte                    'bit 1: unused flag
                                    'bit 2: tag is an update
                                    'bit 3: CRC-32 data available
                                    'bit 4: tag restrictions
                                    'bit 5-8: unused flags
End Type

'helps loading and saving data
Private Type ID3v2_FRAMEHEADER      '(10 bytes)
    ID As String * 4                'characters always A-Z and 0-9; first character X, Y and Z reserved for experimental tags
    Size As Long                    'size of the header: 4 synchsafe bytes, first bit of a byte always zero
    StatusFlag As Byte              'bit 1: unused flag
                                    'bit 2: tag alternation preserve
                                    'bit 3: file alternation preserve
                                    'bit 4: read only
                                    'bit 5-8: unused flags
    FormatFlag As Byte              'bit 1: unused flag
                                    'bit 2: group member
                                    'bit 3,4: unused flags
                                    'bit 5: compression
                                    'bit 6: encryption
                                    'bit 7: unsynchronisation
                                    'bit 8: data length indicator
End Type

'storaging data for use
Private Type ID3v2_HEADERDATA
    Size As Long                    'tag size (without header & footer)
    Version As String * 3           'tag version
    FlagUnsync As Boolean           'unsynchronisation
    FlagExtended As Boolean         'tag has extended header
    FlagExperimental As Boolean     'tag is experimental
    FlagFooter As Boolean           'tag has a footer
    FlagUpdate As Boolean           'tag is an update
    FlagCRC32 As Boolean            'extended header includes CRC-32
    FlagRestricted As Boolean       'tag restrictions
    CRC32 As Long                   'CRC-32 value
    ResSize As ID3V2_RESTRICT_SIZE  'restrict size
    ResCharCode As Boolean          'restrict character coding
    ResChar As ID3V2_RESTRICT_CHARS 'restrict characters in frame data
    ResPngJpeg As Boolean           'restrict images to PNG and JPEG
    ResImg As ID3V2_RESTRICT_IMAGE  'restrict image size
    Corrupt As Boolean              'is corrupt according to CRC-32
End Type

Private Type ID3v2_FRAMEDATA
    ID As String * 4                'characters always A-Z and 0-9; first character X, Y and Z reserved for experimental tags
    Data As String                  'the raw, uncompressed and undecrypted data
    TagAlter As Boolean             'preserve tag if altered and not known what it is for
    FileAlter As Boolean            'preserve tag if file is altered and not known what it is for
    ReadOnly As Boolean             'tag is readonly and should not be modified or removed
    GroupMember As Boolean          'is a member of a group
    GroupID As Byte                 'group byte id
    Compressed As Boolean           'is data compressed
    Encrypted As Boolean            'is data encrypted
    Unsync As Boolean               'is data unsynchronized
    DLI As Boolean                  'data length indicator
    Length As Long                  'data length
End Type


'API declarations

Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

'variables

Private CRC32CheckSumTable(0 To 255) As Long    'CRC-32 bytewise checksum table

Private ErrorsList As New Collection            'for storing the errors

Private filename As String                      'the filename
Private FileOpened As Boolean                   'to know in GetStatus if a file is open or not
                                                '(to prevent extra errors)

Private ID3v1genres(255) As String              'ID3v1 genrelist storage for fast access
Private ID3v1store As ID3v1_STORE               'ID3v1 tag information

Private ID3v2data() As ID3v2_FRAMEDATA          'ID3v2 tag information
Private ID3v2extended As ID3v2_EXTENDEDHEADER   'ID3v2 tag extended header
Private ID3v2extendedData() As Byte             'ID3v2 tag extended header data
Private ID3v2header As ID3v2_HEADERFOOTER       'ID3v2 tag header
Private ID3v2info As ID3v2_HEADERDATA           'ID3v2 tag header information


'mp3 class initialize
Private Sub Class_Initialize()
    Dim A As Integer, B As Byte, C As Long
    Const CRC32limit = &HEDB88320
    'init CRC-32 bytewise checksum table
    For A = 0 To 255 Step 1
        C = A
        For B = 0 To 7 Step 1
            If C And 1 Then
                C = (((C And &HFFFFFFFE) \ 2) And &H7FFFFFFF) Xor CRC32limit
            Else
                C = ((C And &HFFFFFFFE) \ 2) And &H7FFFFFFF
            End If
        Next B
        'add the checksum of the byte to the table
        CRC32CheckSumTable(A) = C
    Next A
    'init ID3v2 tag
    ReDim ID3v2data(0)
    'initialize genres for fast access
    'based on WinAmp genres
    ID3v1genres(147) = "Synthpop"
    ID3v1genres(146) = "JPop"
    ID3v1genres(145) = "Anime"
    ID3v1genres(144) = "Thrash Metal"
    ID3v1genres(143) = "Salsa"
    ID3v1genres(142) = "Merengue"
    ID3v1genres(141) = "Christian Rock"
    ID3v1genres(140) = "Contemporary Christian"
    ID3v1genres(139) = "Crossover"
    ID3v1genres(138) = "Black Metal"
    ID3v1genres(137) = "Heavy Metal"
    ID3v1genres(136) = "Christian Gangsta Rap"
    ID3v1genres(135) = "Beat"
    ID3v1genres(134) = "Polsk Punk"
    ID3v1genres(133) = "Negerpunk"
    ID3v1genres(132) = "BritPop"
    ID3v1genres(131) = "Indie"
    ID3v1genres(130) = "Terror"
    ID3v1genres(129) = "HardMain"
    ID3v1genres(128) = "Club-House"
    ID3v1genres(127) = "Drum & Bass"
    ID3v1genres(126) = "Goa"
    ID3v1genres(125) = "Dance Hall"
    ID3v1genres(124) = "Euro-House"
    ID3v1genres(123) = "A Cappella"
    ID3v1genres(122) = "Drum Solo"
    ID3v1genres(121) = "Punk Rock"
    ID3v1genres(120) = "Duet"
    ID3v1genres(119) = "Freestyle"
    ID3v1genres(118) = "Rhythmic Soul"
    ID3v1genres(117) = "Power Ballad"
    ID3v1genres(116) = "Ballad"
    ID3v1genres(115) = "Folklore"
    ID3v1genres(114) = "Samba"
    ID3v1genres(113) = "Tango"
    ID3v1genres(112) = "Club"
    ID3v1genres(111) = "Slow Jam"
    ID3v1genres(110) = "Satire"
    ID3v1genres(109) = "Porn Groove"
    ID3v1genres(108) = "Primus"
    ID3v1genres(107) = "Booty Bass"
    ID3v1genres(106) = "Symphony"
    ID3v1genres(105) = "Sonata"
    ID3v1genres(104) = "Chamber Music"
    ID3v1genres(103) = "Opera"
    ID3v1genres(102) = "Chanson"
    ID3v1genres(101) = "Speech"
    ID3v1genres(100) = "Humour"
    ID3v1genres(99) = "Acoustic"
    ID3v1genres(98) = "Easy Listening"
    ID3v1genres(97) = "Chorus"
    ID3v1genres(96) = "Big Band"
    ID3v1genres(95) = "Slow Rock"
    ID3v1genres(94) = "Symphonic Rock"
    ID3v1genres(93) = "Psychedelic Rock"
    ID3v1genres(92) = "Progressive Rock"
    ID3v1genres(91) = "Gothic Rock"
    ID3v1genres(90) = "Avantgarde"
    ID3v1genres(89) = "Bluegrass"
    ID3v1genres(88) = "Celtic"
    ID3v1genres(87) = "Revival"
    ID3v1genres(86) = "Latin"
    ID3v1genres(85) = "Bebob"
    ID3v1genres(84) = "Fast-Fusion"
    ID3v1genres(83) = "Swing"
    ID3v1genres(82) = "National Folk"
    ID3v1genres(81) = "Folk/Rock"
    ID3v1genres(80) = "Folk"
    ID3v1genres(79) = "Hard Rock"
    ID3v1genres(78) = "Rock & Roll"
    ID3v1genres(77) = "Musical"
    ID3v1genres(76) = "Retro"
    ID3v1genres(75) = "Polka"
    ID3v1genres(74) = "Acid Jazz"
    ID3v1genres(73) = "Acid Punk"
    ID3v1genres(72) = "Tribal"
    ID3v1genres(71) = "Lo-Fi"
    ID3v1genres(70) = "Trailer"
    ID3v1genres(69) = "Showtunes"
    ID3v1genres(68) = "Rave"
    ID3v1genres(67) = "Psychedelic"
    ID3v1genres(66) = "New Wave"
    ID3v1genres(65) = "Cabaret"
    ID3v1genres(64) = "Native American"
    ID3v1genres(63) = "Jungle"
    ID3v1genres(62) = "Pop/Funk"
    ID3v1genres(61) = "Christian Rap"
    ID3v1genres(60) = "Top 40"
    ID3v1genres(59) = "Gangsta Rap"
    ID3v1genres(58) = "Cult"
    ID3v1genres(57) = "Comedy"
    ID3v1genres(56) = "Southern Rock"
    ID3v1genres(55) = "Dream"
    ID3v1genres(54) = "Eurodance"
    ID3v1genres(53) = "Pop-Folk"
    ID3v1genres(52) = "Electronic"
    ID3v1genres(51) = "Techno-Industrial"
    ID3v1genres(50) = "Darkwave"
    ID3v1genres(49) = "Gothic"
    ID3v1genres(48) = "Ethnic"
    ID3v1genres(47) = "Instrumental Rock"
    ID3v1genres(46) = "Instrumental Pop"
    ID3v1genres(45) = "Meditative"
    ID3v1genres(44) = "Space"
    ID3v1genres(43) = "Punk"
    ID3v1genres(42) = "Soul"
    ID3v1genres(41) = "Bass"
    ID3v1genres(40) = "Alt. Rock"
    ID3v1genres(39) = "Noise"
    ID3v1genres(38) = "Gospel"
    ID3v1genres(37) = "Sound Clip"
    ID3v1genres(36) = "Game"
    ID3v1genres(35) = "House"
    ID3v1genres(34) = "Acid"
    ID3v1genres(33) = "Instrumental"
    ID3v1genres(32) = "Classical"
    ID3v1genres(31) = "Trance"
    ID3v1genres(30) = "Fusion"
    ID3v1genres(29) = "Jazz+Funk"
    ID3v1genres(28) = "Vocal"
    ID3v1genres(27) = "Trip-Hop"
    ID3v1genres(26) = "Ambient"
    ID3v1genres(25) = "Euro-Techno"
    ID3v1genres(24) = "Soundtrack"
    ID3v1genres(23) = "Pranks"
    ID3v1genres(22) = "Death Metal"
    ID3v1genres(21) = "Ska"
    ID3v1genres(20) = "Alternative"
    ID3v1genres(19) = "Industrial"
    ID3v1genres(18) = "Techno"
    ID3v1genres(17) = "Rock"
    ID3v1genres(16) = "Reggae"
    ID3v1genres(15) = "Rap"
    ID3v1genres(14) = "R&B"
    ID3v1genres(13) = "Pop"
    ID3v1genres(12) = "Other"
    ID3v1genres(11) = "Oldies"
    ID3v1genres(10) = "New Age"
    ID3v1genres(9) = "Metal"
    ID3v1genres(8) = "Jazz"
    ID3v1genres(7) = "Hip-Hop"
    ID3v1genres(6) = "Grunge"
    ID3v1genres(5) = "Funk"
    ID3v1genres(4) = "Disco"
    ID3v1genres(3) = "Dance"
    ID3v1genres(2) = "Country"
    ID3v1genres(1) = "Classic Rock"
    ID3v1genres(0) = "Blues"
End Sub


'MP3 class error handling
Private Sub AddError(ByVal Where As String, ByVal Number As Long, ByVal Description As String)
    'add the error always as the first item
    If ErrorsList.Count Then
        ErrorsList.Add Where & "|" & Number & "|" & Trim$(Replace$(Description, Chr(0), "")), , 1
    Else
        ErrorsList.Add Where & "|" & Number & "|" & Trim$(Replace$(Description, Chr(0), ""))
    End If
End Sub
Private Function CheckError(ByVal Where As String, ByVal ErrorNumber As Long) As Boolean
    Dim ErrorString As String * 128
    'returns true if no error
    CheckError = (ErrorNumber = 0)
    'get the error if occurred
    If Not CheckError Then
        'get error string
        mciGetErrorString ErrorNumber, ErrorString, Len(ErrorString)
        'store error
        AddError Where, ErrorNumber, ErrorString
    End If
End Function
Public Sub ClearErrors()
    'remove all errors
    Do While ErrorsList.Count: ErrorsList.Remove 1: Loop
End Sub
Public Function Errors() As Integer
    'return the number of errors
    Errors = ErrorsList.Count
End Function
Public Function GetError(ByVal Index As Integer) As String
    'check for valid index
    If Index < 1 Or Index > ErrorsList.Count Then Exit Function
    'return error
    GetError = ErrorsList(Index)
End Function
Public Function GetLastError() As String
    'return last error
    If ErrorsList.Count Then GetLastError = ErrorsList(1)
End Function


'MP3 class file handling
Public Function CloseMP3() As Boolean
    'close file
    CloseMP3 = CheckError("CloseMP3", mciSendString("close MP3Play", 0&, 0&, 0&))
    FileOpened = False
End Function
Public Function GetFilename() As String
    Dim FileLength As Integer
    'get length
    FileLength = Len(filename)
    'check if long enough to process and return
    If FileLength > 2 Then
        'remove quotes and return the filename
        GetFilename = mid$(filename, 2, FileLength - 2)
    End If
End Function
Public Function GetStatus() As MP3_CLASS_STATUS
    Dim Temp As String * 30
    'check if there is a file open: if not, report not ready
    If Not FileOpened Then GetStatus = mp3classnotready: Exit Function
    'get status and return value accordingly
    If CheckError("GetStatus", mciSendString("status MP3Play mode", Temp, Len(Temp), 0&)) Then
        Select Case LCase(Trim$(Replace(Temp, Chr(0), " ")))
            Case "not ready"
                GetStatus = mp3classnotready
            Case "playing"
                GetStatus = mp3classplaying
            Case "paused"
                GetStatus = mp3classpaused
            Case "stopped"
                GetStatus = mp3classstopped
            Case "seeking"
                GetStatus = mp3classseeking
            Case "recording"
                GetStatus = mp3classrecording
            Case Else
                GetStatus = mp3classunknown
        End Select
    Else
        GetStatus = mp3classerror
    End If
End Function
Public Function OpenMP3(ByVal File As String) As Boolean
    Dim A As Long
    'if there is no file defined, exit
    If File = "" Then Exit Function
    'skip error
    On Error Resume Next
    'check for file length
    A = FileLen(File)
    'on error or zero length file, exit
    If Err Or A = 0 Then Exit Function
    'don't skip errors
    On Error GoTo 0
    'set filename
    filename = Chr(34) & File & Chr(34)
    'get tags
    GetID3v1 File
    GetID3v2 File
    'open the file
    OpenMP3 = CheckError("OpenMP3", mciSendString("open " & filename & " type MPEGVideo Alias MP3Play", 0&, 0&, 0&))
    If Not OpenMP3 Then
        CloseMP3
    Else
        FileOpened = True
    End If
End Function


' MP3 class playback handling
Public Function PauseMP3() As Boolean
    'pause playback
    PauseMP3 = CheckError("PauseMP3", mciSendString("pause MP3Play", 0&, 0&, 0&))
End Function
Public Function PlayMP3() As Boolean
    'play the file
    PlayMP3 = CheckError("PlayMP3", mciSendString("play MP3Play", 0&, 0&, 0&))
End Function
Public Function ResumeMP3() As Boolean
    'resume playback
    ResumeMP3 = CheckError("ResumeMP3", mciSendString("resume MP3Play", 0&, 0&, 0&))
End Function
Public Function StopMP3() As Boolean
    'stop playback
    StopMP3 = CheckError("StopMP3", mciSendString("stop MP3Play", 0&, 0&, 0&))
    'seek to beginning
    If StopMP3 Then SeekTo 0
End Function


'MP3 class playback position handling
Public Function Length() As Long
    Dim Temp As String * 30
    'make sure we use milliseconds
    If CheckError("Length", mciSendString("set MP3Play time format milliseconds", 0&, 0&, 0&)) Then
        If CheckError("Length", mciSendString("status MP3Play length", Temp, Len(Temp), 0&)) Then
            'return length in milliseconds
            Length = CLng(Temp)
        Else
            'return true (= -1)
            Length = True
        End If
    Else
        'return true (= -1)
        Length = True
    End If
End Function
Public Function Position() As Long
    Dim Temp As String * 30
    'make sure we use milliseconds
    If CheckError("Position", mciSendString("set MP3Play time format milliseconds", 0&, 0&, 0&)) Then
        If CheckError("Position", mciSendString("status MP3Play position", Temp, Len(Temp), 0&)) Then
            'return position in milliseconds
            Position = CLng(Temp)
        Else
            'return true (= -1)
            Position = True
        End If
    Else
        'return true (= -1)
        Position = True
    End If
End Function
Public Function Remaining() As Long
    Dim MP3pos As Long, MP3len As Long
    'get position
    MP3pos = Position
    'error, exit
    If MP3pos = True Then Remaining = True: Exit Function
    'get length
    MP3len = Length
    'error, exit
    If MP3len = True Then Remaining = True: Exit Function
    'return remaining
    Remaining = MP3len - MP3pos
End Function
Public Function SeekTo(ByVal NewPosition As Long) As Boolean
    'make sure we use milliseconds
    If CheckError("SeekTo", mciSendString("set MP3Play time format milliseconds", 0&, 0&, 0&)) Then
        'if playing...
        If GetStatus = mp3classplaying Then
            SeekTo = CheckError("SeekTo", mciSendString("play MP3Play from " & CStr(NewPosition), 0&, 0&, 0&))
        Else 'if not playing...
            SeekTo = CheckError("SeekTo", mciSendString("seek MP3Play to " & CStr(NewPosition), 0&, 0&, 0&))
        End If
    Else
        'error
        SeekTo = False
    End If
End Function


'mp3 class id3v1 handling
Public Function GetGenre(ByVal Index As Byte) As String
    'return genre from genres list
    GetGenre = ID3v1genres(Index)
End Function
Private Sub GetID3v1(ByVal File As String)
    Dim IDtag As ID3v1_TAG
    On Error GoTo ErrorHandler
    'open file for read
    Open File For Binary Access Read As #1
        'check if file is big enough for it to contain a tag
        If LOF(1) < Len(IDtag) Then Close #1: Exit Sub
        'read the tag
        Get #1, LOF(1) - Len(IDtag) + 1, IDtag
    Close #1
    'check if there is a tag
    If IDtag.Tag <> "TAG" Then
        'no tag, clear all info
        With ID3v1store
            .Album = ""
            .Artist = ""
            .Title = ""
            .Year = 0
            .Comment = ""
            .Track = 0
            .Genre = 255 '255 is reserved for none
        End With
        Exit Sub
    End If
    'start reading and formatting the information
    With ID3v1store
        .Album = Trim$(IDtag.Album)
        .Artist = Trim$(IDtag.Artist)
        .Title = Trim$(IDtag.Title)
        If IsNumeric(IDtag.Year) Then
            .Year = CInt(IDtag.Year)
        Else
            .Year = 0
        End If
        .Genre = IDtag.Genre
        'check if there is track information (byte before the last byte is zero if track information exists)
        If Asc(mid$(IDtag.Comment, 29, 1)) = 0 Then
            .Comment = Trim$(Left$(IDtag.Comment, 28))
            'convert character to character code
            .Track = Asc(Right$(IDtag.Comment, 1))
        Else
            .Comment = Trim$(IDtag.Comment)
            .Track = 0
        End If
    End With
    Exit Sub
ErrorHandler:
    Close #1
    AddError "GetID3v1", Err.Number, Err.Description
End Sub
Public Sub GetTagV1(ByRef Title As String, ByRef Artist As String, ByRef Album As String, ByRef Year As Integer, ByRef Comment As String, ByRef Track As Byte, ByRef Genre As Byte)
    'return tag information
    With ID3v1store
        Title = .Title
        Artist = .Artist
        Album = .Album
        Year = .Year
        Comment = .Comment
        Track = .Track
        Genre = .Genre
    End With
End Sub
Public Function SetTagV1(ByVal Title As String, ByVal Artist As String, ByVal Album As String, ByVal Year As Integer, ByVal Comment As String, ByVal Track As Byte, ByVal Genre As Byte) As Boolean
    Dim WriteTag As ID3v1_TAG, ReadTag As ID3v1_TAG, tmpStatus As Long, tmpPosition As Long, File As String
    'error correction
    If Len(Title) > 30 Then Title = Left$(Title, 30)
    If Len(Artist) > 30 Then Artist = Left$(Artist, 30)
    If Len(Album) > 30 Then Album = Left$(Album, 30)
    If Year > Val(Format(Date, "yyyy")) Then Year = Val(Format(Date, "yyyy"))
    If Len(Comment) > 30 Then Comment = Left$(Comment, 30)
    'set up tag to be written
    WriteTag.Tag = "TAG"
    WriteTag.Title = Title
    WriteTag.Artist = Artist
    WriteTag.Album = Album
    WriteTag.Year = Format$(Year, "0000")
    WriteTag.Comment = Comment
    'if there is track information to be written...
    If Track <> 0 Then WriteTag.Comment = Left$(WriteTag.Comment, 28) & Chr$(0) & Chr$(Track)
    WriteTag.Genre = Genre
    'get some information and close file so we can write to the file
    File = GetFilename
    tmpStatus = GetStatus
    tmpPosition = Position
    CloseMP3
    'open file for processing
    'On Error GoTo ErrorHandler
    Open File For Binary As #1
        If LOF(1) < Len(ReadTag) Then Close #1: Exit Function
        'read a tag
        Get #1, LOF(1) - Len(ReadTag) + 1, ReadTag
        'check if there is a tag
        If ReadTag.Tag = "TAG" Then
            'overwrite old tag
            Put #1, LOF(1) - Len(WriteTag) + 1, WriteTag
        Else
            'go to end of file
            Seek #1, LOF(1) + 1
            'add tag
            Put #1, , WriteTag
        End If
    Close #1
    On Error GoTo 0
    'have a break
    DoEvents
    'open file and restore state to what it was before saving
    OpenMP3 File
    Select Case tmpStatus
        Case mp3classplaying
            If SeekTo(tmpPosition) Then PlayMP3
        Case mp3classpaused
            If SeekTo(tmpPosition) Then PauseMP3
        Case mp3classstopped
            If SeekTo(tmpPosition) Then StopMP3
        Case mp3classseeking
            SeekTo tmpPosition
    End Select
    'success!
    SetTagV1 = True
    Exit Function
ErrorHandler:
    Close #1
    AddError "SetTagV1", Err.Number, Err.Description
    'just in case...
    On Error Resume Next
    'restore state to what it was before saving attempt
    OpenMP3 File
    Select Case tmpStatus
        Case mp3classplaying
            If SeekTo(tmpPosition) Then PlayMP3
        Case mp3classpaused
            If SeekTo(tmpPosition) Then PauseMP3
        Case mp3classstopped
            If SeekTo(tmpPosition) Then StopMP3
        Case mp3classseeking
            SeekTo tmpPosition
    End Select
End Function


'mp3 class id3v2 handling
Private Sub GetID3v2(ByVal File As String)
    'static keeps these in memory, so no need to recreate these each time the sub is run
    Static EmptyExtended As ID3v2_EXTENDEDHEADER, EmptyHeader As ID3v2_HEADERFOOTER, EmptyInfo As ID3v2_HEADERDATA
    Dim FrameHeader As ID3v2_FRAMEHEADER, FrameData As String, CRCdata() As Byte
    Dim A As Long, B As Byte
    'On Error GoTo ErrorHandler
    'clear data
    ReDim ID3v2data(0)
    ID3v2info = EmptyInfo
    If Not IsEmpty(ID3v2extendedData) Then Erase ID3v2extendedData
    'open file for read
    Open File For Binary Access Read As #1
        'if file is smaller than 21 bytes, it can't be a valid ID3v2 tag
        '(must include header [10 bytes], frame [10 bytes] and some data in the frame [one byte minimum])
        If LOF(1) < 21 Then Close #1: Exit Sub
        'read header
        Get #1, , ID3v2header
        'check if is a ID3v2 tag; if not, quit
        If Not ID3v2header.Tag = "ID3" Then
            Close #1
            'clear data
            ID3v2extended = EmptyExtended
            ID3v2header = EmptyHeader
            Exit Sub
        End If
        With ID3v2info
            'save information for easier access
            .Version = ID3v2header.Version & "." & ID3v2header.Minor
            .Size = SynchToLong(ID3v2header.Size)
            .FlagUnsync = CBool(ID3v2header.Flag And FLAG_HEADER_UNSYNC)
            .FlagExperimental = CBool(ID3v2header.Flag And FLAG_HEADER_EXPERIMENTAL)
            .FlagExtended = CBool(ID3v2header.Flag And FLAG_HEADER_EXTENDED)
            .FlagFooter = CBool(ID3v2header.Flag And FLAG_HEADER_FOOTER)
        End With
        'check if there is an extended header
        If ID3v2info.FlagExtended Then
            'read extended header
            Get #1, , ID3v2extended
            'check for full size of the header
            A = SynchToLong(ID3v2extended.Size) - 6
            'read extended header data if any
            If A > 0 Then
                'using preserve is faster
                ReDim Preserve ID3v2extendedData(1 To A) As Byte
                'read extended header data
                Get #1, , ID3v2extendedData
            End If
            With ID3v2info
                'save information for easier access
                .FlagUpdate = CBool(ID3v2extended.Flag And FLAG_EXTENDED_UPDATE)
                .FlagCRC32 = CBool(ID3v2extended.Flag And FLAG_EXTENDED_CRC32)
                .FlagRestricted = CBool(ID3v2extended.Flag And FLAG_EXTENDED_RESTRICT)
                If .FlagCRC32 Then
                    'get CRC-32
                    .CRC32 = CLng(SynchToCurrency(Chr$(ID3v2extendedData(1)) & Chr$(ID3v2extendedData(2)) & Chr$(ID3v2extendedData(3)) & Chr$(ID3v2extendedData(4)) & Chr$(ID3v2extendedData(5))))
                    'store the current position of the file
                    A = Loc(1)
                    'calculate the length of the extended header
                    If IsEmpty(ID3v2extendedData) Then B = 6 Else B = 6 + UBound(ID3v2extendedData)
                    'prepare read buffer
                    ReDim Preserve CRCdata(ID3v2info.Size - B) As Byte
                    'read data
                    Get #1, , CRCdata
                    'get CRC-32 and compare it to the current information to see if the data is corrupt
                    ID3v2info.Corrupt = Not ID3v2info.CRC32 = GetCRC32(0, CRCdata, ID3v2info.Size - B)
                    'move back to old position in the file
                    Seek #1, A
                End If
                If .FlagRestricted Then
                    B = ID3v2extendedData(UBound(ID3v2extendedData))
                    If B And FLAG_RESTRICT_30CHARACTERS Then
                        .ResChar = id3restrict30
                    ElseIf B And FLAG_RESTRICT_1024CHARACTERS Then
                        .ResChar = id3restrict1024
                    ElseIf B And FLAG_RESTRICT_128CHARACTERS Then
                        .ResChar = id3restrict128
                    End If
                    .ResCharCode = CBool(B And FLAG_RESTRICT_ISO88591UTF8ONLY)
                    If B And FLAG_RESTRICT_64PIXELSONLY Then
                        .ResImg = id3restrict64only
                    ElseIf B And FLAG_RESTRICT_256PIXELS Then
                        .ResImg = id3restrict256
                    ElseIf B And FLAG_RESTRICT_64PIXELS Then
                        .ResImg = id3restrict64
                    End If
                    .ResPngJpeg = CBool(B And FLAG_RESTRICT_PNGJPEGONLY)
                    If B And FLAG_RESTRICT_32FRAMES4KB Then
                        .ResSize = id3restrict4KB
                    ElseIf B And FLAG_RESTRICT_64FRAMES128KB Then
                        .ResSize = id3restrict128KB
                    ElseIf B And FLAG_RESTRICT_32FRAMES40KB Then
                        .ResSize = id3restrict40KB
                    Else
                        .ResSize = id3restrict1MB
                    End If
                End If
            End With
        Else
            'clear as there is no extended header
            ID3v2extended = EmptyExtended
        End If
        Do Until Loc(1) >= ID3v2info.Size + 10
            'read frame header
            Get #1, , FrameHeader
            'check if a valid frame id; if not, exit
            If Not IsValidID(FrameHeader.ID) Then Exit Do
            'add new frame
            If IsValidID(ID3v2data(UBound(ID3v2data)).ID) Then ReDim Preserve ID3v2data(UBound(ID3v2data) + 1)
            With ID3v2data(UBound(ID3v2data))
                'start storaging of the data
                .Compressed = CBool(FrameHeader.FormatFlag And FLAG_FORMAT_COMPRESSED)
                .DLI = CBool(FrameHeader.FormatFlag And FLAG_FORMAT_DLI)
                .Encrypted = CBool(FrameHeader.FormatFlag And FLAG_FORMAT_ENCRYPT)
                .FileAlter = CBool(FrameHeader.StatusFlag And FLAG_STATUS_FILEALTER)
                .GroupMember = CBool(FrameHeader.FormatFlag And FLAG_FORMAT_GROUP)
                .ID = FrameHeader.ID
                .ReadOnly = CBool(FrameHeader.StatusFlag And FLAG_STATUS_READONLY)
                .TagAlter = CBool(FrameHeader.StatusFlag And FLAG_STATUS_TAGALTER)
                .Unsync = CBool(FrameHeader.FormatFlag And FLAG_FORMAT_UNSYNC)
                'read possible additional information before reading actual data
                If .GroupMember Then Get #1, , .GroupID
                If .DLI Then Get #1, , .Length
                'resize frame data buffer
                FrameData = String$(SynchToLong(FrameHeader.Size), Chr(0))
                'read frame data
                Get #1, , FrameData
                'store frame data
                .Data = FrameData
            End With
        Loop
    Close #1
    Exit Sub
ErrorHandler:
    Close #1
    AddError "GetID3v2", Err.Number, Err.Description
End Sub
Public Sub GetTagV2(ByVal Index As Integer, ByRef ID As String, ByRef Data As Variant, Optional ByRef TagAlter As Boolean, Optional ByRef FileAlter As Boolean, Optional ByRef ReadOnly As Boolean, Optional ByRef GroupMember As Boolean, Optional ByRef GroupID As Byte, Optional ByRef Compressed As Boolean, Optional ByRef Encrypted As Boolean, Optional ByRef Unsync As Boolean, Optional ByRef DLI As Boolean, Optional ByRef Length As Long)
    'check we didn't get an index too small or big
    If Index < 1 Or Index > UBound(ID3v2data) + 1 Then Exit Sub
    'return the information
    With ID3v2data(Index - 1)
        ID = .ID
        Data = .Data
        If IsMissing(TagAlter) Then Exit Sub
        TagAlter = .TagAlter
        If IsMissing(FileAlter) Then Exit Sub
        FileAlter = .FileAlter
        If IsMissing(ReadOnly) Then Exit Sub
        ReadOnly = .ReadOnly
        If IsMissing(GroupMember) Then Exit Sub
        GroupMember = .GroupMember
        If IsMissing(GroupID) Then Exit Sub
        GroupID = .GroupID
        If IsMissing(Compressed) Then Exit Sub
        Compressed = .Compressed
        If IsMissing(Encrypted) Then Exit Sub
        Encrypted = .Encrypted
        If IsMissing(Unsync) Then Exit Sub
        Unsync = .Unsync
        If IsMissing(DLI) Then Exit Sub
        DLI = .DLI
        If IsMissing(Length) Then Exit Sub
        Length = .Length
    End With
End Sub
Public Sub GetTagV2Extended(ByRef Update As Boolean, ByRef CRC32 As Boolean, ByRef Restricted As Boolean, ByRef Corrupt As Boolean, ByRef ResSize As ID3V2_RESTRICT_SIZE, ByRef ResCharCode As Boolean, ByRef ResChar As ID3V2_RESTRICT_CHARS, ByRef ResPngJpeg As Boolean, ByRef ResImg As ID3V2_RESTRICT_IMAGE)
    'return tag extended header information
    With ID3v2info
        Update = .FlagUpdate
        CRC32 = .FlagCRC32
        Restricted = .FlagRestricted
        ResSize = .ResSize
        ResCharCode = .ResCharCode
        ResChar = .ResChar
        ResPngJpeg = .ResPngJpeg
        ResImg = .ResImg
        Corrupt = .Corrupt
    End With
End Sub
Public Sub GetTagV2Header(ByRef Version As String, ByRef Size As Long, ByRef Extended As Boolean, Optional ByRef Experimental As Boolean, Optional ByRef Unsync As Boolean)
    'return tag header information
    With ID3v2info
        Version = .Version
        Size = .Size
        Extended = .FlagExtended
        If IsMissing(Experimental) Then Exit Sub
        Experimental = .FlagExperimental
        If IsMissing(Unsync) Then Exit Sub
        Unsync = .FlagUnsync
    End With
End Sub
Public Function TagsV2() As Integer
    'return the number of tags
    If UBound(ID3v2data) = 0 Then If ID3v2data(0).Data = "" Then TagsV2 = 0: Exit Function
    'there is information
    TagsV2 = UBound(ID3v2data) + 1
End Function


'mp3 class id3v2 help functions
Private Function Bin2Dec(strBinary As String) As Currency
    Dim I As Integer
    'convert binary string to value
    For I = 0 To Len(strBinary) - 1
        If mid$(strBinary, Len(strBinary) - I, 1) = "1" Then Bin2Dec = Bin2Dec + (2 ^ I)
    Next
End Function
Private Function Dec2Bin(Value As Long, Optional Numbers As Byte = 0) As String
    Dim A As Currency, Temp As String
    'convert value to binary string
    A = 1
    Do
        Temp = CStr(Abs((Value And A) > 0)) & Temp
        A = A * 2
    Loop While A <= Value
    If Numbers Then Dec2Bin = Format$(Temp, String$(Numbers, "0")) Else Dec2Bin = Temp
End Function
Private Function GetCRC32(ByVal PreviousCRC As Long, ByRef Buffer() As Byte, ByVal BufferSize As Double) As Long
    Dim A As Double, B As Long
    'get CRC-32 of a data
    'check if array is empty
    If UBound(Buffer) < LBound(Buffer) Then
        ' Files with no data have a CRC of 0
        GetCRC32 = 0
    Else
        B = PreviousCRC Xor &HFFFFFFFF
        'go through the whole buffer
        For A = 0 To BufferSize
            'calculate CRC-32
            B = (((B And &HFFFFFF00) \ &H100) And &HFFFFFF) Xor (CRC32CheckSumTable((B And &HFF) Xor Buffer(A)))
        Next A
        'return CRC-32
        GetCRC32 = B Xor &HFFFFFFFF
    End If
End Function
Private Function IsValidID(ByVal ID As String) As Boolean
    'check for valid tag id
    IsValidID = ID Like "[A-Z0-9][A-Z0-9][A-Z0-9][A-Z0-9]"
End Function
Private Function SynchToCurrency(ByVal Value As String) As Currency
    Dim Temp As String, A As Integer
    'convert synchronized value to Currency (64-bit)
    'the first bit of each byte is used as a check bit (always zero)
    'convert input string to binary
    For A = 1 To Len(Value)
        'convert to binary and drop first bit
        Temp = Temp & mid$(Dec2Bin(Asc(mid$(Value, A, 1)), 8), 2)
    Next A
    SynchToCurrency = Bin2Dec(Temp)
End Function
Private Function SynchToLong(ByVal Value As Long) As Long
    Dim Temp As String
    'convert synchronized 28-bit value to Long (32-bit)
    'the first bit of each byte is used as a check bit (always zero)
    'convert to binary
    Temp = Dec2Bin(Value, 32)
    'invert and drop first bits
    SynchToLong = CLng(Bin2Dec(mid$(Temp, 26, 7) & mid$(Temp, 18, 7) & mid$(Temp, 10, 7) & mid$(Temp, 2, 7)))
End Function


'mp3 class additional functions
Public Function ConvertText(ByVal Text As String) As String
    Dim Temp As String, A As Integer
    On Error GoTo ErrorHandler
    If Len(Text) < 3 Then ConvertText = Text: Exit Function
    'check if a certain coding is used
    Select Case Asc(Left$(Text, 1))
        Case 0, 3 'ISO-8859-1 and UTF-8
            Temp = mid$(Text, 2)
        Case 1, 2 'UTF-16 and UTF-16BE
            'check for BOM (shouldn't be included with UTF-16BE)
            If mid$(Text, 2, 2) = Chr$(255) & Chr$(254) Or mid$(Text, 2, 2) = Chr$(254) & Chr$(255) Then
                Text = mid$(Text, 4)
            Else
                Text = mid$(Text, 2)
            End If
            'convert to UTF-16
            For A = 1 To Len(Text) Step 2
                Temp = Temp & ChrW(Asc(mid$(Text, A, 1)) + Asc(mid$(Text, A + 1, 1)) * 256)
            Next A
        Case Else
            Temp = Text
    End Select
    'return converted text
    ConvertText = Temp
    Exit Function
ErrorHandler:
    AddError "ConvertText", Err.Number, Err.Description
    'return what we managed to get
    ConvertText = Temp
End Function

