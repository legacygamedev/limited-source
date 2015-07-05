Attribute VB_Name = "modCompression"
Private Type Pattern
    Text As String
    TimesRepeated As Integer
    Position As Long
End Type

Public Function Compress(Text As String, Optional ByVal MaxPatternLen As Byte = 5)
Dim Patterns() As Pattern
Dim PatternLen As Long
Dim Char As String
Dim Compressed As Integer
Dim ShortestPattern As Byte

    If MaxPatternLen > Len(Text) Then MaxPatternLen = Len(Text) 'this can save alot of time
    ShortestPattern = 4 + Len(STR(Len(Text)))
    If ShortestPattern > MaxPatternLen Then ShortestPattern = MaxPatternLen

    ReDim Patterns(1 To 1)
    Do Until Text = ""
PatternLoop:
        If Text = "" Then Exit Do 'Sometimes control is directed here when it shouldn't be.
        For CurPatternLen = 1 To MaxPatternLen
            If MaxPatternLen > Len(Text) Then MaxPatternLen = Len(Text) 'this can save alot of time
            Char = Left(Text, CurPatternLen)
            If Left(Text, CurPatternLen * 2) = Char & Char Then
                PatternLen = CurPatternLen
                Do Until Right(Left(Text, PatternLen + CurPatternLen), CurPatternLen) <> Char Or PatternLen = Len(Text)
                    PatternLen = PatternLen + CurPatternLen
                Loop
                   
                If PatternLen > ShortestPattern And PatternLen > 6 Then
                    ReDim Preserve Patterns(1 To UBound(Patterns) + 1)
                    Patterns(UBound(Patterns)).Text = Char
                    Patterns(UBound(Patterns)).TimesRepeated = PatternLen / CurPatternLen
                    Patterns(UBound(Patterns)).Position = Len(Compress)
                   
                    Text = Right(Text, Len(Text) - PatternLen)
                Else
                    Compress = Compress & Left(Text, PatternLen)
                    Text = Right(Text, Len(Text) - PatternLen)
                End If
                GoTo PatternLoop
            End If
        Next
        Compress = Compress & Left(Text, 1)
        Text = Right(Text, Len(Text) - 1)
    Loop
   
    For X = 1 To UBound(Patterns)
        If Patterns(X).Text <> "" Then
            Compress = Patterns(X).Text & Compress
            Compress = Patterns(X).Position & " " & Compress
            Compress = Patterns(X).TimesRepeated & " " & Compress
            Compress = Len(Patterns(X).Text & STR(Patterns(X).TimesRepeated & Patterns(X).Position)) + 2 & " " & Compress
           
            Compressed = Compressed + 1
        End If
    Next
    Compress = Compressed & " " & Compress
       
End Function
Public Function Decompress(Text As String)
Dim Patterns() As Pattern
Dim Xstr As String
    If Left(Text, InStr(Text, " ") - 1) = 0 Then
        Decompress = Right(Text, Len(Text) - 2)
        Exit Function
    End If
    ReDim Patterns(1 To Left(Text, InStr(Text, " ") - 1))
    Text = Right(Text, Len(Text) - InStr(Text, " "))
   
    For X = 1 To UBound(Patterns)
        Xstr = Left(Text, InStr(Text, " ") - 1)
        Text = Right(Text, Len(Text) - Len(Xstr))
        Xstr = Left(Text, Xstr)
        Text = Right(Text, Len(Text) - Len(Xstr))
        Xstr = Right(Xstr, Len(Xstr) - 1)
       
        Patterns(X).TimesRepeated = Left(Xstr, InStr(Xstr, " ") - 1)
        Xstr = Right(Xstr, Len(Xstr) - InStr(Xstr, " "))
        Patterns(X).Position = Left(Xstr, InStr(Xstr, " "))
        Xstr = Right(Xstr, Len(Xstr) - InStr(Xstr, " "))
        Patterns(X).Text = Xstr
        Xstr = ""
    Next
   
    'Instrt Patterns into text
    For X = 1 To UBound(Patterns)
        Xstr = ""
        For Y = 1 To Patterns(X).TimesRepeated
            Xstr = Xstr & Patterns(X).Text
        Next
        Text = Left(Text, Patterns(X).Position) & Xstr & Right(Text, Len(Text) - Patterns(X).Position)
    Next

    Decompress = Text
End Function
