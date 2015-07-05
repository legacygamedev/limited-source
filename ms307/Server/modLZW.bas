Attribute VB_Name = "modLZW"
' Special thanks to Chris Dodge for repo
'     rting the bug
Option Explicit


Private Type BNode
    DictIdx As Long
    pLeft As Long
    pRight As Long
    End Type
    Dim Dict(4096) As String
    Dim NextDictIdx As Long
    Dim Heap(4096) As BNode
    Dim NextHeapIdx As Long
    Dim pStr As Long


Sub InitDict()
    Dim i As Integer
    


    For i = 0 To 255
        Dict(i) = Chr(i)
    Next i
    ' Not really necessary
    '
    ' For i = 256 To 4095
    ' Dict(i) = ""
    ' Next i
    
NextDictIdx = 256
NextHeapIdx = 0
End Sub


Function AddToDict(s As String) As Long


    If NextDictIdx > 4095 Then
    NextDictIdx = 256
NextHeapIdx = 0
End If



If Len(s) = 1 Then
AddToDict = Asc(s)
Else
AddToDict = AddToBTree(0, s)
End If
End Function


Function AddToBTree(ByRef Node As Long, ByRef s As String) As Long
    Dim i As Integer
    


    If Node = -1 Or NextHeapIdx = 0 Then
        Dict(NextDictIdx) = s
        Heap(NextHeapIdx).DictIdx = NextDictIdx
    NextDictIdx = NextDictIdx + 1
    Heap(NextHeapIdx).pLeft = -1
    Heap(NextHeapIdx).pRight = -1
    Node = NextHeapIdx
NextHeapIdx = NextHeapIdx + 1
AddToBTree = -1
Else
i = StrComp(s, Dict(Heap(Node).DictIdx))


If i < 0 Then
    AddToBTree = AddToBTree(Heap(Node).pLeft, s)
ElseIf i > 0 Then
    AddToBTree = AddToBTree(Heap(Node).pRight, s)
Else
    AddToBTree = Heap(Node).DictIdx
End If
End If
End Function


Private Sub WriteStrBuf(s As String, s2 As String)


    Do While pStr + Len(s2) - 1 > Len(s)
        s = s & Space(100000)
    Loop
    Mid$(s, pStr) = s2
    pStr = pStr + Len(s2)
End Sub


Function Compress(IPStr As String) As String
    Dim TmpStr As String
    Dim Ch As String
    Dim DictIdx As Integer
    Dim LastDictIdx As Integer
    Dim FirstInPair As Boolean
    Dim HalfCh As Integer
    Dim i As Long
    Dim ostr As String
    
    InitDict
    FirstInPair = True
    pStr = 1
    


    For i = 1 To Len(IPStr)
        Ch = Mid$(IPStr, i, 1)
        
        DictIdx = AddToDict(TmpStr & Ch)


        If DictIdx = -1 Then


            If FirstInPair Then
                HalfCh = (LastDictIdx And 15) * 16
            Else
                WriteStrBuf ostr, Chr(HalfCh Or (LastDictIdx And 15))
            End If
            WriteStrBuf ostr, Chr(LastDictIdx \ 16)
            FirstInPair = Not FirstInPair
            TmpStr = Ch
            LastDictIdx = Asc(Ch)
        Else
            TmpStr = TmpStr & Ch
            LastDictIdx = DictIdx
        End If
    Next i
    
    WriteStrBuf ostr, _
    IIf(FirstInPair, Chr(LastDictIdx \ 16) & Chr((LastDictIdx And 15) * 16), _
    Chr(HalfCh Or (LastDictIdx And 15)) & Chr(LastDictIdx \ 16))
    
    Compress = Left(ostr, pStr - 1)
    
End Function


Function GC(str As String, position As Long) As Integer
    GC = Asc(Mid$(str, position, 1))
End Function


Function DeCompress(IPStr As String) As String
    Dim DictIdx As Integer
    Dim FirstInPair As Boolean
    Dim i As Long
    Dim s As String
    Dim s2 As String
    InitDict
    pStr = 1
    i = 1
    FirstInPair = True
    


    Do While i < Len(IPStr)


        If FirstInPair Then
            DictIdx = (GC(IPStr, i) * 16) Or (GC(IPStr, i + 1) \ 16)
            i = i + 1
        Else
            DictIdx = (GC(IPStr, i + 1) * 16) Or (GC(IPStr, i) And 15)
            i = i + 2
        End If
        FirstInPair = Not FirstInPair
        


        If i > 2 Then


            If DictIdx = NextDictIdx Or (DictIdx = 256 And NextDictIdx = 4096) Then
                AddToDict s2 & Left$(s2, 1)
            Else
                AddToDict s2 & Left$(Dict(DictIdx), 1)
            End If
        End If
        s2 = Dict(DictIdx)
        WriteStrBuf s, s2
    Loop
    
    DeCompress = Left(s, pStr - 1)
End Function



