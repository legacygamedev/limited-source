VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RLE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function Compress(ByVal wStr As String) As String

    Dim iLoop As Long
    Dim nLoop As Long
    Dim OutPt As String
    Dim lPart As String
    Dim nPart As String
    
    For iLoop = 1 To Len(wStr)
        lPart = Mid(wStr, iLoop, 1)
        
        Do
            nPart = Mid(wStr, iLoop + nLoop, 1)
            nLoop = nLoop + 1
        Loop While nPart = lPart And _
          nLoop <= 255
        nLoop = nLoop - 1
        
        If nLoop >= 3 Then
            OutPt = OutPt & Chr(255) & Chr(nLoop) & lPart
        Else
            OutPt = OutPt & Mid(wStr, iLoop, nLoop)
        End If
        
        iLoop = iLoop + (nLoop - 1)
        nLoop = 0
        
    Next iLoop

    Compress = OutPt

End Function

Public Function UnCompress(ByVal wStr As String) As String

    Dim iLoop As Long
    
    Dim OutPt As String
    Dim lPart As String

    For iLoop = 1 To Len(wStr)
        lPart = Mid(wStr, iLoop, 1)
        
        If lPart = Chr(255) Then
            OutPt = OutPt & String(Asc(Mid(wStr, iLoop + 1, 1)), _
              Mid(wStr, iLoop + 2, 1))
            iLoop = iLoop + 2
        Else
            OutPt = OutPt & lPart
        End If
    Next iLoop
    
    UnCompress = OutPt

End Function


