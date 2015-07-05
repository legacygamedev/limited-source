Attribute VB_Name = "modIO"
Option Explicit

Public Sub Place(ByVal Path As String, ByVal Key As String, ByVal Text As String, Optional Instance As Long = 1, Optional ReWrite As Boolean = False)
Dim f As Integer, n As Long, closerfound As Boolean, FilePath As String, TF As Long

FilePath = Path

Dim LineText() As String

f = FreeFile
closerfound = False
n = 1
TF = 0

Open FilePath For Input As #f
    Do Until closerfound
        ReDim Preserve LineText(1 To n)
        Line Input #f, LineText(n)
        If Mid(LCase$(Trim$(LineText(n))), 1, Len(Key)) = LCase$(Key) Then
            ' Adds one to the total found count
            TF = TF + 1
            
            ' Checks if the total found meets the instance wanted
            If TF = Instance Then
                If ReWrite Then
                    LineText(n) = Text
                Else
                    LineText(n) = Mid(LineText(n), 1, Len(Key)) & Text
                End If
            End If
        End If
        
        ' Check for closing LineText, if so,
        ' make closerfound = True to stop LineText input
        If LCase$(Trim$(LineText(n))) = "[endfile]" Then closerfound = True
        ' If the end of file is reached, close
        If EOF(f) = True Then closerfound = True
        n = n + 1
    Loop
Close #f

closerfound = False
n = 1

Open FilePath For Output As #f
    ' Error check
    If LineText(1) = "[endfile]" Then Exit Sub
    
    Do Until closerfound
        Print #f, LineText(n)
        
        ' Check for closing LineText, if so,
        ' make closerfound = True to stop LineText input
        If LCase$(Trim$(LineText(n))) = "[endfile]" Then closerfound = True
        ' If the end of file is reached, close
        If EOF(f) = True Then closerfound = True
        n = n + 1
    Loop
Close #f
End Sub

Public Function Read(ByVal Path As String, ByVal Key As String, Optional Instance As Long = 1) As String
Dim f As Integer, n As Long, closerfound As Boolean, FilePath As String, TF As Long

FilePath = Path
Dim LineText() As String

f = FreeFile
closerfound = False
n = 1
TF = 0

' presets function return
Read = Chr(215)

Open FilePath For Input As #f
    Do Until closerfound
        ReDim Preserve LineText(1 To n)
        Line Input #f, LineText(n)
        If Mid(LCase$(Trim$(LineText(n))), 1, Len(Key)) = LCase$(Key) Then
            ' Adds one to the total found count
            TF = TF + 1
            
            ' Checks if the total found meets the instance wanted
            If TF = Instance Then
                Read = Mid(Trim$(LineText(n)), Len(Key) + 1)
                Exit Function
            End If
        End If
        
        ' Check for closing LineText, if so,
        ' make closerfound = True to stop LineText input
        If LCase$(Trim$(LineText(n))) = "[endfile]" Then closerfound = True
        ' If the end of file is reached, close
        If EOF(f) = True Then closerfound = True
        n = n + 1
    Loop
Close #f
End Function
