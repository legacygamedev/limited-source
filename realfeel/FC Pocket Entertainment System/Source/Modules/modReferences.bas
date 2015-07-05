Attribute VB_Name = "modReferences"
Option Explicit

' load the inet class of mine
Public Inet As New clsInet
Public SavedText As String
Public BeginJump As Boolean
Public RestartAnimation As Boolean

Private Const MAX_LINES = 100

Public Function Exists(ByVal Path As String)
On Error Resume Next
Exists = False
If Dir(Path) <> "" Then Exists = True
End Function

Public Sub Place(ByVal Path As String, ByVal Key As String, ByVal Text As String, Optional Instance As Long = 1, Optional ReWrite As Boolean = False)
Dim f As Integer, n As Long, closerfound As Boolean, FilePath As String, TF As Long

FilePath = Path

Dim LineText(1 To MAX_LINES) As String

f = FreeFile
closerfound = False
n = 1
TF = 0

Open FilePath For Input As #f
    Do Until closerfound
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
        n = n + 1
    Loop
Close #f
End Sub

Public Function Read(ByVal Path As String, ByVal Key As String, Optional Instance As Long = 1) As String
Dim f As Integer, n As Long, closerfound As Boolean, FilePath As String, TF As Long

FilePath = Path
Dim LineText(1 To MAX_LINES) As String

f = FreeFile
closerfound = False
n = 1
TF = 0

' presets function return
Read = Chr(215)

Open FilePath For Input As #f
    Do Until closerfound
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
        n = n + 1
    Loop
Close #f
End Function

Public Sub ColorKeyword(ByVal src As RichTextBox, ByVal dst As RichTextBox, ByVal Word As String, ByVal Color As Long, Optional Bold As Boolean = False, Optional Italic As Boolean = False, Optional WholeLine As Boolean = False)
On Error GoTo Quit
Dim Start As Long, Place As Long, Length As Long, EndLine As Boolean

'Set start to 0
Start = 0

'Set EndLine to false
EndLine = False

'Continuously loop to find all instances of the word
Do While True
Start = src.Find(Word, Start, , rtfMatchCase)
If Start <> -1 Then
    If WholeLine = False Then
        src.SelStart = Start
        src.SelLength = Len(Word)
        src.SelColor = Color
        If Bold = True Then src.SelBold = True
        If Italic = True Then src.SelItalic = True
        src.SelText = Word
        Start = Start + Len(Word)
    Else
        Length = 1
        src.SelStart = Start
        EndLine = False
        If Mid(src.Text, Start + (Length - 1), 1) <> vbNewLine Then
            Do While Not EndLine
                Length = Length + 1
                If Mid(src.Text, Start + (Length - 1), 1) = vbNewLine Then EndLine = True
                'Prevents continuous looping
                If Len(src.Text) < Length Then EndLine = True
            Loop
        End If
        src.SelLength = Length
        src.SelColor = Color
        If Bold = True Then src.SelBold = True
        If Italic = True Then src.SelItalic = True
        src.SelText = Mid(src.Text, Start, Length)
        Start = Start + Length
    End If
Else
    Exit Do
End If
Loop

Quit:
'Set place to old position of cursor
Place = dst.SelStart

'Update cursor position back to original
dst.SelStart = Place
End Sub

Public Sub ColorBBCode(ByVal src As RichTextBox, ByVal dst As RichTextBox, ByVal Initial As String, ByVal Closer As String, ByVal Color As Long, Optional Bold As Boolean = False, Optional Italic As Boolean = False, Optional Underline As Boolean = False, Optional pLeft As Boolean = False, Optional pCenter As Boolean = False, Optional pRight As Boolean = False, Optional DelIdentifiers As Boolean = False)
On Error GoTo Quit
Dim Start As Long, Place As Long, Length As Long, EndLine As Boolean, Change As Boolean, Txt As String

'Set start to 0
Start = 0
src.SelStart = 0

'Set EndLine to false
EndLine = False

'This prevents an odd error
'src.Text = " " & src.Text & " "

'Continuously loop to find all instances of the word
Do While True
Start = src.Find(Initial, Start, , rtfMatchCase)
Debug.Print "Start: " & Start
If Start <> -1 Then
    Length = 0
    src.SelStart = Start
    EndLine = False
    Change = True
    Do While Not EndLine
        Length = Length + 1
        If Right(Mid(src.Text, Start + 1, Length), Len(Closer)) = Closer Then EndLine = True
        
        'Prevents continuous looping, doesn't change anything
        If Len(src.Text) < Length Then
            EndLine = True
            Change = False
        End If
    Loop
    'Checks for the change boolean in case there is no closing tag
    If Change = True Then
        Length = Len(Mid(src.Text, Start + 1, Length))
        Txt = Mid(src.Text, Start + Len(Initial) + 1, Length - Len(Initial))
        Debug.Print Txt
        Txt = Mid(Txt, 1, Len(Txt) - Len(Closer))
        Debug.Print Txt
        src.SelLength = Len(Initial) + Len(Txt) + Len(Closer)
        src.SelColor = Color
        If Bold = True Then src.SelBold = True
        If Italic = True Then src.SelItalic = True
        If Underline = True Then src.SelUnderline = True
        If pLeft = True Then src.SelAlignment = rtfLeft
        If pCenter = True Then src.SelAlignment = rtfCenter
        If pRight = True Then src.SelAlignment = rtfRight
        If DelIdentifiers = True Then src.SelText = Txt
    End If
    Start = Start + Length
Else
    Exit Do
End If
Loop

'This prevents an odd error
'src.Text = Left(src.Text, Len(src.Text) - 1)
'src.Text = Right(src.Text, Len(src.Text) - 1)

Quit:
'Set place to old position of cursor
Place = dst.SelStart

'Update cursor position back to original
dst.SelStart = Place
End Sub

Public Sub FilterText(ByVal frm As Form, ByVal Text As String)
frm.rtbCopy.Text = Text
' Find and make bold words
Call ColorBBCode(frm.rtbCopy, frm.rtbNews, "[b]", "[/b]", RGB(0, 0, 0), True, False, False, False, False, False, True)
' Find and make italic words
Call ColorBBCode(frm.rtbCopy, frm.rtbNews, "[i]", "[/i]", RGB(0, 0, 0), False, True, False, False, False, False, True)
' Find and make underlined words
Call ColorBBCode(frm.rtbCopy, frm.rtbNews, "[u]", "[/u]", RGB(0, 0, 0), False, False, True, False, False, False, True)
' Find and adjust alignment
Call ColorBBCode(frm.rtbCopy, frm.rtbNews, "[left]", "[/left]", RGB(0, 0, 0), False, False, False, True, False, False, True)
Call ColorBBCode(frm.rtbCopy, frm.rtbNews, "[center]", "[/center]", RGB(0, 0, 0), False, False, False, False, True, False, True)
Call ColorBBCode(frm.rtbCopy, frm.rtbNews, "[right]", "[/right]", RGB(0, 0, 0), False, False, False, False, False, True, True)
' Find and set the correct color
Call ColorBBCode(frm.rtbCopy, frm.rtbNews, "[color=red]", "[/color]", RGB(255, 0, 0), False, False, False, False, False, False, True)
Call ColorBBCode(frm.rtbCopy, frm.rtbNews, "[color=blue]", "[/color]", RGB(0, 0, 255), False, False, False, False, False, False, True)
Call ColorBBCode(frm.rtbCopy, frm.rtbNews, "[color=green]", "[/color]", RGB(0, 255, 0), False, False, False, False, False, False, True)
Call ColorBBCode(frm.rtbCopy, frm.rtbNews, "[color=yellow]", "[/color]", RGB(255, 255, 0), False, False, False, False, False, False, True)
Call ColorBBCode(frm.rtbCopy, frm.rtbNews, "[color=cyan]", "[/color]", RGB(0, 255, 255), False, False, False, False, False, False, True)
Call ColorBBCode(frm.rtbCopy, frm.rtbNews, "[color=pink]", "[/color]", RGB(255, 0, 255), False, False, False, False, False, False, True)
frm.rtbNews = frm.rtbCopy
End Sub
