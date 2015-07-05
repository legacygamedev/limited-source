VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.UserControl rtbSyntax 
   ClientHeight    =   2895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4095
   ScaleHeight     =   2895
   ScaleWidth      =   4095
   ToolboxBitmap   =   "rtbSyntax.ctx":0000
   Begin RichTextLib.RichTextBox rtb 
      Height          =   735
      Left            =   1320
      TabIndex        =   0
      Top             =   1080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1296
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      RightMargin     =   1.00000e5
      TextRTF         =   $"rtbSyntax.ctx":0532
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "rtbSyntax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' rtbSyntax Control.
' Aaron Bennear
' April 7, 2002
'
' Adds syntax highlighting to the RichTextBox control. A design goal was to
' accomplish this with only one file, the ctl file. (There is an associated
' ctx file, but this is not required and simply supplies the toolbox icon.)
' This makes the control a drop in replacement for the standard RichTextBox,
' and very easy to add to a project. Therefore, the list of keywords
' highlighted is composed of constants rather than loaded from a file. Also,
' the syntax parsing is not object oriented and does not use classes.
'
' The list of keywords to highlight can be modified. However, the syntax
' highlighting is designed to handle VBScript code. It therefore does not
' handle non-VB conventions such as multiline comments.
'
' As much as possible, the control uses delegation to expose the properties
' and events of the underlying RichTextBox control. The Data and OLE related
' properties proved problematic and are not supported. Also, properties that
' are read-only at run time can not be delegated from user control to
' constituent control. For RichTextBox these are: Appearance, DisableNoScroll,
' MultiLine, and Scrollbars. Of course, these property limitations can dealt
' with by editing the control.
'
' The parser scans line by line. Within each line the text is read left to
' right and the type of syntax context is kept track of. As the context
' changes, from keyword to string for instance, coloring is done for the
' subsection just completed.
'
' As the text in the RichTextBox changes, the parsing needs to be redone. To
' improve performance, an attempt is made to only reparse the lines that are
' changee. This is done by keeping track of the current and previous point of
' insertion. Also, an API call is used to disable the repainting of the
' RichTextBox as it is being colored, to prevent unsightly selection changes
' from flashing by.
'
' The tedious task of matching positions in the text being parsed with its
' absolute position in the RichTextBox is made more so by the fact that VB
' string functions are 1-indexed while the RichTextBox is 0-indexed.
'
' The control exposed only one additionaly method: HighlightRefresh. This
' reparses all the text. It should not be necessary to call this function
' from code using the control, but is provided just in case.
'

Option Explicit

' assumes one character long comment
Const COMMENT = "'"

Const DELIMITER = vbTab & " ,(){}[]-+*%/='~!&|\<>?:;."

' Space surrounding each word is significant. It allows searching on whole
' words. Note that these constant declares are long and could reach the line
' length limit of 1023 characters. If so, simply split to 2 constants and
' combine into a third constant with the appropriate name.
Const RESERVED As String = " And Call Case Const Dim Do Each Else ElseIf Empty End Eqv Erase Error Exit Explicit False For Function If Imp In Is Loop Mod Next Not Nothing Null On Option Or Private Public Randomize ReDim Resume Select Set Step Sub Then To True Until Wend While Xor "
Const FUNC_OBJ As String = " Anchor Array Asc Atn CBool CByte CCur CDate CDbl Chr CInt CLng Cos CreateObject CSng CStr Date DateAdd DateDiff DatePart DateSerial DateValue Day Dictionary Document Element Err Exp FileSystemObject  Filter Fix Int Form FormatCurrency FormatDateTime FormatNumber FormatPercent GetObject Hex History Hour InputBox InStr InstrRev IsArray IsDate IsEmpty IsNull IsNumeric IsObject Join LBound LCase Left Len Link LoadPicture Location Log LTrim RTrim Trim Mid Minute Month MonthName MsgBox Navigator Now Oct Replace Right Rnd Round ScriptEngine ScriptEngineBuildVersion ScriptEngineMajorVersion ScriptEngineMinorVersion Second Sgn Sin Space Split Sqr StrComp String StrReverse Tan Time TextStream TimeSerial TimeValue TypeName UBound UCase VarType Weekday WeekDayName Window Year "
Const KEYWORD_PAD As String = " "

' These are Red, Green, Blue sets of values used to color text. The InitParser
' function converts them using the RBG function to values the RichTextBox can
' use.
Const RGB_COMMENT As String = "0,128,0"
Const RGB_STRING As String = "255,0,255"
Const RGB_RESERVED As String = "0,0,255"
Const RGB_FUNC_OBJ As String = "255,0,0"
Const RGB_DELIMITER As String = "0,0,0"
Const RGB_NORMAL As String = "0,0,0"

Enum SyntaxTypes
    ColorComment = 0
    ColorString = 1
    ColorReserved = 2
    ColorFuncObj = 3
    ColorDelimiter = 4
    ColorNormal = 5
End Enum

' Global variable used to suppress parsing until the end of a series of
' changes. Or, in the Change event itself to prevent cascaded Change events.
Private mbInChange As Boolean

' RGB values derived from constants
Private mrgbComment As Long
Private mrgbString As Long
Private mrgbReserved As Long
Private mrgbFuncObj As Long
Private mrgbDelimiter As Long
Private mrgbNormal As Long

' One WinAPI call. Used to suppress repainting during parsing.
Private Const WM_SETREDRAW = &HB
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

' Keeping track of current and previous insertion point. Used to determine
' what portion of text has changed.
Private mlPrevSelStart As Long
Private mlCurSelStart As Long

'
' Delegation code
'

'Default Property Values:
Const m_def_ForeColor = 0
Const m_def_hWnd = 0
'Property Variables:
Dim m_ForeColor As Long
Dim m_hWnd As Long
'Event Declarations:
Event Change() 'MappingInfo=rtb,rtb,-1,Change
Attribute Change.VB_Description = "Indicates that the contents of a control have changed."
Event Click() 'MappingInfo=rtb,rtb,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=rtb,rtb,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=rtb,rtb,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=rtb,rtb,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=rtb,rtb,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=rtb,rtb,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses a mouse button."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=rtb,rtb,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=rtb,rtb,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user presses and releases a mouse button."
Event SelChange() 'MappingInfo=rtb,rtb,-1,SelChange
Attribute SelChange.VB_Description = "Occurs when the current selection of text in the RichTextBox control has changed or the insertion point has moved."

'
' Sub UserControl_Initialize
' Position constituate control, call initialization.
'
Private Sub UserControl_Initialize()
    rtb.Top = 0
    rtb.Left = 0
    
    InitParser
    mlPrevSelStart = 0
End Sub

'
' Sub InitParser
' Derive color values.
'
Private Sub InitParser()
    Dim vArr
    
    vArr = Split(RGB_COMMENT, ",")
    mrgbComment = RGB(vArr(0), vArr(1), vArr(2))
    
    vArr = Split(RGB_STRING, ",")
    mrgbString = RGB(vArr(0), vArr(1), vArr(2))
    
    vArr = Split(RGB_RESERVED, ",")
    mrgbReserved = RGB(vArr(0), vArr(1), vArr(2))
    
    vArr = Split(RGB_FUNC_OBJ, ",")
    mrgbFuncObj = RGB(vArr(0), vArr(1), vArr(2))
    
    vArr = Split(RGB_DELIMITER, ",")
    mrgbDelimiter = RGB(vArr(0), vArr(1), vArr(2))
    
    vArr = Split(RGB_NORMAL, ",")
    mrgbNormal = RGB(vArr(0), vArr(1), vArr(2))
    
End Sub

'
' Sub rtb_Change
' Determine the changed region and feed to the parser.
'
Private Sub rtb_Change()
    RaiseEvent Change
    
    If mbInChange = True Then
        ' change is being blocked or deferred
        GoTo ExitSub
    End If
    
    ' suppress change events generated during this change event
    '
    mbInChange = "True"
    
        
    Dim srtbText As String      ' working string
    ' add final cariage return so last line is processed
    srtbText = rtb.Text & vbCrLf
    
    ' preserve selection and restore at end
    '
    Dim lOrigSelStart As Long
    Dim lOrigSelLength As Long
    lOrigSelStart = rtb.SelStart
    lOrigSelLength = rtb.SelLength
    
    
    Dim lStartPos As Long
    Dim lEndPos As Long
    
    If mlPrevSelStart < rtb.SelStart Then
        lStartPos = mlPrevSelStart
        lEndPos = rtb.SelStart
    Else
        lStartPos = rtb.SelStart
        lEndPos = mlPrevSelStart
    End If
    
    
    If lStartPos > 1 Then
        ' set start position to beginning of line
        If InStrRev(srtbText, vbCrLf, lStartPos - 1) > 0 Then
            lStartPos = InStrRev(srtbText, vbCrLf, lStartPos - 1) + Len(vbCrLf) - 1
        Else
            lStartPos = 0
        End If
    Else
        lStartPos = 0
    End If
    
    ' set end position to end of line
    If InStr(lEndPos + 1, srtbText, vbCrLf) > 0 Then
        lEndPos = InStr(rtb.SelStart + 1, srtbText, vbCrLf) - 1
    Else
        lEndPos = Len(srtbText) - 1
    End If
    
    
    Dim x As Long
    
    'prevent textbox from repainting
    x = SendMessage(rtb.hWnd, WM_SETREDRAW, 0, 0)
    
    ' send affected text to the parser, along with its position in the
    ' RichTextBox
    If lStartPos <> lEndPos Then
        ParseLines Mid(srtbText, lStartPos + 1, lEndPos - lStartPos), rtb, lStartPos
    End If
    
    rtb.SelStart = lOrigSelStart
    rtb.SelLength = lOrigSelLength

    'allow texbox to repaint
    x = SendMessage(rtb.hWnd, WM_SETREDRAW, 1, 0)
    'force repaint
    rtb.Refresh
    
    mbInChange = False
    
ExitSub:
    
End Sub

'
' Sub rtb_SelChange
' Keep track of previous SelStart to allow determination of
' affected region.
'
Private Sub rtb_SelChange()
    RaiseEvent SelChange

    mlPrevSelStart = mlCurSelStart
    mlCurSelStart = rtb.SelStart
    

End Sub

'
' Sub rtb_KeyDown
' Normally, tabbing leaves the control, but instead, we want to insert
' tab into edited text.
'
Private Sub rtb_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)

    If KeyCode = Asc(vbTab) Then  ' TAB key was pressed.
      ' Ignore the TAB key, so focus doesn't leave the control
      KeyCode = 0
      
      ' Replace selected text with the tab character
      rtb.SelText = vbTab
    End If


End Sub

'
' Sub HighlightRefresh
' Manipulate tracked previous selection and current selection
' to force reparsing of entire text.
'
Public Sub HighlightRefresh()
    'prevent textbox from repainting
    Dim x As Long
    x = SendMessage(rtb.hWnd, WM_SETREDRAW, 0, 0)

    Dim lOrigSelStart As Long
    Dim lOrigSelLength As Long
    lOrigSelStart = rtb.SelStart
    lOrigSelLength = rtb.SelLength

    mlPrevSelStart = 0
    rtb.SelStart = Len(rtb.Text)
    rtb_Change
    rtb.SelStart = lOrigSelStart
    rtb.SelLength = lOrigSelStart

    'allow texbox to repaint
    x = SendMessage(rtb.hWnd, WM_SETREDRAW, 1, 0)
    'force repaint
    rtb.Refresh
End Sub

'
' Sub ParseLines
' Feed text, line by line, to the parser.
'
Private Sub ParseLines(ByVal s As String, rtb As RichTextBox, ByVal RTBPos As Long)
    Dim lStartPos As Long
    Dim lEndPos As Long
    
    lStartPos = 1
    
    s = s & vbCrLf
    lEndPos = InStr(lStartPos, s, vbCrLf)
    Do While lEndPos > 0
        ParseLine Mid(s, lStartPos, lEndPos - lStartPos), rtb, RTBPos + lStartPos - 1
        lStartPos = lEndPos + Len(vbCrLf)
        lEndPos = InStr(lStartPos, s, vbCrLf)
    Loop
    
        
        
End Sub

'
' Sub ParseLine
' Lines are treated independently. Parseline is the main parsing code. Scan
' line from left to right, emitting text to be colored.
'
Private Sub ParseLine(ByVal s As String, rtb As RichTextBox, ByVal RTBPos As Long)
    'Debug.Print s
    
    Dim bInString As Boolean    ' are we in a quoted string?
    bInString = False
    
    Dim bInWord As Boolean      ' are we in a word? (not a string, comment,
                                ' or delimiter)
    bInWord = False
    
    Dim sCurString As String        ' the current set of characters
    Dim lCurStringStart As Long     '   - where it starts
    Dim sCurChar As String          ' the current character
    
    Dim i As Long
    
    For i = 1 To Len(s)
        sCurChar = Mid(s, i, 1)
        If sCurChar = COMMENT Then
            ' if comment character occurs within a quoted string, it doesn't
            ' count
            If Not bInString Then
                ' this is a comment. we are done with the line
                If bInWord Then
                    ' before we encounterd the comment we were processing a word
                    Highlight rtb, ParseWord(sCurString), lCurStringStart + RTBPos - 1, i - lCurStringStart
                    sCurString = ""
                    bInWord = False
                End If
            
                Highlight rtb, ColorComment, i + RTBPos - 1, Len(s) - i + 1
                GoTo ExitSub    ' rest of line is comment
            End If
        End If
        
        If sCurChar = """" Then
            ' if not already in a string, then this quote begins a string
            ' otherwise, we are in a string, and this quote ends it
            If bInString Then
                sCurString = sCurString & sCurChar
                Highlight rtb, ColorString, lCurStringStart + RTBPos - 1, i - lCurStringStart + 1
                sCurString = ""
                bInString = False
            Else
                If bInWord Then
                    ' before we encounterd the string we were processing a word
                    Highlight rtb, ParseWord(sCurString), lCurStringStart + RTBPos - 1, i - lCurStringStart
                    sCurString = ""
                    bInWord = False
                End If
                
                bInString = True
                sCurString = sCurChar
                lCurStringStart = i
            End If
            
            GoTo Next_i ' get next character
        End If
                
        If InStr(1, DELIMITER, sCurChar) > 0 Then
            If bInWord Then
                ' before we encounterd the delimiter we were processing a word
                Highlight rtb, ParseWord(sCurString), lCurStringStart + RTBPos - 1, i - lCurStringStart
                sCurString = ""
                bInWord = False
            End If
            
            Highlight rtb, ColorDelimiter, i + RTBPos - 1, 1
            GoTo Next_i
        End If
            
        If (Not bInWord) And (Not bInString) Then
            bInWord = True
            sCurString = sCurChar
            lCurStringStart = i
            
            GoTo Next_i ' get next character
        End If
            
        ' add current character to the "word" we are in the middle of
        sCurString = sCurString & sCurChar
Next_i:     ' VB style continue
    Next
    
    If bInString Then
        ' before we encounterd the end of the line we were processing a string
        Highlight rtb, ColorString, lCurStringStart + RTBPos - 1, i - lCurStringStart
    ElseIf bInWord Then
        ' before we encounterd the end of the line we were processing a word
        Highlight rtb, ParseWord(sCurString), lCurStringStart + RTBPos - 1, i - lCurStringStart
    End If

ExitSub:
    Exit Sub
End Sub

'
' Function ParseWord
' Determine color for this word by checking for its existence in the keyword
' lists. The word being checked it padded with spaces to prevent matches
' with substrings of keywords.
'
Private Function ParseWord(ByVal Word As String) As SyntaxTypes
    If InStr(1, RESERVED, KEYWORD_PAD & Word & KEYWORD_PAD, vbTextCompare) > 0 Then
        ParseWord = ColorReserved
    ElseIf InStr(1, FUNC_OBJ, KEYWORD_PAD & Word & KEYWORD_PAD, vbTextCompare) > 0 Then
        ParseWord = ColorFuncObj
    Else
        ParseWord = ColorNormal
    End If
End Function

'
' Sub Highlight
' Color this range in the RichTextBox. Note that you could also apply bold,
' italic, etc. to the selection at the same time.
'
Private Sub Highlight(rtb As RichTextBox, SyntaxType As SyntaxTypes, StartPos As Long, Length As Long)
        rtb.SelStart = StartPos
        rtb.SelLength = Length
    
    Select Case SyntaxType
        Case SyntaxTypes.ColorComment
            rtb.SelColor = mrgbComment
        Case SyntaxTypes.ColorString
            rtb.SelColor = mrgbString
        Case SyntaxTypes.ColorReserved
            rtb.SelColor = mrgbReserved
        Case SyntaxTypes.ColorFuncObj
            rtb.SelColor = mrgbFuncObj
        Case SyntaxTypes.ColorDelimiter
            rtb.SelColor = mrgbDelimiter
        Case Else
            rtb.SelColor = mrgbNormal
    End Select

End Sub

'
' Sub UserControl_Resize
' Constituate control is always same size as user control.
'
Private Sub UserControl_Resize()
    rtb.Width = UserControl.ScaleWidth
    rtb.Height = UserControl.ScaleHeight
End Sub

' *****************************************************************************
' Properties
' For the most part this code is generated by the VB ActiveX Control Wizard
' *****************************************************************************


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtb,rtb,-1,AutoVerbMenu
Public Property Get AutoVerbMenu() As Boolean
Attribute AutoVerbMenu.VB_Description = "Returns/sets a value that indicating whether the selected object's verbs will be displayed in a popup menu when the right mouse button is clicked."
    AutoVerbMenu = rtb.AutoVerbMenu
End Property

Public Property Let AutoVerbMenu(ByVal New_AutoVerbMenu As Boolean)
    rtb.AutoVerbMenu() = New_AutoVerbMenu
    PropertyChanged "AutoVerbMenu"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtb,rtb,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color of an object."
    BackColor = rtb.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    rtb.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get ForeColor() As Long
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As Long)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtb,rtb,-1,BorderStyle
Public Property Get BorderStyle() As BorderStyleConstants
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = rtb.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleConstants)
    rtb.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtb,rtb,-1,BulletIndent
Public Property Get BulletIndent() As Single
Attribute BulletIndent.VB_Description = "Returns or sets the amount of indent used in a RichTextBox control when SelBullet is set to True."
    BulletIndent = rtb.BulletIndent
End Property

Public Property Let BulletIndent(ByVal New_BulletIndent As Single)
    rtb.BulletIndent() = New_BulletIndent
    PropertyChanged "BulletIndent"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtb,rtb,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = rtb.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    rtb.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtb,rtb,-1,FileName
Public Property Get FileName() As String
Attribute FileName.VB_Description = "Returns/sets the filename of the file loaded into the RichTextBox control at design time."
    FileName = rtb.FileName
End Property

Public Property Let FileName(ByVal New_FileName As String)
    rtb.FileName() = New_FileName
    
    PropertyChanged "FileName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtb,rtb,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = rtb.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set rtb.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtb,rtb,-1,HideSelection
Public Property Get HideSelection() As Boolean
Attribute HideSelection.VB_Description = "Returns/sets a value that specifies if the selected item remains highlighted when a control loses focus."
    HideSelection = rtb.HideSelection
End Property

Public Property Let HideSelection(ByVal New_HideSelection As Boolean)
    rtb.HideSelection() = New_HideSelection
    PropertyChanged "HideSelection"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = m_hWnd
End Property

Public Property Let hWnd(ByVal New_hWnd As Long)
    m_hWnd = New_hWnd
    PropertyChanged "hWnd"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtb,rtb,-1,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Returns/sets a value indicating whether the contents in a RichTextBox control can be edited."
    Locked = rtb.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    rtb.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtb,rtb,-1,MaxLength
Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Returns/sets a value indicating whether there is a maximum number of characters a RichTextBox control can hold and, if so, specifies the maximum number of characters."
    MaxLength = rtb.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    rtb.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtb,rtb,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = rtb.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set rtb.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtb,rtb,-1,MousePointer
Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets a value indicating the type of mouse pointer displayed when the mouse is over the control at run time."
    MousePointer = rtb.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    rtb.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtb,rtb,-1,RightMargin
Public Property Get RightMargin() As Single
Attribute RightMargin.VB_Description = "Sets the right margin used for textwrap, centering, etc."
    RightMargin = rtb.RightMargin
End Property

Public Property Let RightMargin(ByVal New_RightMargin As Single)
    rtb.RightMargin() = New_RightMargin
    PropertyChanged "RightMargin"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtb,rtb,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in an object."
    Text = rtb.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    mbInChange = True
    rtb.Text() = New_Text
    mbInChange = False
    HighlightRefresh
    
    PropertyChanged "Text"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtb,rtb,-1,Find
Public Function Find(ByVal bstrString As String, Optional ByVal vStart As Variant, Optional ByVal vEnd As Variant, Optional ByVal vOptions As Variant) As Long
Attribute Find.VB_Description = "Searches the text in a RichTextBox control for a given string."
    Find = rtb.Find(bstrString, vStart, vEnd, vOptions)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtb,rtb,-1,GetLineFromChar
Public Function GetLineFromChar(ByVal lChar As Long) As Long
Attribute GetLineFromChar.VB_Description = "Returns the number of the line containing a specified character position in a RichTextBox control."
    GetLineFromChar = rtb.GetLineFromChar(lChar)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtb,rtb,-1,LoadFile
Public Sub LoadFile(ByVal bstrFilename As String, Optional ByVal vFileType As Variant)
Attribute LoadFile.VB_Description = "Loads an .RTF file or text file into a RichTextBox control."
    rtb.LoadFile bstrFilename, vFileType
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtb,rtb,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a control."
    rtb.Refresh
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtb,rtb,-1,SaveFile
Public Sub SaveFile(ByVal bstrFilename As String, Optional ByVal vFlags As Variant)
Attribute SaveFile.VB_Description = "Saves the contents of a RichTextBox control to a file."
    rtb.SaveFile bstrFilename, vFlags
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtb,rtb,-1,SelPrint
Public Sub SelPrint(ByVal lHDC As Long, Optional ByVal vStartDoc As Variant)
Attribute SelPrint.VB_Description = "Sends formatted text in a RichTextBox control to a device for printing."
    rtb.SelPrint lHDC, vStartDoc
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtb,rtb,-1,Span
Public Sub Span(ByVal bstrCharacterSet As String, Optional ByVal vForward As Variant, Optional ByVal vNegate As Variant)
Attribute Span.VB_Description = "Selects text in a RichTextBox control based on a set of specified characters."
    rtb.Span bstrCharacterSet, vForward, vNegate
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtb,rtb,-1,UpTo
Public Sub UpTo(ByVal bstrCharacterSet As String, Optional ByVal vForward As Variant, Optional ByVal vNegate As Variant)
Attribute UpTo.VB_Description = "Moves the insertion point up to, but not including, the first character that is a member of the specified character set in a RichTextBox control."
    rtb.UpTo bstrCharacterSet, vForward, vNegate
End Sub

Private Sub rtb_Click()
    RaiseEvent Click
End Sub

Private Sub rtb_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub rtb_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub rtb_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub rtb_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub rtb_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub rtb_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_ForeColor = m_def_ForeColor
    m_hWnd = m_def_hWnd
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    ' prevent parsing while file is loading
    mbInChange = True
    
    rtb.AutoVerbMenu = PropBag.ReadProperty("AutoVerbMenu", False)
    rtb.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    rtb.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    rtb.BulletIndent = PropBag.ReadProperty("BulletIndent", 0)
    rtb.Enabled = PropBag.ReadProperty("Enabled", True)
    rtb.FileName = PropBag.ReadProperty("FileName", "")
    Set rtb.Font = PropBag.ReadProperty("Font", Ambient.Font)
    rtb.HideSelection = PropBag.ReadProperty("HideSelection", True)
    m_hWnd = PropBag.ReadProperty("hWnd", m_def_hWnd)
    rtb.Locked = PropBag.ReadProperty("Locked", False)
    rtb.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    rtb.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    rtb.RightMargin = PropBag.ReadProperty("RightMargin", 0)
    rtb.Text = PropBag.ReadProperty("Text", "")
    
    mbInChange = False
    HighlightRefresh
    
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("AutoVerbMenu", rtb.AutoVerbMenu, False)
    Call PropBag.WriteProperty("BackColor", rtb.BackColor, &H80000005)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("BorderStyle", rtb.BorderStyle, 1)
    Call PropBag.WriteProperty("BulletIndent", rtb.BulletIndent, 0)
    Call PropBag.WriteProperty("Enabled", rtb.Enabled, True)
    Call PropBag.WriteProperty("FileName", rtb.FileName, "")
    Call PropBag.WriteProperty("Font", rtb.Font, Ambient.Font)
    Call PropBag.WriteProperty("HideSelection", rtb.HideSelection, True)
    Call PropBag.WriteProperty("hWnd", m_hWnd, m_def_hWnd)
    Call PropBag.WriteProperty("Locked", rtb.Locked, False)
    Call PropBag.WriteProperty("MaxLength", rtb.MaxLength, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", rtb.MousePointer, 0)
    Call PropBag.WriteProperty("RightMargin", rtb.RightMargin, 0)
    Call PropBag.WriteProperty("Text", rtb.Text, "")
End Sub

' *****************************************************************************
' Run Time Only Properties
' NOT generated by the ActiveX Control Wizard. Each of these procedures has
' its Procedure Attribute "Don't Show In Property Browser" set to true.
' *****************************************************************************

Public Property Get SelAlignment() As SelAlignmentConstants
Attribute SelAlignment.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If
    
    SelAlignment = rtb.SelAlignment
End Property

Public Property Let SelAlignment(ByVal New_SelAlignment As SelAlignmentConstants)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If
    
    rtb.SelAlignment = New_SelAlignment
End Property

Public Property Get SelBold() As Boolean
Attribute SelBold.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If
    
    SelBold = rtb.SelBold
End Property

Public Property Let SelBold(ByVal New_SelBold As Boolean)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If
    
    rtb.SelBold = New_SelBold
End Property

Public Property Get SelItalic() As Boolean
Attribute SelItalic.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If
    
    SelItalic = rtb.SelItalic
End Property

Public Property Let SelItalic(ByVal New_SelItalic As Boolean)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If
    
    rtb.SelItalic = New_SelItalic
End Property

Public Property Get SelStrikethru() As Boolean
Attribute SelStrikethru.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If
    
    SelStrikethru = rtb.SelStrikethru
End Property

Public Property Let SelStrikethru(ByVal New_SelStrikethru As Boolean)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If
    
    rtb.SelStrikethru = New_SelStrikethru
End Property

Public Property Get SelUnderline() As Boolean
Attribute SelUnderline.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If
    
    SelUnderline = rtb.SelUnderline
End Property

Public Property Let SelUnderline(ByVal New_SelUnderline As Boolean)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If
    
    rtb.SelUnderline = New_SelUnderline
End Property

Public Property Get SelBullet() As Variant
Attribute SelBullet.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If
    
    SelBullet = rtb.SelBullet
End Property

Public Property Let SelBullet(ByVal New_SelBullet As Variant)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If
    
    rtb.SelBullet = New_SelBullet
End Property

Public Property Get SelCharOffset() As Variant
Attribute SelCharOffset.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If
    
    SelCharOffset = rtb.SelCharOffset
End Property

Public Property Let SelCharOffset(ByVal New_SelCharOffset As Variant)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If
    
    rtb.SelCharOffset = New_SelCharOffset
End Property

Public Property Get SelRTF() As String
Attribute SelRTF.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If
    
    SelRTF = rtb.SelRTF
End Property

Public Property Let SelRTF(ByVal New_SelRTF As String)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If
    
    rtb.SelRTF = New_SelRTF
End Property

Public Property Get SelTabCount() As Integer
Attribute SelTabCount.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If
    
    SelTabCount = rtb.SelTabCount
End Property

Public Property Let SelTabCount(ByVal New_SelTabCount As Integer)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If
    
    rtb.SelTabCount = New_SelTabCount
End Property

Public Property Get SelTabs(Index As Integer) As Integer
Attribute SelTabs.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If
    
    SelTabs = rtb.SelTabs(Index)
End Property

Public Property Let SelTabs(Index As Integer, ByVal New_SelTabs As Integer)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If
    
    rtb.SelTabs(Index) = New_SelTabs
End Property

Public Property Get SelColor() As Variant
Attribute SelColor.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If
    
    SelColor = rtb.SelColor
End Property

Public Property Let SelColor(ByVal New_SelColor As Variant)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If
    
    rtb.SelColor = New_SelColor
End Property

Public Property Get SelHangingIndent() As Integer
Attribute SelHangingIndent.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If
    
    SelHangingIndent = rtb.SelHangingIndent
End Property

Public Property Let SelHangingIndent(ByVal New_SelHangingIndent As Integer)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If
    
    rtb.SelHangingIndent = New_SelHangingIndent
End Property

Public Property Get SelIndent() As Integer
Attribute SelIndent.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If
    
    SelIndent = rtb.SelIndent
End Property

Public Property Let SelIndent(ByVal New_SelIndent As Integer)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If
    
    rtb.SelIndent = New_SelIndent
End Property

Public Property Get SelRightIndent() As Integer
Attribute SelRightIndent.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If
    
    SelRightIndent = rtb.SelRightIndent
End Property

Public Property Let SelRightIndent(ByVal New_SelRightIndent As Integer)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If
    
    rtb.SelRightIndent = New_SelRightIndent
End Property

Public Property Get SelLength() As Long
Attribute SelLength.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If
    
    SelLength = rtb.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If
    
    rtb.SelLength = New_SelLength
End Property

Public Property Get SelStart() As Long
Attribute SelStart.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If
    
    SelStart = rtb.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If
    
    rtb.SelStart = New_SelStart
End Property

Public Property Get SelText() As String
Attribute SelText.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If
    
    SelText = rtb.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If
    
    rtb.SelText = New_SelText
End Property

Public Property Get SelProtected() As Variant
Attribute SelProtected.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If
    
    SelProtected = rtb.SelProtected
End Property

Public Property Let SelProtected(ByVal New_SelProtected As Variant)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If
    
    rtb.SelProtected = New_SelProtected
End Property

Public Property Get TextRTF() As String
Attribute TextRTF.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If
    TextRTF = rtb.TextRTF
End Property

Public Property Let TextRTF(ByVal New_TextRTF As String)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If
        
    mbInChange = True
    rtb.TextRTF = New_TextRTF
    mbInChange = False
    
    HighlightRefresh

End Property
