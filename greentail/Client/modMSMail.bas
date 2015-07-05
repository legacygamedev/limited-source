Attribute VB_Name = "modMSMail"
'=========================================================================================
'  This Module was adapted from external source
'  Published Date: 29/09/2003
'  WebSite: http://www.deepeshagarwal.tk
' E-mail: agarwal_deepesh@indiatimes.com
'=========================================================================================
Option Explicit
 
Type MAPIMessage
    Reserved As Long
    Subject As String
    NoteText As String
    MessageType As String
    DateReceived As String
    ConversationID As String
    Flags As Long
    RecipCount As Long
    FileCount As Long
End Type
 
Type MapiRecip
    Reserved As Long
    RecipClass As Long
    Name As String
    Address As String
 
    EIDSize As Long
    EntryID As String
End Type
 
Type MapiFile
    Reserved As Long
    Flags As Long
    Position As Long
    PathName As String
    filename As String
    FileType As String
End Type
 
Declare Function MAPISendMail Lib "MAPI32.DLL" Alias "BMAPISendMail" (ByVal Session&, ByVal UIParam&, Message As MAPIMessage, Recipient() As MapiRecip, File() As MapiFile, ByVal Flags&, ByVal Reserved&) As Long
 

Global Const SUCCESS_SUCCESS = 0
 
Global Const MAPI_TO = 1
Global Const MAPI_CC = 2
Global Const MAPI_BCC = 3
 
Global Const MAPI_LOGON_UI = &H1
 
Function CountTokens(ByVal pstrSource As String, ByVal pstrDelim As String)
    On Error GoTo CountTokens_Err
    '*************************************************************
    
    ' FUNCTION NAME: CountTokens
    '
    ' PURPOSE:
    '   Given a string of delimited items and the delimiter, the number
    '   of tokens in the string will be returned. This function is useful
    '   for dimensioning an array to store the delimited items prior to
    '   calling ParseTokens.
    '
    ' INPUT PARAMETERS:
    '   pstrSource: A delimited list of tokens
    '   pstrDelim:  The delimiter used to delimit pstrSource
    '
    ' RETURN
    '   The number of tokens in pstrSource, which is the number of delimiters
    
    '   plus 1. If pstrSource is empty, 0 is returned.
    
    '*************************************************************
 
    Dim iDelimPos As Integer
    Dim iCount As Integer
 
    ' Number of tokens = 0 if the source string is empty
    If pstrSource = "" Then
        CountTokens = 0
 
    ' Otherwise number of tokens = number of delimiters + 1
    Else
        iDelimPos = InStr(1, pstrSource, pstrDelim)
 
        Do Until iDelimPos = 0
        iCount = iCount + 1
        iDelimPos = InStr(iDelimPos + 1, pstrSource, pstrDelim)
        Loop
        CountTokens = iCount + 1
    End If
 

CountTokens_Exit:
    On Error Resume Next
    Exit Function
 
CountTokens_Err:
    Select Case Err
        Case Else
            MsgBox "Error " & Err.Number & " in CountTokens()"
    End Select
    Resume CountTokens_Exit
    Resume
 
End Function
 
Function GetToken(sSource As String, ByVal sDelim As String) As String
    On Error GoTo GetToken_Err
    '*************************************************************
    ' FUNCTION NAME: GetToken
    '
    ' PURPOSE:
    '   Given a string of delimited items, the first item will be
    '   removed from the list and returned.
    '
    ' INPUT PARAMETERS:
    '   sSource: A delimited list of tokens
    
    '   sDelim:  The delimiter used to delimit sSource
    '
    ' RETURN
    '   sSource will have the first token removed. The function
    '   returns the token removed from sSource.
    
    '*************************************************************
 
    Dim iDelimPos As Integer
 
    ' Find the first delimiter
    iDelimPos = InStr(1, sSource, sDelim)
 
    ' If no delimiter was found, return the existing string and set
 
    ' .. the source to an empty string.
    If (iDelimPos = 0) Then
        GetToken = Trim$(sSource)
        sSource = ""
 
    ' Otherwise, return everything to the left of the delimiter and
    ' .. return the source string with it removed.
    Else
        GetToken = Trim$(Left$(sSource, iDelimPos - 1))
        sSource = Mid$(sSource, iDelimPos + 1)
    End If
 

GetToken_Exit:
    On Error Resume Next
    Exit Function
 
GetToken_Err:
    Select Case Err
        Case Else
            MsgBox Err, Error, 16, "Error " & Err & " in GetToken()"
    End Select
    Resume GetToken_Exit
    Resume
 
End Function
 
Function Mail()
    On Error GoTo Mail_Err
    '*************************************************************
    ' FUNCTION NAME: Mail
    '
    ' PURPOSE:
    '   Passes information on the active forms To, Subject, CC,
 
    '   Attach, and Message text boxes to the SendMail function.
    '   It ensures that each box does not have a NULL value. It also
    '   displays an error message if SendMail fails.
    '   This function is called from the OnPush property of the form.
    '
    ' INPUT PARAMETERS:
    '   None
    '
    ' RETURN
    '   None
 
    '*************************************************************
        
    Dim F As Form, result
    Set F = Screen.ActiveForm
 
    ' Make sure user has something in the To: box
    If IsNull(F!txtTo) Or F!txtTo = "" Then Exit Function
 
    ' Make sure no Null values are in the other boxes
    If IsNull(F!txtsubject) Then F!txtsubject = ""
    If IsNull(F!txtCC) Then F!txtCC = ""
    If IsNull(F!txtAttach) Then F!txtAttach = ""
    If IsNull(F!txtMessage) Then F!txtMessage = ""
 
    ' Send the message, passing information from the form
    result = SendMail((F!txtsubject), (F!txtTo), (F!txtCC), (F!txtAttach), (F!txtMessage))
 
    ' Test the result for any errors
    If result <> SUCCESS_SUCCESS Then
        MsgBox "Error sending mail: " & result, 16, "Mail"
        'MsgBox tLookup("Msg", "zstblMAPIError", "ErrNo=" & result), 16, "MAPI Error #" & result
    Else
        MsgBox "Message sent successfully!", 64, "Mail"
    End If
 

Mail_Exit:
    On Error Resume Next
    Exit Function
 
Mail_Err:
    Select Case Err
        Case Else
            MsgBox Err, Error, 16, "Error " & Err.Description & " in Mail()"
    End Select
    Resume Mail_Exit
    Resume
 
End Function
 
Function SendMail(sSubject As String, sTo As String, sCC As String, sAttach As String, sMessage As String)
    On Error GoTo SendMail_Err
 
    '*************************************************************
    ' FUNCTION NAME: SendMail
    '
    ' PURPOSE:
    '   This is the front-end function to the MAPISendMail function. You
    '   pass a semicolon-delimited list of To and CC recipients, a
    
    '   subject, a message, and a delimited list of file attachments.
    '   This function prepares MapiRecip and MapiFile structures with the
    '   data parsed from the information provided using the ParseRecord
    '   sub. Once the structures are prepared, the MAPISendMail function
    '   is called to send the message.
    '
    ' INPUT PARAMETERS:
    '   sSubject: The text to appear in the subject line of the message
    '   sTo:      Semicolon-delimited list of names to receive the
    
    '             message
    '   sCC:      Semicolon-delimited list of names to be CC'd
    '   sAttach:  Semicolon-delimited list of files to attach to
    '             the message
    ' RETURN
    '   SUCCESS_SUCCESS if successful, or a MAPI error if not.
    
    '*********************************************************** **
 
    Dim i, cTo, cCC, cAttach          ' variables holding counts
 
    Dim MAPI_Message As MAPIMessage
 
    ' Count the number of items in each piece of the mail message
    cTo = CountTokens(sTo, ";")
    cCC = CountTokens(sCC, ";")
    cAttach = CountTokens(sAttach, ";")
 
    ' Create arrays to store the semicolon delimited mailing
    ' .. information after it is parsed
    ReDim rTo(0 To cTo) As String
    ReDim rCC(0 To cCC) As String
    ReDim rAttach(0 To cAttach) As String
 
    ' Parse the semicolon delimited information into the arrays.
 
    ParseTokens rTo(), sTo, ";"
    ParseTokens rCC(), sCC, ";"
    ParseTokens rAttach(), sAttach, ";"
 
    ' Create the MAPI Recip structure to store all the To and CC
    ' .. information to be passed to the MAPISendMail function
    ReDim MAPI_Recip(0 To cTo + cCC - 1) As MapiRecip
 
    ' Setup the "TO:" recipient structures
    For i = 0 To cTo - 1
        MAPI_Recip(i).Name = rTo(i)
        MAPI_Recip(i).RecipClass = MAPI_TO
    Next i
 
    ' Setup the "CC:" recipient structures
    For i = 0 To cCC - 1
        MAPI_Recip(cTo + i).Name = rCC(i)
        MAPI_Recip(cTo + i).RecipClass = MAPI_CC
    Next i
 

    ' Create the MAPI File structure to store all the file attachment
    ' .. information to be passed to the MAPISendMail function
    ReDim MAPI_File(0 To cAttach) As MapiFile
 
    ' Setup the file attachment structures
    MAPI_Message.FileCount = cAttach
    For i = 0 To cAttach - 1
 
        MAPI_File(i).Position = -1
        MAPI_File(i).PathName = rAttach(i)
    Next i
 
    ' Set the mail message fields
    MAPI_Message.Subject = sSubject
    MAPI_Message.NoteText = sMessage
    MAPI_Message.RecipCount = cTo + cCC
 
    ' Send the mail message
    SendMail = MAPISendMail(0&, 0&, MAPI_Message, MAPI_Recip(), MAPI_File(), MAPI_LOGON_UI, 0&)
 

SendMail_Exit:
    On Error Resume Next
    Exit Function
 
SendMail_Err:
    Select Case Err
        Case Else
            MsgBox Err, Error, 16, "Error " & Err & " in SendMail()"
    End Select
    Resume SendMail_Exit
    Resume
 
End Function
 
Sub ParseTokens(pstrArray() As String, ByVal sTokens As String, ByVal sDelim As String)
    On Error GoTo ParseTokens_Err
 
    '*************************************************************
    
    ' SUB NAME: ParseTokens
    '
    ' PURPOSE:
    '   Extracts information from a delimited list of items and places
    '   it in an array.
    '
    ' INPUT PARAMETERS:
    '   Array(): A one-dimensional array of strings in which the parsed
    '   tokens will be place
    '   sTokens: A delimited list of tokens
    '   sDelim:  The delimiter used to delimit sTokens
    '
    ' RETURN
    '   None
    
    '*************************************************************
    Dim i As Integer
    For i = LBound(pstrArray) To UBound(pstrArray)
        pstrArray(i) = GetToken(sTokens, sDelim)
    Next
 

ParseTokens_Exit:
    On Error Resume Next
    Exit Sub
 
ParseTokens_Err:
    Select Case Err
        Case Else
            MsgBox Err, Error, 16, "Error " & Err & " in ParseTokens()"
    End Select
    Resume ParseTokens_Exit
    Resume
 
End Sub
 
Sub SendReport()
    Dim strrep As String
    Dim varerror As Variant
    'The path to the saved report
    Const strreport As String = "c:\autoexec.bat" '"c:\temp\SystemInfo.txt"
    'The Message
    Const strmess As String = "Please find attached the SSI Report, which was generated on Machine Name X , By User Y." & _
    "" & vbCrLf & vbCrLf & "The Makers of SSI Accept no responsibility for this Email"
    'The Subject
    Const strsub As String = "SSI Report"
    
    'This accepts an Email address can be replaced with a form etc...
    strrep = InputBox("Please Enter the recipient for this SSI Report:", "Email Report")
    If Len("" & strrep) = 0 Then
        MsgBox "Email Canceled", vbInformation + vbOKOnly, "Warning!"
        Exit Sub
    End If
    If ValidEmail(strrep) = True Then
        varerror = SendMail(strsub, strrep, "", strreport, strmess)
        If varerror <> 0 Then
            'This should be where the errors are reported
            'using the error numbers i gave you.
            Select Case varerror
                Case 3
                    MsgBox "Error " & varerror & ", You could not be logged in to your E-mail Client or User Cancelled log on!", vbCritical + vbOKOnly, "Error"
                Case Else
                    MsgBox "The Email was not sent. An error has ocurred, Error " & varerror & "!", vbCritical + vbOKOnly, "Error"
            End Select
        End If
    Else
        MsgBox "That E-mail address is not valid!", vbInformation + vbOKOnly, "Warning!"
    End If
End Sub
 
Function ValidEmail(strin As String)
    'Performs a simple check on the Email address to see
    'if it contains a @ and a .
    If InStr(strin, "@") And InStr(strin, ".") Then
        ValidEmail = True
    Else
        ValidEmail = False
    End If
    
End Function
 

