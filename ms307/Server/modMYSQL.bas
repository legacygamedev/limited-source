Attribute VB_Name = "modMYSQL"
Option Explicit

Public Function Cleanse(Dirty As String) As String
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 10/02/2003  Shannara   Created Function
'****************************************************************

'THIS FUNCTION WILL ESCAPE ALL SINGLE QUOTE CHARACTERS IN AN EFFORT
'TO PREVENT SQL INJECTION ATTACKS. IT IS RECCOMENDED THAT ALL TAINTED DATA BE
'PASSED THROUGH THIS FUNCTION PRIOR TO BEING USED IN DYNAMIC SQL QUERIES.
'
'*******************************************
'NOTE: YOUR BROWSER MAY SHOW SPACES IN THE STRINGS (I.E.  " '  " ) THERE SHOULD BE NO WHITESPACES IN ANY OF THE STRINGS
'*******************************************
'
'WRITTEN BY: MIKE HILLYER
'LAST MODIFIED: 14JUN2003
    Cleanse = Replace(Dirty, "'", "\'")
'CLEVER HACKERS COULD PASS \' TO THIS FUNCTION, WHICH WOULD BECOME \\'
' \\' GETS INTERPRETED AS \', WITH THE \ BEING IGNORED AND THE ' GETTING
'INTERPRETED, THUS BYPASSING THIS FUNCTION, SO WE SHALL LOOP UNTIL WE ARE LEFT
'WITH JUST \' WHICH ESCAPES THE QUOTE, LOOP IS NEEDED BECAUSE A HACKER COULD TYPE
' \\\' IF WE SIMPLY CHECKED FOR \\' AFTER DOING THE INITIAL REPLACE.
    Do While InStr(Cleanse, "\\'")
        Cleanse = Replace(Cleanse, "\\'", "\'")
    Loop
End Function

Public Function DBAccountExist(ByVal Name As String) As Boolean
Dim Conn As ADODB.Connection
Dim RS As ADODB.Recordset
Dim cLogin As String
Dim iExist As Boolean

'Cleanse the strings
cLogin = Cleanse(Trim(Name))

'Setup database stuff
Set Conn = New ADODB.Connection
Set RS = New ADODB.Recordset

Conn.CursorLocation = adUseServer
Conn.ConnectionString = strCONN
Conn.Open
RS.Open "SELECT * FROM tb_accounts WHERE `login`='" & cLogin & "'", Conn, adOpenStatic, adLockReadOnly

iExist = False
If Not (RS.BOF And RS.EOF) Then
    RS.MoveFirst
    iExist = True
Else
    iExist = False
End If
RS.Close
Conn.Close

Set RS = Nothing
Set Conn = Nothing

DBAccountExist = iExist

End Function

Public Function DBCharExist(ByVal Index As Long, ByVal CharNum As Long) As Boolean
    If Trim(Player(Index).Char(CharNum).Name) <> "" Then
        CharExist = True
    Else
        CharExist = False
    End If
End Function

Public Function DBPasswordOK(ByVal Name As String, ByVal Password As String) As Boolean
Dim cLogin As String
Dim cPassword As String
Dim iExist As Boolean
Dim Conn As ADODB.Connection
Dim RS As ADODB.Recordset

'Cleanse the strings
cLogin = Cleanse(Trim(Login))
cPassword = Password

'Setup database stuff
Set Conn = New ADODB.Connection
Set RS = New ADODB.Recordset

Conn.CursorLocation = adUseServer
Conn.ConnectionString = strCONN
Conn.Open
RS.Open "SELECT * FROM tb_accounts WHERE `login`='" & cLogin & "' AND `password`='" & cPassword & "'", Conn, adOpenStatic, adLockReadOnly

iExist = False
If Not (RS.BOF And RS.EOF) Then
    RS.MoveFirst
    iExist = True
Else
    iExist = False
End If
RS.Close
Conn.Close

Set RS = Nothing
Set Conn = Nothing

DBPasswordOK = iExist
End Function


