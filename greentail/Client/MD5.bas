Attribute VB_Name = "MD5"
Option Explicit

Function XOREncrypt(PlainText As String, Password As String) As String
    Dim PTLength As Long
    Dim PWDLength As Integer
    Dim X As Long
   
    PTLength = Len(PlainText) - 1
    PWDLength = Len(Password)
    XOREncrypt = ""
   
 
    For X = 0 To PTLength
        XOREncrypt = XOREncrypt + CStr(Asc(Mid$(PlainText, X + 1, 1)) Xor Asc(Mid$(Password, (X Mod PWDLength) + 1, 1))) + " "
    Next X

    XOREncrypt = Trim$(XOREncrypt)
End Function

Function XORDecrypt(Cipher As String, Password As String) As String
    Dim TempArray As Variant
                               
    Dim X As Long
    Dim PWDLength As Integer
   
   
    On Error Resume Next
    'leave
   
    TempArray = Split(Cipher, " ")
    PWDLength = Len(Password)
    XORDecrypt = ""
   
    For X = 0 To UBound(TempArray)
        XORDecrypt = XORDecrypt + Chr(Int(TempArray(X)) Xor Asc(Mid$(Password, (X Mod PWDLength) + 1, 1)))
    Next X
End Function
