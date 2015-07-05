Attribute VB_Name = "ModServerSettings"
Option Explicit

Type RegRec
Name As String
regdate As String
regtime As String
End Type

Public Reg(1 To 1) As RegRec

Sub LoadREG()
Dim filename As String
Dim I As Long
Dim f As Long
    Call CheckREG
    For I = 1 To 1
        Call SetStatus("Verifying Database Files... " & Int((I / 1) * 100) & "%")
        filename = App.Path & "\Main\support\crypto" & I & ".dat"
        f = FreeFile
        Open filename For Binary As #f
        Get #f, , Reg(I)
        Close #f
        DoEvents
    Next
End Sub
Sub CheckREG()
Dim filename As String
Call ClearReg
Dim I As Long
    For I = 1 To 1
        filename = "Main\support\crypto" & I & ".dat"
        If Not FileExist(filename) Then
            Call SetStatus("Updating Database Files... " & Int((I / 1) * 100) & "%")
            DoEvents
            Call SaveREG(I)
        End If
    Next
End Sub
Sub SaveREG(ByVal RegNum As Long)
Dim filename As String
Dim f As Long
    filename = App.Path & "\Main\support\crypto" & RegNum & ".dat"
    f = FreeFile
    Open filename For Binary As #f
    Put #f, , Reg(RegNum)
    Close #f
End Sub
Sub ClearReg()
Reg(1).Name = ""
End Sub
