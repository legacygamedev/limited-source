Attribute VB_Name = "modMain"
Option Explicit

Public Codes As Object

Public Const DirUp = 0
Public Const DirDown = 1
Public Const DirLeft = 2
Public Const DirRight = 3

'Public Const LocalSocket = 1
Public Const Username = "TriviaBot"
Public Const Password = "botpass"
Public Const CharIndex = 1

Public Const BotMap = 1
Public Const BotX = 4
Public Const BotY = 3
Public Const BotDir = DirDown

Type TriviaQuestions
    Question As String
    Answer As String
End Type

Public Trivia() As TriviaQuestions
Public BotIndex As Integer
Public LocalSocket As Integer
Public PluginNum As Integer

Public SecondsLeft As Integer
Public SecondsTill As Integer
Public CurrentQuestion As String
Public CurrentAnswer As String

Public Function UpdateStatus(newStatus As String)
    frmMain.lblStatus.Caption = newStatus
End Function

Public Function SetupTrivia()
On Error Resume Next
    Dim tmpStr As String
    Dim TheFile As String
    Dim TrivCount As Long
    
    TheFile = App.Path & "\questions.txt"
    
    TrivCount = 0
    
    Open TheFile For Input As #1
    While Not EOF(1)
        Input #1, tmpStr
        
        TrivCount = TrivCount + 1
        ReDim Preserve Trivia(1 To TrivCount) As TriviaQuestions
        Trivia(TrivCount).Question = Split(tmpStr, ":")(0)
        Trivia(TrivCount).Answer = Split(tmpStr, ":")(1)
        
        DoEvents
    Wend
    Close #1
    
    Call SendEmoteMSG("Questions Loaded. Waiting 5 seconds to start.")
    Call NewGame
End Function

Public Function GetRandom() As Integer
Dim rndInt As Integer
Start:
    Randomize
    rndInt = Int(UBound(Trivia) * Rnd)
    If rndInt = 0 Then
        GoTo Start:
    End If
    GetRandom = rndInt
End Function

Public Function NewGame()
Dim TriviaNumber As Integer
    TriviaNumber = GetRandom()
    CurrentQuestion = Trivia(TriviaNumber).Question
    CurrentAnswer = LCase(Trivia(TriviaNumber).Answer)
    SecondsLeft = 65
    SecondsTill = 5
    frmMain.tmrNewGame.Enabled = True
End Function

Public Function CheckWin(Who As String, ThereMessage As String)
    If frmMain.tmrTrivia.Enabled = False Then Exit Function

    If Trim(LCase(ThereMessage)) = Trim(LCase(CurrentAnswer)) Then
        Call SendEmoteMSG(Who & " Has Won This Round! Lets Give Him/Her A Round Of Applause!")
        frmMain.tmrTrivia.Enabled = False
        frmMain.tmrNewGame.Enabled = False
        Call NewGame
    End If
End Function
