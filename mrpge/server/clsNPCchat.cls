VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNPCchat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'Private strChatText As String
'Private strAnswers As NpcChatAns
'Private lngPrevAns As Long''

'Public Property Get ChatText() As String
'    ChatText = strChatText
'End Property

'Public Property Get Answers() As NpcChatAns
'    Answers = strAnswers
'End Property

'Sub LoadNpcChat(ByVal Level As Long, ByVal PrevAnswer As Long, ByVal NpcNum As Long)
'    strChatText = GetVar(App.Path & "\NPC AI\talk\" & NpcNum & ".txt", "ChatLev" & Level, "ChatTxt")
'    strAnswers.ans1 = GetVar(App.Path & "\NPC AI\talk\" & NpcNum & ".txt", "ChatLev" & Level, "Ans1")
'    strAnswers.ans2 = GetVar(App.Path & "\NPC AI\talk\" & NpcNum & ".txt", "ChatLev" & Level, "Ans2")
'    strAnswers.ans3 = GetVar(App.Path & "\NPC AI\talk\" & NpcNum & ".txt", "ChatLev" & Level, "Ans3")
'    strAnswers.ans4 = GetVar(App.Path & "\NPC AI\talk\" & NpcNum & ".txt", "ChatLev" & Level, "Ans4")
'End Sub
