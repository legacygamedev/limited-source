Attribute VB_Name = "modFunctions"
Option Explicit

Public Function SendMapMSG(Message As String)
    Call Codes.SendTheMapMSG(Message, LocalSocket)
End Function

Public Function SendEmoteMSG(Message As String)
    Call Codes.SendTheEmoteMSG(Message, LocalSocket)
End Function

Public Function SendBroadcastMSG(Message As String)
    Call Codes.SendTheBroadcastMSG(Message, LocalSocket)
End Function

Public Function KickPlayer(Person As String)
    Call Codes.KickThePlayer(Person, BotIndex)
End Function
