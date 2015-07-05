Attribute VB_Name = "modCustomScripts"
Option Explicit

Public Sub CustomScript(Index As Long, caseID As Long)
    Select Case caseID
        Case Else
            PlayerMsg Index, "You just activated custom script " & caseID & ". This script is not yet programmed.", BrightRed
    End Select
End Sub
