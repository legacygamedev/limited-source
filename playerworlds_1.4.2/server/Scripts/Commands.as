Sub ProcessCommand(Command, Data)
Call DebugPrint("Command: " & Command & ", Data: " & Data)

Select Case LCase(Command)
   Case "/print"
      Call DebugPrint("It works!!!")
      Call GlobalMessage("This is a global message!")
      Exit Sub

   Case "/who"
      
      Exit Sub
End Select   



End Sub