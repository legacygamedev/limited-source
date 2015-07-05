Attribute VB_Name = "modInformation"
Option Explicit

Public Const VERSION As String = "1.0.0a"
Public Ext_Path As String

Public Sub LoadHelpForm()
frmHelp.txtHelp.Text = "----Dual Solace Registering Application--" & vbCrLf & _
                       "--Version: " & VERSION & "--" & vbCrLf & _
                       "" & vbCrLf & _
                       "Compiler: Visual Basic 6" & vbCrLf & _
                       "Created by smchronos" & vbCrLf & _
                       "" & vbCrLf & _
                       "Certified by Dual Solace 2008" & vbCrLf & _
                       "" & vbCrLf & _
                       "" & vbCrLf & _
                       "Run this program once to generate two files:" & vbCrLf & vbCrLf & _
                       "config.ini points towards the folder where all extensions to register are located." & vbCrLf & _
                       "exts.ini gives helpful information about extensions." & vbCrLf & _
                       "" & vbCrLf & _
                       "*Note: All files listed on loadup will overwrite any existing copies in the system32 folder if registered!" & vbCrLf & _
                       "" & vbCrLf & _
                       "Please report all errors to www.dualsolace.com"

' Make the help form visible
frmHelp.Show
End Sub
