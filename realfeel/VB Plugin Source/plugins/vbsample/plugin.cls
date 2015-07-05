VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "plugin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True


'Function RegisterPlugin(intMenu As Integer,              <-- which menu to add to
'                         strMenuName As String,          <-- menu text
'                         intStartupArgument As Integer)  <--argument to pass back on select


Private frmMain As Object

Sub SetHost(newref As Object)
    Set frmMain = newref
    frmMain.RegisterPlugin 0, "Sample VB Plugin", 0
End Sub

Sub StartUp(intArg As Integer)
    
    MsgBox "Vb Plugin - Argument =" & intArg
    MsgBox "FrmMain.Caption = " & frmMain.Caption
    
    'Select Case intArg
    '    Case 0:
    'End Select
    
End Sub

 

