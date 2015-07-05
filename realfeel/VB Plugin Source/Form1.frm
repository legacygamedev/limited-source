VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Plugin Host Example - http://sandsprite.com"
   ClientHeight    =   3300
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label3 
      Caption         =   $"Form1.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   60
      TabIndex        =   2
      Top             =   2040
      Width           =   5535
   End
   Begin VB.Label Label2 
      Caption         =   "    On its first run, make sure to let it register the dlls first for you by answering NO to the messagebox."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   60
      TabIndex        =   1
      Top             =   1080
      Width           =   5535
   End
   Begin VB.Label Label1 
      Caption         =   "Plugin Host - This exe will find and dynamically load all the plugins it finds in the /plugins/ directory. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   5475
   End
   Begin VB.Menu mnuPlugins 
      Caption         =   "Plugins"
      Begin VB.Menu mnuPluginList 
         Caption         =   "-"
         Index           =   0
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Author: David Zimmer
'Site:   http://sandsprite.com

'This is a simple efficient plugin host implementation...it is designed to be as
'readable as possible without to much bulk, to use this in a production quality
'application, you will need to add error handling, specifically around the
'createobject routines. If the createObject fails, it could be because the
'plugin was not yet registered, you can test for this with the err.number, if this
'is the case you can then programatically register it and try createobject again.
'If it still doesnt work, then you can reject it as not a plugin.
'
'-Dave


Dim plugins() As Object

Private Sub Form_Load()
    Dim tmp() As String, i As Integer, progid As String
    Dim wsc() As String
    
    On Error GoTo hell
    
    
    tmp() = GetFolderFiles(App.path & "\plugins", "*dll")
    wsc() = GetFolderFiles(App.path & "\plugins", "*wsc")
    
    If Not AryIsEmpty(wsc) Then    'add any windows script components to plugin list
        For i = 0 To UBound(wsc)   'they can be used as plugins too. (.wsc files)
            push tmp, wsc(i)
        Next
    End If
    
    
   'for the demo, we will just let the user register this way if they want
    If MsgBox("Did you register all of the dlls & the wsc file with regsvr32 already?", vbYesNo) = vbNo Then
         For i = 0 To UBound(tmp)
            Shell "regsvr32 """ & tmp(i) & """", vbNormalFocus
         Next
    End If
    
    ReDim plugins(0)
    
    For i = 0 To UBound(tmp)
        ReDim Preserve plugins(i)
        progid = GetBaseName(tmp(i)) & ".plugin"
        Set plugins(i) = CreateObject(progid)
        plugins(i).sethost Me
        
        '_____________________________________________________________________________________
        'pass in a ref to this form which supports IDispatch
        'plugin::SetHost(newref As Object)
        'sethost implementation will then call RegisterPlugin on this form
        'to register its functionality and we will preform the actual menu manipulations and
        'bookwork of associating x client with y menu item and z startup arg
        '-------------------------------------------------------------------------------------
        
    Next
    
Exit Sub
hell: MsgBox tmp(i) & " - " & Err.Description
      Resume Next
End Sub


Function RegisterPlugin(intMenu As Integer, strMenuName As String, intStartupArgument As Integer)
    'here right after sethost in loadplugins sub
    Dim i As Integer
    
    'If intMenu = 0 Then
        i = mnuPluginList.Count
        Load mnuPluginList(i)
        mnuPluginList(i).Caption = strMenuName
        mnuPluginList(i).Visible = True
        mnuPluginList(i).Tag = UBound(plugins) & "." & intStartupArgument
    'Else
     'same thing to some other menu
     
End Function

Private Sub mnuPluginList_Click(Index As Integer)
    Dim tmp() As String
    
    On Error GoTo hell
    
    'plugin was selected from the list fire it up with arg specified when it registered tiself
    'format of menuitem.Tag = index of object in plugin() array . startuparg it expects
    'notice how all plugins are added to menu item arrays, this way it is easy for us
    'to forward the events back to the clients.
    
    tmp = Split(mnuPluginList(Index).Tag, ".")
    
    plugins(CInt(tmp(0))).startup CInt(tmp(1))
    
    
    Exit Sub
hell:     MsgBox Err.Description
End Sub






'---------------------------------------------------------------
'general library support functions below
'---------------------------------------------------------------
Function GetFolderFiles(folder, Optional filter = ".*", Optional retFullPath As Boolean = True) As String()
   Dim fnames() As String, extension, fs
   
   If Not FolderExists(folder) Then
        'returns empty array if fails
        GetFolderFiles = fnames()
        Exit Function
   End If
   
   folder = IIf(Right(folder, 1) = "\", folder, folder & "\")
   If Left(filter, 1) = "*" Then extension = Mid(filter, 2, Len(filter))
   If Left(filter, 1) <> "." Then filter = "." & filter
   
   fs = Dir(folder & "*" & filter, vbHidden Or vbNormal Or vbReadOnly Or vbSystem)
   While fs <> ""
     If fs <> "" Then push fnames(), IIf(retFullPath = True, folder & fs, fs)
     fs = Dir()
   Wend
   
   GetFolderFiles = fnames()
End Function

Function FolderExists(path) As Boolean
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbDirectory) <> "" Then FolderExists = True
End Function

Function GetBaseName(path) As String
    Dim tmp() As String, ub
    tmp = Split(path, "\")
    ub = tmp(UBound(tmp))
    If InStr(1, ub, ".") > 0 Then
       GetBaseName = Mid(ub, 1, InStrRev(ub, ".") - 1)
    Else
       GetBaseName = ub
    End If
End Function

Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    Dim i As Integer
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function
