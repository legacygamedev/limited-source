VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form frmUpdate 
   Caption         =   "M:RPGe Updater"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4125
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   4125
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   450
      Left            =   1455
      TabIndex        =   1
      Top             =   4020
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock socketUpdate 
      Left            =   0
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "root12.co.uk"
      RemotePort      =   80
   End
   Begin VB.PictureBox picLogo 
      Height          =   2535
      Left            =   0
      ScaleHeight     =   2475
      ScaleWidth      =   4035
      TabIndex        =   0
      Top             =   0
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Overall Progress:"
      Height          =   270
      Left            =   480
      TabIndex        =   3
      Top             =   3360
      Width           =   3030
   End
   Begin VB.Label lblCurrentFile 
      Height          =   270
      Left            =   465
      TabIndex        =   2
      Top             =   2595
      Width           =   3030
   End
   Begin VB.Shape shpOverall 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   255
      Left            =   495
      Top             =   3600
      Width           =   3000
   End
   Begin VB.Shape Shape2 
      Height          =   270
      Left            =   480
      Top             =   3585
      Width           =   3015
   End
   Begin VB.Shape shpCurFile 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   255
      Left            =   495
      Top             =   2880
      Width           =   3000
   End
   Begin VB.Shape shpCurrentFile 
      Height          =   270
      Left            =   480
      Top             =   2865
      Width           =   3015
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private gettingFileList As Boolean
Private updating As Boolean
Private connected As Boolean
Private arrData() As String
Private strHeaders As String
Private MD5Hasher As New MD5
Private ArrFilesToDownload() As String
Private needsUpdate As Boolean
Private downloadArr() As String
Private currentFile As Long
Private downloadPart As Long
Private runCopy As Boolean
Private SizeOfFile As Single
Private BytesTransfered As Single

Sub getFileList()
    gettingFileList = True
    lblCurrentFile.Caption = "Connecting..."
    ReDim arrData(1)
    socketUpdate.Connect
    
End Sub

Private Sub Form_Load()
    On Error Resume Next
    MkDir (App.Path & "\data")
    MkDir (App.Path & "\data\audio")
    MkDir (App.Path & "\data\audio\music")
    MkDir (App.Path & "\data\audio\sound")
    MkDir (App.Path & "\data\bmp")
    MkDir (App.Path & "\maps")
    Dim freeFile1 As Long
    freeFile1 = FreeFile
'    Open App.Path & "\setup.txt" For Output As #freeFile1
'    Close #freeFile1
    runCopy = False
    getFileList
End Sub

Private Sub socketUpdate_Close()
If updating = True Then
    shpCurFile.Width = 3000
    socketUpdate.Close
    savefile
Else
    socketUpdate.Close
End If
End Sub

Private Sub socketUpdate_Connect()
    If gettingFileList = True Then
        lblCurrentFile.Caption = "Connected"
        strHeaders = "GET http://www.root12.co.uk/updater/filelist.txt HTTP/1.1" & vbCrLf
        strHeaders = strHeaders & "Range: bytes0-" & vbCrLf
        strHeaders = strHeaders & "Host: root12.co.uk" & vbCrLf
        strHeaders = strHeaders & "Accept: */*" & vbCrLf
        strHeaders = strHeaders & "User-Agent: AfterDarkness Updater V0.1" & vbCrLf
        strHeaders = strHeaders & "Connection: close" & vbCrLf
        socketUpdate.SendData (strHeaders & vbCrLf & vbCrLf)
        DoEvents
    ElseIf updating = True Then
        
        
        Call socketUpdate.SendData(getHeaders(ArrFilesToDownload(currentFile)))
        DoEvents
    Else
        'MsgBox "done"
        If runCopy = True Then
            lblCurrentFile.Caption = "Complete"
            socketUpdate.Close
            
            Call Shell(App.Path & "\copy.exe", vbNormalFocus)
            End
            DoEvents
        Else
            lblCurrentFile.Caption = "Complete"
            socketUpdate.Close
            'Load frmSplash
            'frmSplash.Show
            Unload Me
        End If
    Exit Sub
    End If
End Sub

Private Sub socketUpdate_DataArrival(ByVal bytesTotal As Long)
Dim tempdata As String

    If gettingFileList = True Then
        socketUpdate.GetData tempdata
        arrData(0) = tempdata
        Call checkLocalFiles
        DoEvents
    ElseIf updating = True Then
        lblCurrentFile.Caption = "Getting: " & ArrFilesToDownload(currentFile)
        ReDim Preserve downloadArr(downloadPart)
        socketUpdate.GetData tempdata
        Debug.Print tempdata
        downloadArr(downloadPart) = tempdata
        'Debug.Print tempdata
        downloadPart = downloadPart + 1
        getSizeOfFile
        workoutStatusBar bytesTotal
    End If
    
    
End Sub

Sub workoutStatusBar(ByVal Bytes As Long)
On Error Resume Next
    BytesTransfered = BytesTransfered + Bytes
    Me.shpCurFile.Width = (BytesTransfered / SizeOfFile) * 3000
End Sub


Private Sub getSizeOfFile()
On Error Resume Next
    Dim splitarr() As String
    splitarr = Split(downloadArr(0), vbCrLf)
    SizeOfFile = Val(Right(splitarr(6), Len(splitarr(6)) - 16))
End Sub

Private Sub checkLocalFiles()
Dim i As Long
Dim a As Long
Dim counter As Long
    Dim tempArr() As String
    Dim splitarr() As String
    counter = 0
    ReDim ArrFilesToDownload(counter)
    lblCurrentFile.Caption = "Checking Local Files"
    'MsgBox arrData(0)
    tempArr = Split(arrData(0), vbCrLf)
    needsUpdate = False
    For i = 12 To UBound(tempArr) Step 1
        splitarr = Split(tempArr(i), vbLf)
        On Error Resume Next
        For a = 0 To UBound(splitarr) Step 2
        Debug.Print checkMD5hash(splitarr(a), splitarr(a + 1))
        If checkMD5hash(splitarr(a), splitarr(a + 1)) = False Then
            ReDim Preserve ArrFilesToDownload(counter)
            needsUpdate = True
            If splitarr(a) = "adnew.exe" Then
                ArrFilesToDownload(counter) = splitarr(a)
                ReDim Preserve ArrFilesToDownload(counter + 1)
                ArrFilesToDownload(counter + 1) = "copy.exe"
                runCopy = True
                counter = counter + 1
            Else
                ArrFilesToDownload(counter) = splitarr(a)
            End If
            
            counter = counter + 1
        End If
        Next a
    Next i
    
    If needsUpdate = True Then
        'old files need replacing
        gettingFileList = False
        'MsgBox "bad file"
        startUpdate
    Else
        ' all done
        lblCurrentFile.Caption = "Complete"
        socketUpdate.Close
        'Load frmSplash
        'frmSplash.Show
        Unload Me
    End If
End Sub

Public Function checkMD5hash(ByVal filename As String, ByVal hash As String) As Boolean
'If filename = "adnew.exe" Then
    'filename = "adste.exe"
'End If
    Debug.Print MD5Hasher.GetCheckSumFromFile(App.Path & "\" & filename)
    Debug.Print hash
    If MD5Hasher.GetCheckSumFromFile(App.Path & "\" & filename) <> hash Then
        checkMD5hash = False
    Else
        checkMD5hash = True
    End If
End Function

Private Sub startUpdate()
    updating = True
    ReDim downloadArr(0)
    socketUpdate.Close
    socketUpdate.Connect
    currentFile = 0
    downloadPart = 0
    DoEvents
End Sub

Private Function getHeaders(ByVal filename As String) As String
        getHeaders = "GET http://www.root12.co.uk/updater/" & filename & " HTTP/1.1" & vbCrLf
        getHeaders = getHeaders & "Range: bytes0-" & vbCrLf
        getHeaders = getHeaders & "Host: root12.co.uk" & vbCrLf
        getHeaders = getHeaders & "Accept: */*" & vbCrLf
        getHeaders = getHeaders & "User-Agent: AfterDarkness Updater V0.1" & vbCrLf
        getHeaders = getHeaders & "Connection: close" & vbCrLf & vbCrLf & vbCrLf
        downloadPart = 0
        BytesTransfered = 0
End Function


Private Sub savefile()
    
    Dim splitarr() As String
    Dim strData As String
    Dim filenum As Long
    Dim i As Long
    lblCurrentFile.Caption = "Saving File"
    splitarr = Split(downloadArr(0), vbCrLf)
    For i = 12 To UBound(splitarr) - 1 Step 1
        strData = strData & splitarr(i) & vbCrLf
    Next i
    strData = strData & splitarr(UBound(splitarr))
'    For i = 1 To UBound(downloadArr) Step 1
'        strData = strData & downloadArr(i)
'
'    Next i
    filenum = FreeFile
    On Error Resume Next
    Kill (App.Path & "\" & ArrFilesToDownload(currentFile))
    'If Dir(App.Path & "\" & ArrFilesToDownload(currentFile)) Then Kill (App.Path & "\" & ArrFilesToDownload(currentFile))
    DoEvents
    Open App.Path & "\" & ArrFilesToDownload(currentFile) For Binary As #filenum
        Put #filenum, , strData
    Close #filenum
    
    Open App.Path & "\" & ArrFilesToDownload(currentFile) For Binary As #filenum
        DoEvents
        Seek #filenum, Len(strData) + 1
        
        For i = 1 To UBound(downloadArr()) Step 1
            Me.shpCurFile.Width = Int(i / UBound(downloadArr) * 3000)
            Put #filenum, , downloadArr(i)
            DoEvents
        Next i
    Close #filenum
    
    
    If currentFile >= UBound(ArrFilesToDownload) Then
        updating = False
    End If
    currentFile = currentFile + 1
    ReDim downloadArr(0)
    socketUpdate.Connect
    Me.shpOverall.Width = ((currentFile - 1) / UBound(ArrFilesToDownload)) * 3000
End Sub
