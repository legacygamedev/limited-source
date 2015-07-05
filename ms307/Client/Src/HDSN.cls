VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HDSN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' Antonio Giuliana, 2001-2003

' Costanti per l'individuazione della versione di OS
Private Const VER_PLATFORM_WIN32S = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

' Costanti per la comunicazione con il driver IDE
Private Const DFP_RECEIVE_DRIVE_DATA = &H7C088

' Costanti per la CreateFile
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3
Private Const CREATE_NEW = 1

' Enumerazione dei comandi per la CmnGetHDData
Private Enum HDINFO
    HD_MODEL_NUMBER
    HD_SERIAL_NUMBER
    HD_FIRMWARE_REVISION
End Enum

' Struttura per l'individuazione della versione di OS
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

' Struttura per il campo irDriveRegs della struttura SENDCMDINPARAMS
Private Type IDEREGS
    bFeaturesReg As Byte
    bSectorCountReg As Byte
    bSectorNumberReg As Byte
    bCylLowReg As Byte
    bCylHighReg As Byte
    bDriveHeadReg As Byte
    bCommandReg As Byte
    bReserved As Byte
End Type

' Struttura per l'I/O dei comandi al driver IDE
Private Type SENDCMDINPARAMS
    cBufferSize As Long
    irDriveRegs As IDEREGS
    bDriveNumber As Byte
    bReserved(1 To 3) As Byte
    dwReserved(1 To 4) As Long
End Type

' Struttura per il campo DStatus della struttura SENDCMDOUTPARAMS
Private Type DRIVERSTATUS
    bDriveError As Byte
    bIDEStatus As Byte
    bReserved(1 To 2) As Byte
    dwReserved(1 To 2) As Long
End Type

' Struttura per l'I/O dei comandi al driver IDE
Private Type SENDCMDOUTPARAMS
    cBufferSize As Long
    DStatus As DRIVERSTATUS     ' ovvero DriverStatus
    bBuffer(1 To 512) As Byte
End Type

' Per ottenere la versione del SO
Private Declare Function GetVersionEx _
    Lib "kernel32" Alias "GetVersionExA" _
    (lpVersionInformation As OSVERSIONINFO) As Long

' Per ottenere un handle al device IDE
Private Declare Function CreateFile _
    Lib "kernel32" Alias "CreateFileA" _
    (ByVal lpFileName As String, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    ByVal lpSecurityAttributes As Long, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long) As Long

' Per chiudere l'handle del device IDE
Private Declare Function CloseHandle _
    Lib "kernel32" _
    (ByVal hObject As Long) As Long

' Per comunicare con il driver IDE
Private Declare Function DeviceIoControl _
    Lib "kernel32" _
    (ByVal hDevice As Long, _
    ByVal dwIoControlCode As Long, _
    lpInBuffer As Any, _
    ByVal nInBufferSize As Long, _
    lpOutBuffer As Any, _
    ByVal nOutBufferSize As Long, _
    lpBytesReturned As Long, _
    ByVal lpOverlapped As Long) As Long
    
' Per azzerare buffer di scambio dati
Private Declare Sub ZeroMemory _
    Lib "kernel32" Alias "RtlZeroMemory" _
    (dest As Any, _
    ByVal numBytes As Long)

' Per copiare porzioni di memoria
Private Declare Sub CopyMemory _
    Lib "kernel32" Alias "RtlMoveMemory" _
    (Destination As Any, _
    Source As Any, _
    ByVal Length As Long)

Private Declare Function GetLastError _
    Lib "kernel32" () As Long

Private mvarCurrentDrive As Byte    ' Drive corrente
Private mvarPlatform As String      ' Piattaforma usata

Public Property Get Copyright() As String
    
    ' Copyright
    Copyright = "HDSN Vrs. 1.00, (C) Antonio Giuliana, 2001-2003"
    
End Property

' Metodo GetModelNumber
Public Function GetModelNumber() As String
    
    ' Ottiene il ModelNumber
    GetModelNumber = CmnGetHDData(HD_MODEL_NUMBER)
    
End Function

' Metodo GetSerialNumber
Public Function GetSerialNumber() As String
    
    ' Ottiene il SerialNumber
    GetSerialNumber = CmnGetHDData(HD_SERIAL_NUMBER)
    
End Function

' Metodo GetFirmwareRevision
Public Function GetFirmwareRevision() As String
    
    ' Ottiene la FirmwareRevision
    GetFirmwareRevision = CmnGetHDData(HD_FIRMWARE_REVISION)
    
End Function

' Proprieta' CurrentDrive
Public Property Let CurrentDrive(ByVal vData As Byte)
    
    ' Controllo numero di drive fisico IDE
    If vData < 0 Or vData > 3 Then
        Err.Raise 10000, , "Illegal drive number"   ' IDE drive 0..3
    End If
    
    ' Nuovo drive da considerare
    mvarCurrentDrive = vData

End Property

' Proprieta' CurrentDrive
Public Property Get CurrentDrive() As Byte
    
    ' Restituisce drive fisico corrente (IDE 0..3)
    CurrentDrive = mvarCurrentDrive

End Property

' Proprieta' Platform
Public Property Get Platform() As String
    
    ' Restituisce tipo OS
    Platform = mvarPlatform

End Property

Private Sub Class_Initialize()

    ' Individuazione del tipo di OS
    Dim OS As OSVERSIONINFO
        
    OS.dwOSVersionInfoSize = Len(OS)
    Call GetVersionEx(OS)
    mvarPlatform = "Unk"
    Select Case OS.dwPlatformId
        Case Is = VER_PLATFORM_WIN32S
            mvarPlatform = "32S"                ' Win32S
        Case Is = VER_PLATFORM_WIN32_WINDOWS
            If OS.dwMinorVersion = 0 Then
                mvarPlatform = "W95"            ' Win 95
            Else
                mvarPlatform = "W98"            ' Win 98
            End If
        Case Is = VER_PLATFORM_WIN32_NT
            mvarPlatform = "WNT"                ' Win NT/2000
    End Select

End Sub

Private Function CmnGetHDData(hdi As HDINFO) As String

    ' Rilevazione proprieta' IDE
    
    Dim bin As SENDCMDINPARAMS
    Dim bout As SENDCMDOUTPARAMS
    Dim hdh As Long
    Dim br As Long
    Dim ix As Long
    Dim hddfr As Long
    Dim hddln As Long
    Dim s As String
    
    Select Case hdi             ' Selezione tipo caratteristica richiesta
        Case HD_MODEL_NUMBER
            hddfr = 55          ' Posizione nel buffer del ModelNumber
            hddln = 40          ' Lunghezza nel buffer del ModelNumber
        Case HD_SERIAL_NUMBER
            hddfr = 21          ' Posizione nel buffer del SerialNumber
            hddln = 20          ' Lunghezza nel buffer del SerialNumber
        Case HD_FIRMWARE_REVISION
            hddfr = 47          ' Posizione nel buffer del FirmwareRevision
            hddln = 8           ' Lunghezza nel buffer del FirmwareRevision
        Case Else
            Err.Raise 10001, "Illegal HD Data type" ' Altre informazioni non disponibili (Evoluzione futura)
    End Select
    
    Select Case mvarPlatform
        Case "WNT"
            ' Per Win NT/2000 apertura handle al drive fisico
            hdh = CreateFile("\\.\PhysicalDrive" & mvarCurrentDrive, _
                GENERIC_READ + GENERIC_WRITE, FILE_SHARE_READ + FILE_SHARE_WRITE, _
                0, OPEN_EXISTING, 0, 0)
        Case "W95", "W98"
            ' Per Win 9X apertura handle al driver SMART
            ' (in \WINDOWS\SYSTEM da spostare in \WINDOWS\SYSTEM\IOSUBSYS)
            ' che comunica con il driver IDE
            hdh = CreateFile("\\.\Smartvsd", _
                0, 0, 0, CREATE_NEW, 0, 0)
        Case Else
            ' Piattaforma non supportata (Win32S)
            Err.Raise 10002, , "Illegal platform (only WNT, W98 or W95)"    ' Altre piattaforme non gestite
    End Select
    ' Controllo validità handle
    If hdh = 0 Then
        Err.Raise 10003, , "Error on CreateFile"
    End If
    
    ' Azzeramento strutture per l'I/O da driver
    ZeroMemory bin, Len(bin)
    ZeroMemory bout, Len(bout)
    
    ' Preparazione parametri struttura di richiesta al driver
    With bin
        .bDriveNumber = mvarCurrentDrive
        .cBufferSize = 512
        With .irDriveRegs
            If (mvarCurrentDrive And 1) Then
                .bDriveHeadReg = &HB0
            Else
                .bDriveHeadReg = &HA0
            End If
            .bCommandReg = &HEC
            .bSectorCountReg = 1
            .bSectorNumberReg = 1
        End With
    End With
    
    ' Richiesta al driver
    DeviceIoControl hdh, DFP_RECEIVE_DRIVE_DATA, _
                    bin, Len(bin), bout, Len(bout), br, 0
    
    ' Formazione stringa di risposta
    ' da buffer di uscita
    ' L'ordine dei byte e' invertito
    s = ""
    For ix = hddfr To hddfr + hddln - 1 Step 2
        If bout.bBuffer(ix + 1) = 0 Then Exit For
        s = s & Chr(bout.bBuffer(ix + 1))
        If bout.bBuffer(ix) = 0 Then Exit For
        s = s & Chr(bout.bBuffer(ix))
    Next ix
    
    ' Chiusura handle
    CloseHandle hdh

    ' Restituzione informazione richiesta
    CmnGetHDData = Trim(s)
    
End Function

