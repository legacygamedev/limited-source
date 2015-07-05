Attribute VB_Name = "SysInfo"
'=========================================================================================
'  Coded By: Deepesh Agarwal
'  Published Date: 29/09/2003
'  WebSite: http://www.deepeshagarwal.tk
'  E-mail: agarwal_deepesh@indiatimes.com
'  Collection compiled from various sources
'=========================================================================================
 Type SYSTEM_INFO
            dwOemID As Long
            dwPageSize As Long
            lpMinimumApplicationAddress As Long
            lpMaximumApplicationAddress As Long
            dwActiveProcessorMask As Long
            dwNumberOrfProcessors As Long
            dwProcessorType As Long
            dwAllocationGranularity As Long
            dwReserved As Long
      End Type
      Type OSVERSIONINFO
            dwOSVersionInfoSize As Long
            dwMajorVersion As Long
            dwMinorVersion As Long
            dwBuildNumber As Long
            dwPlatformId As Long
            szCSDVersion As String * 128
      End Type
      Type MEMORYSTATUS
            dwLength As Long
            dwMemoryLoad As Long
            dwTotalPhys As Long
            dwAvailPhys As Long
            dwTotalPageFile As Long
            dwAvailPageFile As Long
            dwTotalVirtual As Long
            dwAvailVirtual As Long
      End Type

      Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
         (LpVersionInformation As OSVERSIONINFO) As Long
      Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As _
         MEMORYSTATUS)
      Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As _
         SYSTEM_INFO)

      Public Const PROCESSOR_INTEL_386 = 386
      Public Const PROCESSOR_INTEL_486 = 486
      Public Const PROCESSOR_INTEL_PENTIUM = 586
      Public Const PROCESSOR_MIPS_R4000 = 4000
      Public Const PROCESSOR_ALPHA_21064 = 21064
Private Const MAX_COMPUTERNAME_LENGTH As Long = 31
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'Declarations For Display Info
Const ENUM_CURRENT_SETTINGS As Long = -1&
Const CCDEVICENAME = 32
Const CCFORMNAME = 32
Private Type DEVMODE
    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
'End Of Declaration For Display Info

'Declaration For Getting GUI Resources Info
Private Const GR_GDIOBJECTS = 0
Private Const GR_USEROBJECTS = 1
Const GFSR_SYSTEMRESOURCES = 0
Const GFSR_GDIRESOURCES = 1
Const GFSR_USERRESOURCES = 2
Private Declare Function GetGuiResources Lib "user32.dll" (ByVal hProcess As Long, ByVal uiFlags As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
'End of Getting GUI Resources Info

'Declaration For Processor Feature
Private Const PF_FLOATING_POINT_PRECISION_ERRATA = 0
Private Const PF_FLOATING_POINT_EMULATED = 1
Private Const PF_COMPARE_EXCHANGE_DOUBLE = 2
Private Const PF_MMX_INSTRUCTIONS_AVAILABLE = 3
Private Const PF_XMMI_INSTRUCTIONS_AVAILABLE = 6
Private Const PF_3DNOW_INSTRUCTIONS_AVAILABLE = 7
Private Const PF_RDTSC_INSTRUCTION_AVAILABLE = 8
Private Const PF_PAE_ENABLED = 9
Private Declare Function IsProcessorFeaturePresent Lib "kernel32.dll" (ByVal ProcessorFeature As Long) As Long
Dim CPUInfo As String
'End of Declaration For processor feature
'Declaration For Environment settings
Private Declare Function GetEnvironmentStrings Lib "kernel32" Alias "GetEnvironmentStringsA" () As Long
Private Declare Function FreeEnvironmentStrings Lib "kernel32" Alias "FreeEnvironmentStringsA" (ByVal lpsz As String) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
'End of Declaration For Environment settings

'Declerations For Enumerating Running Process
Const TH32CS_SNAPHEAPLIST = &H1
Const TH32CS_SNAPPROCESS = &H2
Const TH32CS_SNAPTHREAD = &H4
Const TH32CS_SNAPMODULE = &H8
Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Const TH32CS_INHERIT = &H80000000
Const MAX_PATH As Integer = 260
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)
'End Declerations For Enumerating Running Process

'Decleration For Checking is administartor
Private Const ANYSIZE_ARRAY = 20 'Fixed at this size for comfort. Could be bigger or made dynamic.

'Security APIs
Private Const TokenUser = 1
Private Const TokenGroups = 2
Private Const TokenPrivileges = 3
Private Const TokenOwner = 4
Private Const TokenPrimaryGroup = 5
Private Const TokenDefaultDacl = 6
Private Const TokenSource = 7
Private Const TokenType = 8
Private Const TokenImpersonationLevel = 9
Private Const TokenStatistics = 10

'Token Specific Access Rights
Private Const TOKEN_ASSIGN_PRIMARY = &H1
Private Const TOKEN_DUPLICATE = &H2
Private Const TOKEN_IMPERSONATE = &H4
Private Const TOKEN_QUERY = &H8
Private Const TOKEN_QUERY_SOURCE = &H10
Private Const TOKEN_ADJUST_PRIVILEGES = &H20
Private Const TOKEN_ADJUST_GROUPS = &H40
Private Const TOKEN_ADJUST_DEFAULT = &H80
 
' NT well-known SIDs
Private Const SECURITY_DIALUP_RID = &H1
Private Const SECURITY_NETWORK_RID = &H2
Private Const SECURITY_BATCH_RID = &H3
Private Const SECURITY_INTERACTIVE_RID = &H4
Private Const SECURITY_SERVICE_RID = &H6
Private Const SECURITY_ANONYMOUS_LOGON_RID = &H7
Private Const SECURITY_LOGON_IDS_RID = &H5
Private Const SECURITY_LOCAL_SYSTEM_RID = &H12
Private Const SECURITY_NT_NON_UNIQUE = &H15
Private Const SECURITY_BUILTIN_DOMAIN_RID = &H20

' Well-known domain relative sub-authority values (RIDs)
Private Const DOMAIN_ALIAS_RID_ADMINS = &H220
Private Const DOMAIN_ALIAS_RID_USERS = &H221
Private Const DOMAIN_ALIAS_RID_GUESTS = &H222
Private Const DOMAIN_ALIAS_RID_POWER_USERS = &H223
Private Const DOMAIN_ALIAS_RID_ACCOUNT_OPS = &H224
Private Const DOMAIN_ALIAS_RID_SYSTEM_OPS = &H225
Private Const DOMAIN_ALIAS_RID_PRINT_OPS = &H226
Private Const DOMAIN_ALIAS_RID_BACKUP_OPS = &H227
Private Const DOMAIN_ALIAS_RID_REPLICATOR = &H228

Private Const SECURITY_NT_AUTHORITY = &H5

Type SID_AND_ATTRIBUTES
    Sid As Long
    Attributes As Long
End Type

Type TOKEN_GROUPS
    GroupCount As Long
    Groups(ANYSIZE_ARRAY) As SID_AND_ATTRIBUTES
End Type

Type SID_IDENTIFIER_AUTHORITY
    Value(0 To 5) As Byte
End Type

'Declare Function GetCurrentProcess Lib "Kernel32" () As Long

Declare Function GetCurrentThread Lib "kernel32" () As Long

Declare Function OpenProcessToken Lib "Advapi32" ( _
    ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, _
    TokenHandle As Long) As Long

Declare Function OpenThreadToken Lib "Advapi32" ( _
    ByVal ThreadHandle As Long, ByVal DesiredAccess As Long, _
    ByVal OpenAsSelf As Long, TokenHandle As Long) As Long

Declare Function GetTokenInformation Lib "Advapi32" ( _
    ByVal TokenHandle As Long, TokenInformationClass As Integer, _
    TokenInformation As Any, ByVal TokenInformationLength As Long, _
    ReturnLength As Long) As Long

Declare Function AllocateAndInitializeSid Lib "Advapi32" ( _
    pIdentifierAuthority As SID_IDENTIFIER_AUTHORITY, _
    ByVal nSubAuthorityCount As Byte, ByVal nSubAuthority0 As Long, _
    ByVal nSubAuthority1 As Long, ByVal nSubAuthority2 As Long, _
    ByVal nSubAuthority3 As Long, ByVal nSubAuthority4 As Long, _
    ByVal nSubAuthority5 As Long, ByVal nSubAuthority6 As Long, _
    ByVal nSubAuthority7 As Long, lpPSid As Long) As Long

Declare Function RtlMoveMemory Lib "kernel32" ( _
    Dest As Any, Source As Any, ByVal lSize As Long) As Long

Declare Function IsValidSid Lib "Advapi32" (ByVal pSid As Long) As Long

Declare Function EqualSid Lib "Advapi32" (pSid1 As Any, pSid2 As Any) As Long

Declare Sub FreeSid Lib "Advapi32" (pSid As Any)

'Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long
'End checking IsAdmin

'Declaration For Getting Drives Info
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
'End of Declaration For Getting Drives Info

Public Function GetSysInfo() As String
    Dim sysinfo As SYSTEM_INFO
    Dim msg As Variant
    'Get system name
    Dim dwLen As Long
    Dim strString As String
    'Create a buffer
    dwLen = MAX_COMPUTERNAME_LENGTH + 1
    strString = String(dwLen, "X")
    'Get the computer name
    GetComputerName strString, dwLen
    'get only the actual data
    strString = Left(strString, dwLen)
    msg = msg & String(40, "#") & vbCrLf
    msg = msg & String(15, " ") & "Debug Report" & vbCrLf
    msg = msg & String(40, "#") & vbCrLf
    msg = msg & "Log Date : " & "(" & Date & " - " & Time() & ")" & vbCrLf
    msg = msg & "Debug Log For System : " & strString & vbCrLf
    msg = msg & "ExeName : " & App.EXEName & vbCrLf
    msg = msg & "ExeVersion : " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf
    msg = msg & "ExePath : " & App.Path & vbCrLf
    'Format & Append it into a string variable
    ' Get operating system and version.
    msg = msg & "OS Platform : " & GetOS
    'Check if administartor
    If IsAdmin Then
        msg = msg & "Is User Administrator : True" & vbCrLf
    Else
        msg = msg & "Is User Administrator : False" & vbCrLf
    End If

    ' Get CPU type and operating mode.
    GetSystemInfo sysinfo
    msg = msg & "CPU: "
    Select Case sysinfo.dwProcessorType
        Case PROCESSOR_INTEL_386
            msg = msg & "Intel 386" '& vbCrLf
        Case PROCESSOR_INTEL_486
            msg = msg & "Intel 486" '& vbCrLf
        Case PROCESSOR_INTEL_PENTIUM
            msg = msg & "Intel Pentium" '& vbCrLf
        Case PROCESSOR_MIPS_R4000
            msg = msg & "MIPS R4000" '& vbCrLf
        Case PROCESSOR_ALPHA_21064
            msg = msg & "DEC Alpha 21064" '& vbCrLf
        Case Else
            msg = msg & "(unknown)" '& vbCrLf

    End Select
    msg = msg & vbCrLf
    'get additional cpu info
    msg = msg & GetCPUFeature & vbCrLf
    ' Get free memory.
    Dim memsts As MEMORYSTATUS
    Dim memory As Long
    GlobalMemoryStatus memsts
    memory = memsts.dwTotalPhys
    msg = msg & "Total Physical Memory: "
    msg = msg & FormatFileSize(memory) & vbCrLf

    memory& = memsts.dwAvailPhys
    msg = msg & "Available Physical Memory: "
    msg = msg & FormatFileSize(memory) & vbCrLf

    memory& = memsts.dwTotalVirtual
    msg = msg & "Total Virtual Memory: "
    msg = msg & FormatFileSize(memory) & vbCrLf

    memory& = memsts.dwAvailVirtual
    msg = msg & "Available Virtual Memory: "
    msg = msg & FormatFileSize(memory) & vbCrLf '& vbCrLf
    
    memory& = memsts.dwAvailPageFile
    msg = msg & "Available Page File : "
    msg = msg & FormatFileSize(memory) & vbCrLf
    
    memory& = memsts.dwMemoryLoad
    msg = msg & "Total Memory Load : "
    msg = msg & memory & vbCrLf & vbCrLf
    
    
   'get Windows resourses info
    msg = msg & GetGUIInfo
    'Get System Display Related Info
    msg = msg & GetDisplayInfo()
    'Get Drives Info
    msg = msg & GetDriveInfo
    'get running process info
    msg = msg & String(30, "+") & vbCrLf
    msg = msg & "Currently Running Processes : " & vbCrLf
    msg = msg & String(30, "+") & vbCrLf
    msg = msg & GetProcessInfo '& vbCrLf
    msg = msg & String(30, "+") & vbCrLf
    msg = msg & "End of Processes Info" & vbCrLf
    msg = msg & String(30, "+") & vbCrLf & vbCrLf
    'Get Enviournment Settings info
    msg = msg & String(30, "+") & vbCrLf
    msg = msg & "Dumping Environment Strings" & vbCrLf
    msg = msg & String(30, "+") & vbCrLf
    msg = msg & Trim(GetEnvString)
    msg = msg & String(30, "+") & vbCrLf
    msg = msg & "End of Dumping Environment Strings" & vbCrLf
    msg = msg & String(30, "+") & vbCrLf
    'return information to calling function
    GetSysInfo = msg
End Function


Public Function GetDisplayInfo() As String
    Dim DispInfo As DEVMODE
    Dim SDispRet As String
    Call EnumDisplaySettings(0&, ENUM_CURRENT_SETTINGS, DispInfo)
    SDispRet = "Current screen width: " & DispInfo.dmPelsWidth & " Pixels" & vbCrLf
    SDispRet = SDispRet & "Current screen height: " & DispInfo.dmPelsHeight & " Pixels" & vbCrLf
    SDispRet = SDispRet & "Current color depth: " & DispInfo.dmBitsPerPel & " Bits/Pixel" & vbCrLf
    SDispRet = SDispRet & "Display Driver : " & Trim(DispInfo.dmDeviceName) & vbCrLf
    'Return gathered info to calling function
    GetDisplayInfo = SDispRet
End Function

Public Function GetGUIInfo() As String
    Dim GUIInfoRet As String
    'Get GUI info for our app
    GUIInfoRet = GUIInfoRet & "GDI objects used by this app : " & GetGuiResources(GetCurrentProcess, GR_GDIOBJECTS) & vbCrLf
    GUIInfoRet = GUIInfoRet & "User objects used by this app : " & GetGuiResources(GetCurrentProcess, GR_USEROBJECTS) & vbCrLf
    'Return gathered info to calling function
    GetGUIInfo = GUIInfoRet
End Function

Public Function GetCPUFeature() As String
    ShowFeature PF_FLOATING_POINT_PRECISION_ERRATA, "Floating point error"
    ShowFeature PF_FLOATING_POINT_EMULATED, "Floating-point operations emulated"
    ShowFeature PF_COMPARE_EXCHANGE_DOUBLE, "Compare and exchange double operation available"
    ShowFeature PF_MMX_INSTRUCTIONS_AVAILABLE, "MMX instructions available"
    ShowFeature PF_XMMI_INSTRUCTIONS_AVAILABLE, "XMMI instructions available"
    ShowFeature PF_3DNOW_INSTRUCTIONS_AVAILABLE, "3D-Now instructions available"
    ShowFeature PF_RDTSC_INSTRUCTION_AVAILABLE, "RDTSC instructions available"
    ShowFeature PF_PAE_ENABLED, "Processor is PAE-enabled"
    'Return gathered info to calling function
    GetCPUFeature = CPUInfo
End Function
Private Sub ShowFeature(lIndex As Long, Description As String)

    If IsProcessorFeaturePresent(lIndex) = 0 Then
        CPUInfo = CPUInfo & Description + " : False" & vbCrLf
    Else
        CPUInfo = CPUInfo & Description + " : True" & vbCrLf
    End If
End Sub



Public Function GetEnvString() As String

    Dim lngRet As Long, strDest As String, lLen As Long
    Dim sEnvRet As String
    'retrieve the initial pointer to the environment strings
    lngRet = GetEnvironmentStrings
    Do
        'get the length of the following string
        lLen = lstrlen(lngRet)
        'if the length equals 0, we've reached the end
        If lLen = 0 Then Exit Do
        'create a buffer string
        strDest = Space$(lLen)
        'copy the text from the environment block
        CopyMemory ByVal strDest, ByVal lngRet, lLen
        sEnvRet = sEnvRet & strDest & vbCrLf

        'move the pointer
        lngRet = lngRet + lstrlen(lngRet) + 1
    Loop
    'GetEnvString = sEnvRet
    'clean up
    FreeEnvironmentStrings lngRet
    'Return gathered info to calling function
    GetEnvString = sEnvRet
End Function

Public Function GetProcessInfo() As String
    Dim hSnapShot As Long, uProcess As PROCESSENTRY32
    Dim sRetProcInfo As String
    'Takes a snapshot of the processes and the heaps, modules, and threads used by the processes
    hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)
    'set the length of our ProcessEntry-type
    uProcess.dwSize = Len(uProcess)
    'Retrieve information about the first process encountered in our system snapshot
    r = Process32First(hSnapShot, uProcess)

    Do While r
        sRetProcInfo = sRetProcInfo & Left$(uProcess.szExeFile, IIf(InStr(1, uProcess.szExeFile, Chr$(0)) > 0, InStr(1, uProcess.szExeFile, Chr$(0)) - 1, 0)) & vbCrLf
        'Retrieve information about the next process recorded in our system snapshot
        r = Process32Next(hSnapShot, uProcess)
    Loop
    'close our snapshot handle
    CloseHandle hSnapShot
    'Return gathered info to calling function
    GetProcessInfo = sRetProcInfo
End Function
Public Function IsAdmin() As Boolean
    'By Anderson Mesquita
    Dim hProcessToken       As Long
    Dim BufferSize          As Long
    Dim psidAdmin           As Long
    Dim lResult             As Long
    Dim X                   As Integer
    Dim tpTokens            As TOKEN_GROUPS
    Dim tpSidAuth           As SID_IDENTIFIER_AUTHORITY

    IsAdmin = False
    tpSidAuth.Value(5) = SECURITY_NT_AUTHORITY

    ' Obtain current process token
    If Not OpenThreadToken(GetCurrentThread(), TOKEN_QUERY, True, hProcessToken) Then
        Call OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, hProcessToken)
    End If
    If hProcessToken Then

        ' Deternine the buffer size required
        Call GetTokenInformation(hProcessToken, ByVal TokenGroups, 0, 0, BufferSize) ' Determine required buffer size
        If BufferSize Then
            ReDim InfoBuffer((BufferSize \ 4) - 1) As Long

            ' Retrieve your token information
            lResult = GetTokenInformation(hProcessToken, ByVal TokenGroups, InfoBuffer(0), BufferSize, BufferSize)
            If lResult <> 1 Then Exit Function

            ' Move it from memory into the token structure
            Call RtlMoveMemory(tpTokens, InfoBuffer(0), Len(tpTokens))

            ' Retreive the admins sid pointer
            lResult = AllocateAndInitializeSid(tpSidAuth, 2, SECURITY_BUILTIN_DOMAIN_RID, _
                DOMAIN_ALIAS_RID_ADMINS, 0, 0, 0, 0, 0, 0, psidAdmin)
            If lResult <> 1 Then Exit Function
            If IsValidSid(psidAdmin) Then
                For X = 0 To tpTokens.GroupCount

                    ' Run through your token sid pointers
                    If IsValidSid(tpTokens.Groups(X).Sid) Then

                        ' Test for a match between the admin sid equalling your sid's
                        If EqualSid(ByVal tpTokens.Groups(X).Sid, ByVal psidAdmin) Then
                            IsAdmin = True
                            Exit For
                        End If
                    End If
                Next
            End If
            If psidAdmin Then Call FreeSid(psidAdmin)
        End If
        Call CloseHandle(hProcessToken)
    End If
End Function


Public Function GetOS() As String
    Dim verinfo As OSVERSIONINFO
    Dim sRetOS As String
    Dim build As String, ver_major As String, ver_minor As String
    Dim ret As Long
    verinfo.dwOSVersionInfoSize = Len(verinfo)
    ret = GetVersionEx(verinfo)
    If ret = 0 Then
        MsgBox "Error Getting Version Information"
        End
    End If
    Select Case verinfo.dwPlatformId
        Case 0
            sRetOS = "Windows 32s "
        Case 1
            sRetOS = "Windows 95 "
        Case 2
            sRetOS = "Windows NT "
    End Select

    ver_major = verinfo.dwMajorVersion
    ver_minor = verinfo.dwMinorVersion
    build = verinfo.dwBuildNumber
    sRetOS = sRetOS & ver_major & "." & ver_minor
    sRetOS = sRetOS & " (Build " & build & ")" & vbCrLf & vbCrLf
    'Return gathered info to calling function
    GetOS = sRetOS
End Function

Public Function GetDriveInfo() As String
    On Error Resume Next
    Dim fso As New FileSystemObject, drv As Drive
    Dim strAllDrives As String
    Dim strTmp As String, drvinfo As String
    DoEvents
    'drvinfo = String(30, "+") & vbCrLf
    'drvinfo = drvinfo & "Get All Drive Info" & vbCrLf
    drvinfo = drvinfo & String(35, "+") & vbCrLf
    drvinfo = drvinfo & "Drv Type FSys VName TSize Avail/Quota " & vbCrLf
    drvinfo = drvinfo & String(35, "+") & vbCrLf
    strAllDrives = fGetDrives
    If strAllDrives <> "" Then
        Do
            DoEvents
            strTmp = Mid$(strAllDrives, 1, InStr(strAllDrives, vbNullChar) - 1)
            strAllDrives = Mid$(strAllDrives, InStr(strAllDrives, vbNullChar) + 1)
            Set drv = fso.GetDrive(fso.GetDriveName(Left$(strTmp, 2)))
            If drv.IsReady = True Then
                drvinfo = drvinfo & drv.DriveLetter & " "
                Select Case drv.DriveType
                    Case 0: dtype = "Unknown"
                    Case 1: dtype = "Removable"
                    Case 2: dtype = "Fixed"
                    Case 3: dtype = "Network"
                    Case 4: dtype = "CD-ROM"
                    Case 5: dtype = "RAM Disk"
                End Select
                drvinfo = drvinfo & dtype & " "
                drvinfo = drvinfo & drv.FileSystem & " "
                drvinfo = drvinfo & drv.VolumeName & " "
                drvinfo = drvinfo & FormatFileSize(drv.TotalSize) & " "
                drvinfo = drvinfo & FormatFileSize(drv.AvailableSpace)

            Else
                drvinfo = drvinfo & drv.DriveLetter & " "
                Select Case drv.DriveType
                    Case 0: dtype = "Unknown"
                    Case 1: dtype = "Removable"
                    Case 2: dtype = "Fixed"
                    Case 3: dtype = "Network"
                    Case 4: dtype = "CD-ROM"
                    Case 5: dtype = "RAM Disk"
                End Select
                drvinfo = drvinfo & dtype & "  "
                drvinfo = drvinfo & " --Drive Not Ready--"
            End If

            sRetDrvInfo = sRetDrvInfo & vbCrLf & drvinfo
            drvinfo = ""
            Set drv = Nothing
        Loop While strAllDrives <> ""
        sRetDrvInfo = sRetDrvInfo & String(35, "+") & vbCrLf
        sRetDrvInfo = sRetDrvInfo & "End of Drives Info" & vbCrLf
        sRetDrvInfo = sRetDrvInfo & String(35, "+") & vbCrLf

        'Return gathered info to calling function
        GetDriveInfo = sRetDrvInfo & vbCrLf & vbCrLf
    End If
End Function
Private Function fGetDrives() As String
    'Returns all mapped drives
    Dim lngRet As Long
    Dim strDrives As String * 255
    Dim lngTmp As Long
    lngTmp = Len(strDrives)
    lngRet = GetLogicalDriveStrings(lngTmp, strDrives)
    fGetDrives = Left(strDrives, lngRet)
End Function

Public Function FormatFileSize(ByVal Size As Double) As String
    Dim sRet As String
    Const KB& = 1024
    Const MB& = KB * KB
    ' Return size of file in kilobytes.
    If Size < KB Then
        sRet = Format(Size, "#,##0") & " bytes"
    Else

        Select Case Size / 1024
            Case Is < 10
                sRet = Format(Size / KB, "0.00") & "KB"
            Case Is < 100
                sRet = Format(Size / KB, "0.0") & "KB"
            Case Is < 1000
                sRet = Format(Size / KB, "0") & "KB"
            Case Is < 10000
                sRet = Format(Size / MB, "0.00") & "MB"
            Case Is < 100000
                sRet = Format(Size / MB, "0.0") & "MB"
            Case Is < 1000000
                sRet = Format(Size / MB, "0") & "MB"
            Case Is < 10000000
                sRet = Format(Size / MB / KB, "0.00") & "GB"
            Case Else
                sRet = "Error"
        End Select
        sRet = sRet '& " (" & Format(Size, "#,##0") & " bytes)"
    End If
    FormatFileSize = sRet
End Function

