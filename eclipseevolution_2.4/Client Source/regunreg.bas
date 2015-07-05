Attribute VB_Name = "RegUnregActiveX"
Option Explicit

'All required Win32 SDK functions to register/unregister any ActiveX component

Private Declare Function LoadLibraryRegister Lib "KERNEL32" Alias "LoadLibraryA" _
        (ByVal lpLibFileName As String) As Long

Private Declare Function FreeLibraryRegister Lib "KERNEL32" Alias "FreeLibrary" _
        (ByVal hLibModule As Long) As Long

Private Declare Function CloseHandle Lib "KERNEL32" (ByVal hObject As Long) As Long

Private Declare Function GetProcAddressRegister Lib "KERNEL32" Alias "GetProcAddress" _
        (ByVal hModule As Long, _
        ByVal lpProcName As String) As Long

Private Declare Function CreateThreadForRegister Lib "KERNEL32" Alias "CreateThread" _
        (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, _
        ByVal lpparameter As Long, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long

Private Declare Function WaitForSingleObject Lib "KERNEL32" _
        (ByVal hHandle As Long, _
        ByVal dwMilliseconds As Long) As Long

Private Declare Function GetExitCodeThread Lib "KERNEL32" _
        (ByVal hThread As Long, lpExitCode As Long) As Long

Private Declare Sub ExitThread Lib "KERNEL32" (ByVal dwExitCode As Long)

Private Const STATUS_WAIT_0 = &H0
Private Const WAIT_OBJECT_0 = ((STATUS_WAIT_0) + 0)

Public Enum REGISTER_FUNCTIONS
    DllRegisterServer = 1
    DllUnregisterServer = 2
End Enum

Public Enum Status
    [File Could Not Be Loaded Into Memory Space] = 1
    [Not A Valid ActiveX Component] = 2
    [ActiveX Component Registration Failed] = 3
    [ActiveX Component Registered Successfully] = 4
    [ActiveX Component UnRegistered Successfully] = 5
End Enum


Public Function RegisterComponent(ByVal FileName As String, _
                                  ByVal RegFunction As REGISTER_FUNCTIONS) As Status

    '**********************************************************************************
    'Author: Vasudevan S
    'Helena, MT
    'Function: RegisterComponent
    'Purpose: Registers/Unregisters any ActiveX DLL/EXE/OCX component
    'Entry Points in ActiveX DLL/EXE/OCX are DllRegisterServer and DllUnRegisterServer
    'Input: FileName:       Any valid file with complete path
    'RegFunction:   Enumerated Type(DllRegisterServer, DllUnregisterServer)
    'Returns: Returns the status of the call in a enumerated type
    'Comments: The utility REGSVR32.EXE need not be used to register/unregister ActiveX
    'components. This code can be embedded inside any application that needs
    'to register/unregister any ActiveX component from within the code base
    'SAMPLE FORM IS INCLUDED
    'WORKS IN VB5.0/6.0

    'HOW TO CALL:
    '-----------
    'Dim mEnum As STATUS
    '
    'TO REGISTER A COMPONENT USE
    'mEnum = RegisterComponent("C:\windows\system\filename.dll", DllRegisterServer) 'to Register
    '
    'If mEnum = [File Could Not Be Loaded Into Memory Space] Then
    '   MsgBox "Your Message Here", vbExclamation
    'ElseIf mEnum = [Not A Valid ActiveX Component] Then
    '   MsgBox "Your Message Here", vbExclamation
    'ElseIf mEnum = [ActiveX Component Registration Failed] Then
    '   MsgBox "Your Message Here", vbExclamation
    'ElseIf mEnum = [ActiveX Component Registered Successfully] Then
    '   MsgBox "Your Message Here", vbExclamation
    'End If
    '
    'TO UNREGISTER A COMPONENT USE
    'mEnum = RegisterComponent("C:\windows\system\filename.dll", DllUnRegisterServer) 'to UnRegister
    '
    'If mEnum = [File Could Not Be Loaded Into Memory Space] Then
    '   MsgBox "Your Message Here", vbExclamation
    'ElseIf mEnum = [Not A Valid ActiveX Component] Then
    '   MsgBox "Your Message Here", vbExclamation
    'ElseIf mEnum = [ActiveX Component Registration Failed] Then
    '   MsgBox "Your Message Here", vbExclamation
    'ElseIf mEnum = [ActiveX Component UnRegistered Successfully] Then
    '   MsgBox "Your Message Here", vbExclamation
    'End If
    '************************************************************************************

  Dim lngLib As Long
  Dim  lngProcAddress As Long
  Dim  lpThreadID As Long
  Dim  fSuccess As Long
  Dim  dwExitCode As Long
  Dim  hThread As Long

    If FileName = vbNullString Then Exit Function

    lngLib = LoadLibraryRegister(FileName)

    If lngLib = 0 Then
        RegisterComponent = [File Could Not Be Loaded Into Memory Space]    'Couldn't load component
        Exit Function
    End If

    Select Case RegFunction
     Case REGISTER_FUNCTIONS.DllRegisterServer
        lngProcAddress = GetProcAddressRegister(lngLib, "DllRegisterServer")

     Case REGISTER_FUNCTIONS.DllUnregisterServer
        lngProcAddress = GetProcAddressRegister(lngLib, "DllUnregisterServer")

     Case Else
    End Select

    If lngProcAddress = 0 Then
        RegisterComponent = [Not A Valid ActiveX Component]               'Not a Valid ActiveX Component
        If lngLib Then Call FreeLibraryRegister(lngLib)
        Exit Function
     Else
        hThread = CreateThreadForRegister(ByVal 0&, 0&, ByVal lngProcAddress, ByVal 0&, 0&, lpThreadID)

        If hThread Then
            fSuccess = (WaitForSingleObject(hThread, 10000) = WAIT_OBJECT_0)

            If Not fSuccess Then
                Call GetExitCodeThread(hThread, dwExitCode)
                Call ExitThread(dwExitCode)
                RegisterComponent = [ActiveX Component Registration Failed]        'Couldn't Register.
                If lngLib Then Call FreeLibraryRegister(lngLib)
                Exit Function
             Else

                If RegFunction = DllRegisterServer Then
                    RegisterComponent = [ActiveX Component Registered Successfully]         'Success. OK
                 ElseIf RegFunction = DllUnregisterServer Then
                    RegisterComponent = [ActiveX Component UnRegistered Successfully]         'Success. OK
                End If

            End If
            Call CloseHandle(hThread)
            If lngLib Then Call FreeLibraryRegister(lngLib)
        End If

    End If

End Function

