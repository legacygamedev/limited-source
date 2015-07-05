# Microsoft Developer Studio Project File - Name="COMTools" - Package Owner=<4>
# Microsoft Developer Studio Generated Build File, Format Version 6.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Static Library" 0x0104

CFG=COMTools - Win32 Unicode Debug
!MESSAGE This is not a valid makefile. To build this project using NMAKE,
!MESSAGE use the Export Makefile command and run
!MESSAGE 
!MESSAGE NMAKE /f "COMTools.mak".
!MESSAGE 
!MESSAGE You can specify a configuration when running NMAKE
!MESSAGE by defining the macro CFG on the command line. For example:
!MESSAGE 
!MESSAGE NMAKE /f "COMTools.mak" CFG="COMTools - Win32 Unicode Debug"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "COMTools - Win32 Release" (based on "Win32 (x86) Static Library")
!MESSAGE "COMTools - Win32 Debug" (based on "Win32 (x86) Static Library")
!MESSAGE "COMTools - Win32 Unicode Release" (based on "Win32 (x86) Static Library")
!MESSAGE "COMTools - Win32 Unicode Debug" (based on "Win32 (x86) Static Library")
!MESSAGE 

# Begin Project
# PROP AllowPerConfigDependencies 0
# PROP Scc_ProjName ""$/Web Articles/SocketServers/COMSocketServer2/JetByteTools/COMTools", OCBAAAAA"
# PROP Scc_LocalPath "."
CPP=cl.exe
RSC=rc.exe

!IF  "$(CFG)" == "COMTools - Win32 Release"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "Release"
# PROP BASE Intermediate_Dir "Release"
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "Output\VC6\Release"
# PROP Intermediate_Dir "Output\VC6\Release"
# PROP Target_Dir ""
# ADD BASE CPP /nologo /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_MBCS" /D "_LIB" /YX /FD /c
# ADD CPP /nologo /MT /W3 /GX /O2 /I "..\..\\" /D "NDEBUG" /D "WIN32" /D "_MBCS" /D "_LIB" /D "STRICT" /YX /FD /c
# ADD BASE RSC /l 0x809 /d "NDEBUG"
# ADD RSC /l 0x809 /d "NDEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LIB32=link.exe -lib
# ADD BASE LIB32 /nologo
# ADD LIB32 /nologo

!ELSEIF  "$(CFG)" == "COMTools - Win32 Debug"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir "Debug"
# PROP BASE Intermediate_Dir "Debug"
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 1
# PROP Output_Dir "Output\VC6\Debug"
# PROP Intermediate_Dir "Output\VC6\Debug"
# PROP Target_Dir ""
# ADD BASE CPP /nologo /W3 /Gm /GX /ZI /Od /D "WIN32" /D "_DEBUG" /D "_MBCS" /D "_LIB" /YX /FD /GZ /c
# ADD CPP /nologo /MTd /W3 /Gm /GX /ZI /Od /I "..\..\\" /D "_DEBUG" /D "WIN32" /D "_MBCS" /D "_LIB" /D "STRICT" /YX /FD /GZ /c
# ADD BASE RSC /l 0x809 /d "_DEBUG"
# ADD RSC /l 0x809 /d "_DEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LIB32=link.exe -lib
# ADD BASE LIB32 /nologo
# ADD LIB32 /nologo

!ELSEIF  "$(CFG)" == "COMTools - Win32 Unicode Release"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "COMTools___Win32_Unicode_Release"
# PROP BASE Intermediate_Dir "COMTools___Win32_Unicode_Release"
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "Output\VC6\URelease"
# PROP Intermediate_Dir "Output\VC6\URelease"
# PROP Target_Dir ""
# ADD BASE CPP /nologo /W3 /GX /O2 /I "..\..\\" /D "NDEBUG" /D "WIN32" /D "_MBCS" /D "_LIB" /D "STRICT" /YX /FD /c
# ADD CPP /nologo /MT /W3 /GX /O2 /I "..\..\\" /D "NDEBUG" /D "WIN32" /D "_MBCS" /D "_LIB" /D "STRICT" /D "UNICODE" /D "_UNICODE" /YX /FD /c
# ADD BASE RSC /l 0x809 /d "NDEBUG"
# ADD RSC /l 0x809 /d "NDEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LIB32=link.exe -lib
# ADD BASE LIB32 /nologo
# ADD LIB32 /nologo

!ELSEIF  "$(CFG)" == "COMTools - Win32 Unicode Debug"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir "COMTools___Win32_Unicode_Debug"
# PROP BASE Intermediate_Dir "COMTools___Win32_Unicode_Debug"
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 1
# PROP Output_Dir "Output\VC6\UDebug"
# PROP Intermediate_Dir "Output\VC6\UDebug"
# PROP Target_Dir ""
# ADD BASE CPP /nologo /W3 /Gm /GX /ZI /Od /I "..\..\\" /D "_DEBUG" /D "WIN32" /D "_MBCS" /D "_LIB" /D "STRICT" /YX /FD /GZ /c
# ADD CPP /nologo /MTd /W3 /Gm /GX /ZI /Od /I "..\..\\" /D "_DEBUG" /D "WIN32" /D "_MBCS" /D "_LIB" /D "STRICT" /D "UNICODE" /D "_UNICODE" /YX /FD /GZ /c
# ADD BASE RSC /l 0x809 /d "_DEBUG"
# ADD RSC /l 0x809 /d "_DEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LIB32=link.exe -lib
# ADD BASE LIB32 /nologo
# ADD LIB32 /nologo

!ENDIF 

# Begin Target

# Name "COMTools - Win32 Release"
# Name "COMTools - Win32 Debug"
# Name "COMTools - Win32 Unicode Release"
# Name "COMTools - Win32 Unicode Debug"
# Begin Group "Source Files"

# PROP Default_Filter "cpp;c;cxx;rc;def;r;odl;idl;hpj;bat"
# Begin Source File

SOURCE=.\AsyncServerEventHelper.cpp
# End Source File
# Begin Source File

SOURCE=.\Exception.cpp
# End Source File
# Begin Source File

SOURCE=.\UsesCom.cpp
# End Source File
# Begin Source File

SOURCE=.\Utils.cpp
# End Source File
# End Group
# Begin Group "Header Files"

# PROP Default_Filter "h;hpp;hxx;hm;inl"
# Begin Source File

SOURCE=.\AsyncServerEventHelper.h
# End Source File
# Begin Source File

SOURCE=.\Exception.h
# End Source File
# Begin Source File

SOURCE=.\UsesCom.h
# End Source File
# Begin Source File

SOURCE=.\Utils.h
# End Source File
# End Group
# Begin Group "Lint Options"

# PROP Default_Filter ""
# Begin Source File

SOURCE=.\COMTools.lnt
# End Source File
# Begin Source File

SOURCE=.\std.lnt
# End Source File
# End Group
# Begin Group "Interfaces"

# PROP Default_Filter ""
# Begin Source File

SOURCE=.\AsyncServerEvent.idl

!IF  "$(CFG)" == "COMTools - Win32 Release"

# Begin Custom Build - MIDL
ProjDir=.
InputPath=..\JetByteTools\COMTools\AsyncServerEvent.idl
InputName=AsyncServerEvent

"$(InputName).h" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	midl $(ProjDir)\$(InputName).idl

# End Custom Build

!ELSEIF  "$(CFG)" == "COMTools - Win32 Debug"

# Begin Custom Build - MIDL
ProjDir=.
InputPath=..\JetByteTools\COMTools\AsyncServerEvent.idl
InputName=AsyncServerEvent

"$(InputName).h" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	midl $(ProjDir)\$(InputName).idl

# End Custom Build

!ELSEIF  "$(CFG)" == "COMTools - Win32 Unicode Release"

# Begin Custom Build - MIDL
ProjDir=.
InputPath=..\JetByteTools\COMTools\AsyncServerEvent.idl
InputName=AsyncServerEvent

"$(InputName).h" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	midl $(ProjDir)\$(InputName).idl

# End Custom Build

!ELSEIF  "$(CFG)" == "COMTools - Win32 Unicode Debug"

# Begin Custom Build - MIDL
ProjDir=.
InputPath=..\JetByteTools\COMTools\AsyncServerEvent.idl
InputName=AsyncServerEvent

"$(InputName).h" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	midl $(ProjDir)\$(InputName).idl

# End Custom Build

!ENDIF 

# End Source File
# End Group
# End Target
# End Project
