@echo off
echo Registering dll and ocx controls...
regsvr32 /s TABCTL32.OCX
regsvr32 /s MSINET.OCX
regsvr32 /s MSWINSCK.OCX
echo Done