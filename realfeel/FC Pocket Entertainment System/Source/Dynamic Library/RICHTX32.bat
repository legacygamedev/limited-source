@ECHO OFF
ECHO :: Registering RICHTX32.ocx for the FriendCodes Pocket Entertainment System

COPY RICHTX32.ocx C:\Windows\System32 /Y
REGSVR32 RICHTX32.ocx /S

ECHO Registered Successfully!