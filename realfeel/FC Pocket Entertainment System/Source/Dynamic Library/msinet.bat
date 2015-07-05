@ECHO OFF
ECHO :: Registering msinet.ocx for the FriendCodes Pocket Entertainment System

COPY msinet.ocx C:\Windows\System32 /Y
REGSVR32 msinet.ocx /S

ECHO Registered Successfully!