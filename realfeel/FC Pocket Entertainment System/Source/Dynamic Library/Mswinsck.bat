@ECHO OFF
ECHO :: Registering Mswinsck.ocx for the FriendCodes Pocket Entertainment System

COPY Mswinsck.ocx C:\Windows\System32 /Y
REGSVR32 Mswinsck.ocx /S

ECHO Registered Successfully!