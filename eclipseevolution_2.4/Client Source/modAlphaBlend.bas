Attribute VB_Name = "modAlphaBlend"
Option Explicit
Public Declare Sub AlphaBlend Lib "msimg32.dll" (ByVal destHDC As Long, ByVal destLEFT As Long, ByVal destTOP As Long, ByVal destWIDTH As Long, ByVal destHEIGHT As Long, ByVal sourceHDC As Long, ByVal sourceLEFT As Long, ByVal sourceTOP As Long, ByVal sourceWIDTH As Long, ByVal sourceHEIGHT As Long, ByVal BLENDFUNCT As Long)
Public Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal Length As Long)

Public Const AC_SRC_OVER = &H0

Public Type BLENDFUNCTION
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type

Public BF As BLENDFUNCTION
Public Alpha_lBF As Long
