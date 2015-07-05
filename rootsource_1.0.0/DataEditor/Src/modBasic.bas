Attribute VB_Name = "modBasic"
Option Explicit

Public Const NAME_LENGTH = 20

Public GameData As DataRec

Type DataRec
  IP As String * NAME_LENGTH
  Port As Integer
  Autoupdater As Byte
  SaveLogin As Byte
  Username As String * NAME_LENGTH
  Password As String * NAME_LENGTH
  Music As Byte
  Sound As Byte
  PlayerNames As Byte
  NpcNames As Byte
  SpellGFX As Byte
  MusicExt As String * 5
  ScreenNum As Integer
  WebAddress As String * 500
  GameName As String * 255
End Type

