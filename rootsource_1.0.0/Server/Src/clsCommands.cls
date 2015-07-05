VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Sub SetServerName(ByVal Name As String)
    GAME_NAME = Trim$(Name)
End Sub

Sub SetServerPort(ByVal Port As Long)
    GAME_PORT = Port
End Sub

Sub SetMaxPlayers(ByVal Players As Integer)
    MAX_PLAYERS = Val(Players)
End Sub

Sub SetMaxMaps(ByVal Maps As Integer)
    MAX_MAPS = Val(Maps)
End Sub

Sub SetMaxItems(ByVal Items As Integer)
    MAX_ITEMS = Val(Items)
End Sub

Sub SetMaxShops(ByVal Shops As Integer)
    MAX_SHOPS = Val(Shops)
End Sub

Sub SetMaxSpells(ByVal Spells As Integer)
    MAX_SPELLS = Val(Spells)
End Sub

Sub SetMaxNPCs(ByVal NPCs As Integer)
    MAX_NPCS = Val(NPCs)
End Sub

Sub SetWebsite(ByVal Site As String)
    GAME_WEBSITE = Trim$(Site)
End Sub

Sub SetStartPosition(ByVal Map As Integer, ByVal X As Byte, ByVal y As Byte)
    START_MAP = Val(Map)
    START_X = Val(X)
    START_Y = Val(y)
End Sub
