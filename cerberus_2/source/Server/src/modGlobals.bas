Attribute VB_Name = "modGlobals"
'   This file is part of the Cerberus Engine 2nd Edition.
'
'    The Cerberus Engine 2nd Edition is free software; you can redistribute it
'    and/or modify it under the terms of the GNU General Public License as
'    published by the Free Software Foundation; either version 2 of the License,
'    or (at your option) any later version.
'
'    Cerberus 2nd Edition is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Cerberus 2nd Edition; if not, write to the Free Software
'    Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301  USA

Option Explicit

Public GAME_NAME As String
Public GAME_WEBSITE As String

' Used for respawning items
Public SpawnSeconds As Long

' Used for closing key doors again
Public KeyTimer As Long

' Used for logging
Public ServerLog As Boolean

' Used for player looping
Public HighIndex As Long

' Used for parsing
Public SEP_CHAR As String * 1
Public END_CHAR As String * 1

' Maximum classes
Public Max_Classes As Byte

' Used for placing spawn points
Public RSpawnNum As Byte
Public NSpawnNum As Byte

' Used for Socket Collection
Public GameServer As clsServer
Public Sockets As colSockets

' Defined MAX values
Public MAX_PLAYERS As Long
Public MAX_SPELLS As Long
Public MAX_SKILLS As Long
Public MAX_SHOPS As Long
Public MAX_ITEMS As Long
Public MAX_NPCS As Long
Public MAX_QUESTS As Long
Public MAX_MAPS As Long
Public MAX_MAP_ITEMS As Long
Public MAX_GUILDS As Long
Public MAX_GUILD_MEMBERS As Long
'Public MAX_EMOTICONS As Long
'Public MAX_LEVEL As Long
'Public Scripting As Long

' Global map arrays
Public Map() As MapRec                    ' Re-dimmed in InitServer
Public MapItem() As MapItemRec            ' Re-dimmed
Public MapNpc() As MapNpcRec              ' Re-dimmed
Public MapResource() As MapResourceRec    ' Re-dimmed
Public TempTile() As TempTileRec          ' Re-dimmed
Public PushTile() As PushTileRec          ' Re-dimmed
Public PlayersOnMap() As Long             ' Re-dimmed
' Global player arrays
Public Player() As AccountRec             ' Re-dimmed
Public Spells(1 To MAX_PLAYER_SPELLS) As PlayerSpellRec
Public Skills(1 To MAX_PLAYER_SKILLS) As PlayerSkillRec
Public Quests(1 To MAX_PLAYER_QUESTS) As PlayerQuestRec
Public Maps(1 To MAX_PLAYER_MAPS) As PlayerMapRec
' General global arrays
Public Class() As ClassRec
Public Item() As ItemRec                  ' Re-dimmed
Public Npc() As NpcRec                    ' Re-dimmed
Public Shop() As ShopRec                  ' Re-dimmed
Public Skill() As SkillRec                ' Re-dimmed
Public Spell() As SpellRec                ' Re-dimmed
Public Quest() As QuestRec                ' Re-dimmed
Public Guild() As GuildRec                ' Re-dimmed
' Global Menu & GUI arrays
Public GUI(1 To MAX_GUIS) As GUIRec
Public Background(1 To 7) As GUIBackgroundRec
Public Menu(1 To 5) As GUIDataRec
Public Login(1 To 4) As GUIDataRec
Public NewAcc(1 To 4) As GUIDataRec
Public DelAcc(1 To 4) As GUIDataRec
Public Credits(1 To 2) As GUIDataRec
Public Chars(1 To 5) As GUIDataRec
Public NewChar(1 To 14) As GUIDataRec

