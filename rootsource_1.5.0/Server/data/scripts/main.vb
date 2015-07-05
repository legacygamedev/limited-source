'**************************************************************************
'* This is a sample Main.as script for rootSource. You may modify this    *
'* to suit the needs of your server. More can be added through the code.  *
'**************************************************************************

'**************************
'* Edit the values below. *
'**************************
Public Const GAME_NAME = "rootSource" ' Name of game
Public Const WEB_SITE = "http://www.onlinegamecore.com" ' Website

Public Const GAME_PORT = 7234 ' Run off What Port?
Public Const MAX_PLAYERS = 50 ' Max Players
Public Const MAX_MAPS = 50 ' Max Maps
Public Const MAX_ITEMS = 255 ' Max Items
Public Const MAX_SHOPS = 255 ' Max Shops
Public Const MAX_SPELLS = 255 ' Max Spells
Public Const MAX_NPCS = 255 ' Max NPCs

' Where will New Players begin?
Public Const START_MAP = 5
Public Const START_X = 5
Public Const START_Y = 8

Public Const DEBUG = NO ' Find errors in script file

'************************
'* Game core constants. *
'************************

' Map constants [DO NOT EDIT]
Public Const MAP_MORAL_NONE = 0
Public Const MAP_MORAL_SAFE = 1


' Item constants [DO NOT EDIT]
Public Const ITEM_TYPE_NONE = 0
Public Const ITEM_TYPE_WEAPON = 1
Public Const ITEM_TYPE_ARMOR = 2
Public Const ITEM_TYPE_HELMET = 3
Public Const ITEM_TYPE_SHIELD = 4
Public Const ITEM_TYPE_POTIONADDHP = 5 ' The rest may not be needed
Public Const ITEM_TYPE_POTIONADDMP = 6
Public Const ITEM_TYPE_POTIONADDSP = 7
Public Const ITEM_TYPE_POTIONSUBHP = 8
Public Const ITEM_TYPE_POTIONSUBMP = 9
Public Const ITEM_TYPE_POTIONSUBSP = 10
Public Const ITEM_TYPE_KEY = 11
Public Const ITEM_TYPE_CURRENCY = 12
Public Const ITEM_TYPE_SPELL = 13
Public Const ITEM_TYPE_WARP = 14

' Color constants [DO NOT EDIT]
Public Const BLACK = 0
Public Const BLUE = 1
Public Const GREEN = 2
Public Const CYAN = 3
Public Const RED = 4
Public Const MAGENTA = 5
Public Const BROWN = 6
Public Const GREY = 7
Public Const DARKGREY = 8
Public Const BRIGHTBLUE = 9
Public Const BRIGHTGREEN = 10
Public Const BRIGHTCYAN = 11
Public Const BRIGHTRED = 12
Public Const PINK = 13
Public Const YELLOW = 14
Public Const WHITE = 15

' Boolean values [DO NOT EDIT]
Public Const NO = 0 ' False
Public Const YES = -1 ' True


Sub ServerSet()
'**************************************************
'* This event is fired to setup your information. *
'**************************************************

    ' Initialize server
    Call SetServerName(GAME_NAME)
    Call SetWebsite(WEB_SITE)
    Call SetServerPort(GAME_PORT)
    Call SetMaxPlayers(MAX_PLAYERS)
    Call SetMaxMaps(MAX_MAPS)
    Call SetMaxItems(MAX_ITEMS)
    Call SetMaxShops(MAX_SHOPS)
    Call SetMaxSpells(MAX_SPELLS)
    Call SetMaxNPCs(MAX_NPCS)
    Call SetStartPosition(START_MAP, START_X, START_Y)

    Call SetDebugScripting(DEBUG)

End Sub