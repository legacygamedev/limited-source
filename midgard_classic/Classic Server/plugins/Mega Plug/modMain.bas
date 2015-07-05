Attribute VB_Name = "modMain"
Option Explicit

' Edit the Username/Password with the bot username/password.

Public Codes As Object

Public Const DirUp = 0
Public Const DirDown = 1
Public Const DirLeft = 2
Public Const DirRight = 3

Public Const Username = "TriviaBot"
Public Const Password = "botpass"
Public Const CharIndex = 1

Public Const BotMap = 1
Public Const BotX = 4
Public Const BotY = 3
Public Const BotDir = DirDown

Public BotIndex As Integer
Public LocalSocket As Integer
Public PluginNum As Integer
