Attribute VB_Name = "ModSharedTCP"
'The SharedTCP module holds the ID of packets. These IDs work in the exact same manner as the old method of
'just sending a string for the ID, but are much shorter. For example, instead of sending "MaxInfo", you send
'the variable PacketID.MaxInfo. The server and client must hold the same value for this variable, just like
'they would both have to know the string "maxinfo" with the old method. This is why this module is named
'SharedTCP - because it must be indentical on the server and client. Every time you make a modification to
'this module, copy and paste ALL of the code in the module to the other. Ie if you edit this code in the
'server project, copy it to the client project, or vise versa. Everything else is the exact same, except
'for instead of typing "MaxInfo", you're typing PacketID.MaxInfo, and saving a LOT of bandwidth! :)
'
'There is an issue with this method - you can only have 255 packets. Though, this is easily avoidable. If
'you do ever need over 255 packets, on the 255th packet, instead of giving it a command like the other
'packets, you mark it as a "Packet Extension". This way you can have 255 more packets. So for example, on
'an extended packet, instead of sending a packet in this manner:
'
' = PacketID.MaxInfo & ...
'
'You will have:
'
' = PacketID.PacketExtension & PacketID.MaxInfo & ...
'
'This means your packet will be 2 bytes instead of one for the header, but that is still better than 10. ;)

Option Explicit

Public Type PacketID

    'Server -> Client only
    MaxInfo As String * 1
    NPCHP As String * 1
    AttributeNPCHP As String * 1
    AlertMsg As String * 1
    PlainMsg As String * 1
    AllChars As String * 1
    LoginOK As String * 1
    News As String * 1
    NewCharClasses As String * 1
    ClassesData As String * 1
    GameClock As String * 1
    InGame As String * 1
    PlayerInv As String * 1
    PlayerInvUpdate As String * 1
    PlayerBank As String * 1
    PlayerBankUpdate As String * 1
    OpenBank As String * 1
    PlayerWornEQ As String * 1
    PlayerPoints As String * 1
    CusSprite As String * 1
    PlayerHP As String * 1
    PlayerMP As String * 1
    MapMsg2 As String * 1
    ScriptBubble As String * 1
    PlayerSP As String * 1
    PlayerStatsPacket As String * 1
    PlayerData As String * 1
    PlayerLevel As String * 1
    UpdateSprite As String * 1
    NPCMove As String * 1
    AttributeNPCMove As String * 1
    NPCDir As String * 1
    AttributeNPCDir As String * 1
    PlayerXY As String * 1
    RemoveMembers As String * 1
    UpdateMembers As String * 1
    NPCAttack As String * 1
    AttributeNPCAttack As String * 1
    CheckForMap As String * 1
    TileCheck As String * 1
    TileCheckAttribute As String * 1
    MapItemData As String * 1
    MapNPCData As String * 1
    '//!! MapAttributeNPCData As String * 1
    MapDone As String * 1
    MapMsg As String * 1
    SpawnItem As String * 1
    ItemEditor As String * 1
    UpdateItem As String * 1
    Mouse As String * 1
    MapWeather As String * 1
    SpawnNPC As String * 1
    SpawnAttributeNPC As String * 1
    NPCDead As String * 1
    AttributeNPCDead As String * 1
    NPCEditor As String * 1
    UpdateNPC As String * 1
    MapKey As String * 1
    EditMap As String * 1
    EditHouse As String * 1
    Main As String * 1
    ShopEditor As String * 1
    UpdateShop As String * 1
    SpellEditor As String * 1
    UpdateSpell As String * 1
    GoShop As String * 1
    NameColor As String * 1
    Fog As String * 1
    Lights As String * 1
    BlitPlayerDmg As String * 1
    BlitNPCDmg As String * 1
    PPTrading As String * 1
    OnCanon As String * 1
    CanonOff As String * 1
    DTime As String * 1
    UpdateTradeItem As String * 1
    Trading As String * 1
    PPChatting As String * 1
    Sound As String * 1
    SpriteChange As String * 1
    HouseBuy As String * 1
    ChangeDir As String * 1
    FlashEvent As String * 1
    EmoticonEditor As String * 1
    ElementEditor As String * 1
    QuestEditor As String * 1
    UpdateQuest As String * 1
    SkillEditor As String * 1
    UpdateSkill As String * 1
    SkillInfo As String * 1
    UpdateElement As String * 1
    UpdateEmoticon As String * 1
    ArrowEditor As String * 1
    UpdateArrow As String * 1
    HookShot As String * 1
    CheckSprite As String * 1
    ActionName As String * 1
    Time As String * 1
    Wierd As String * 1
    SpellAnim As String * 1
    ScriptSpellAnim As String * 1
    LevelUp As String * 1
    DamageDisplay As String * 1
    '//!! ItemBreak As String * 1 Unused packet
    ItemWorn As String * 1
    ForceHouseClose As String * 1
    SetSpeed As String * 1
    ShowCustomMenu As String * 1
    CloseCustomMenu As String * 1
    LoadPicCustomMenu As String * 1
    LoadLabelCustomMenu As String * 1
    LoadTextBoxCustomMenu As String * 1
    LoadInternetWindow As String * 1
    ReturnCustomBoxMsg As String * 1
    SayMsg As String * 1
    BroadcastMsg As String * 1
    EndShot As String * 1
    BankMsg As String * 1
    
    'Client -> Server only
    GatGlasses As String * 1
    NewFAccountIED As String * 1
    DelimAccounted As String * 1
    Logination As String * 1
    GiveMeTheMax As String * 1
    AddAChara As String * 1
    DelimboCharu As String * 1
    Usagakarim As String * 1
    Mail As String * 1
    GuildChangeAccess As String * 1
    GuildDisown As String * 1
    GuildLeave As String * 1
    MakeGuild As String * 1
    GuildMember As String * 1
    GuildTrainee As String * 1
    EmoteMsg As String * 1
    EditMain As String * 1
    SaveMain As String * 1
    UseItem As String * 1
    PlayerMoveMouse As String * 1
    Warp As String * 1
    UseStatPoint As String * 1
    PlayerInfoRequest As String * 1
    SetSprite As String * 1
    SetPlayerSprite As String * 1
    GetStats As String * 1
    CanonShoot As String * 1
    RequestNewMap As String * 1
    NeedMap As String * 1
    '//!! NeedMapNum2 As String * 1 Unused packet
    MapGetItem As String * 1
    MapDropItem As String * 1
    MapRespawn As String * 1
    KickPlayer As String * 1
    Banlist As String * 1
    BanDestroy As String * 1
    BanPlayer As String * 1
    RequestEditMap As String * 1
    RequestEditHouse As String * 1
    RequestEditItem As String * 1
    SaveItem As String * 1
    '//!! EnableDayNight As String * 1 Unused packet
    DayNight As String * 1
    RequestEditNPC As String * 1
    SaveNPC As String * 1
    RequestEditShop As String * 1
    SaveShop As String * 1
    RequestEditSpell As String * 1
    SaveSpell As String * 1
    ForgetSpell As String * 1
    Key As String * 1
    SetAccess As String * 1
    WhosOnline As String * 1
    SetMOTD As String * 1
    Buy As String * 1
    SellItem As String * 1
    FixItem As String * 1
    Search As String * 1
    PlayerChat As String * 1
    AChat As String * 1
    DChat As String * 1
    SendChat As String * 1
    PPTrade As String * 1
    ATrade As String * 1
    DTrade As String * 1
    UpdateTradeInv As String * 1
    SwapItems As String * 1
    Party As String * 1
    JoinParty As String * 1
    LeaveParty As String * 1
    '//!! PartyChat As String * 1 Unused packet
    HotScript1 As String * 1
    HotScript2 As String * 1
    HotScript3 As String * 1
    HotScript4 As String * 1
    Cast As String * 1
    RequestLocation As String * 1
    Refresh As String * 1
    BuySprite As String * 1
    ClearOwner As String * 1
    BuyHouse As String * 1
    CheckCommands As String * 1
    RequestEditArrow As String * 1
    SaveArrow As String * 1
    RequestEditEmoticon As String * 1
    RequestEditElement As String * 1
    RequestEditSkill As String * 1
    RequestEditQuest As String * 1
    SaveEmoticon As String * 1
    SaveSkill As String * 1
    SaveQuest As String * 1
    SaveElement As String * 1
    GMTime As String * 1
    WarpTo As String * 1
    ArrowHit As String * 1
    BankDeposit As String * 1
    BankWithdraw As String * 1
    ReloadScripts As String * 1
    CustomMenuClick As String * 1
    ReturningCustomBoxMsg As String * 1

    'Server <-> Client (sent and handled by both)
    'In optimal conditions, every packet would be listed under here to preserve IDs
    PlayerMove As String * 1
    PlayerDir As String * 1
    Attack As String * 1
    MapData As String * 1
    GlobalMsg As String * 1
    PlayerMsg As String * 1
    AdminMsg As String * 1
    EditItem As String * 1
    EditNPC As String * 1
    EditShop As String * 1
    EditSpell As String * 1
    Spells As String * 1
    Weather As String * 1
    OnlineList As String * 1
    QTrade As String * 1
    QChat As String * 1
    Prompt As String * 1
    QueryBox As String * 1
    EditQuest As String * 1
    EditSkill As String * 1
    EditElement As String * 1
    EditEmoticon As String * 1
    EditArrow As String * 1
    CheckArrows As String * 1
    MapReport As String * 1
    CheckEmoticons As String * 1
    ScriptTile As String * 1

End Type
Public PacketID As PacketID


Public Sub InitPacketIDs()

    With PacketID
        'Chr$(0) reserved for SEP_CHAR
        'Chr$(237) reserved for END_CHAR
        .MaxInfo = Chr$(1)
        .NPCHP = Chr$(2)
        .AttributeNPCHP = Chr$(3)
        .AlertMsg = Chr$(4)
        .PlainMsg = Chr$(5)
        .AllChars = Chr$(6)
        .LoginOK = Chr$(7)
        .News = Chr$(8)
        .NewCharClasses = Chr$(9)
        .ClassesData = Chr$(10)
        .GameClock = Chr$(11)
        .InGame = Chr$(12)
        .PlayerInv = Chr$(13)
        .PlayerInvUpdate = Chr$(14)
        .PlayerBank = Chr$(15)
        .PlayerBankUpdate = Chr$(16)
        .OpenBank = Chr$(17)
        .PlayerWornEQ = Chr$(18)
        .PlayerPoints = Chr$(19)
        .CusSprite = Chr$(20)
        .PlayerHP = Chr$(21)
        .PlayerMP = Chr$(22)
        .MapMsg2 = Chr$(23)
        .ScriptBubble = Chr$(24)
        .PlayerSP = Chr$(25)
        .PlayerStatsPacket = Chr$(26)
        .PlayerData = Chr$(27)
        .PlayerLevel = Chr$(28)
        .UpdateSprite = Chr$(29)
        .NPCMove = Chr$(30)
        .AttributeNPCMove = Chr$(31)
        .NPCDir = Chr$(32)
        .AttributeNPCDir = Chr$(33)
        .PlayerXY = Chr$(34)
        .RemoveMembers = Chr$(35)
        .UpdateMembers = Chr$(36)
        .NPCAttack = Chr$(37)
        .AttributeNPCAttack = Chr$(38)
        .CheckForMap = Chr$(39)
        .TileCheck = Chr$(40)
        .TileCheckAttribute = Chr$(41)
        .MapItemData = Chr$(42)
        .MapNPCData = Chr$(43)
        '//!! .MapAttributeNPCData = Chr$(44) Unused packet
        .MapDone = Chr$(45)
        .MapMsg = Chr$(46)
        .SpawnItem = Chr$(47)
        .ItemEditor = Chr$(48)
        .UpdateItem = Chr$(49)
        .Mouse = Chr$(50)
        .MapWeather = Chr$(51)
        .SpawnNPC = Chr$(52)
        .SpawnAttributeNPC = Chr$(53)
        .NPCDead = Chr$(54)
        .AttributeNPCDead = Chr$(55)
        .NPCEditor = Chr$(56)
        .UpdateNPC = Chr$(57)
        .MapKey = Chr$(58)
        .EditMap = Chr$(59)
        .EditHouse = Chr$(60)
        .Main = Chr$(61)
        .ShopEditor = Chr$(62)
        .UpdateShop = Chr$(63)
        .SpellEditor = Chr$(64)
        .UpdateSpell = Chr$(65)
        .GoShop = Chr$(66)
        .NameColor = Chr$(67)
        .Fog = Chr$(68)
        .Lights = Chr$(69)
        .BlitPlayerDmg = Chr$(70)
        .BlitNPCDmg = Chr$(71)
        .PPTrading = Chr$(72)
        .OnCanon = Chr$(73)
        .CanonOff = Chr$(74)
        .DTime = Chr$(75)
        .UpdateTradeItem = Chr$(76)
        .Trading = Chr$(77)
        .PPChatting = Chr$(78)
        .Sound = Chr$(79)
        .SpriteChange = Chr$(80)
        .HouseBuy = Chr$(81)
        .ChangeDir = Chr$(82)
        .FlashEvent = Chr$(83)
        .EmoticonEditor = Chr$(84)
        .ElementEditor = Chr$(85)
        .QuestEditor = Chr$(86)
        .UpdateQuest = Chr$(87)
        .SkillEditor = Chr$(88)
        .UpdateSkill = Chr$(89)
        .SkillInfo = Chr$(90)
        .UpdateElement = Chr$(91)
        .UpdateEmoticon = Chr$(92)
        .ArrowEditor = Chr$(93)
        .UpdateArrow = Chr$(94)
        .HookShot = Chr$(95)
        .CheckSprite = Chr$(96)
        .ActionName = Chr$(97)
        .Time = Chr$(98)
        .Wierd = Chr$(99)
        .SpellAnim = Chr$(100)
        .ScriptSpellAnim = Chr$(101)
        .LevelUp = Chr$(102)
        .DamageDisplay = Chr$(103)
        '//!! .ItemBreak = Chr$(104)
        .ItemWorn = Chr$(105)
        .ForceHouseClose = Chr$(106)
        .SetSpeed = Chr$(107)
        .ShowCustomMenu = Chr$(108)
        .CloseCustomMenu = Chr$(109)
        .LoadPicCustomMenu = Chr$(110)
        .LoadLabelCustomMenu = Chr$(111)
        .LoadTextBoxCustomMenu = Chr$(112)
        .LoadInternetWindow = Chr$(113)
        .ReturnCustomBoxMsg = Chr$(114)

        .GatGlasses = Chr$(115)
        .NewFAccountIED = Chr$(116)
        .DelimAccounted = Chr$(117)
        .Logination = Chr$(118)
        .GiveMeTheMax = Chr$(119)
        .AddAChara = Chr$(120)
        .DelimboCharu = Chr$(121)
        .Usagakarim = Chr$(122)
        .Mail = Chr$(123)
        .GuildChangeAccess = Chr$(124)
        .GuildDisown = Chr$(125)
        .GuildLeave = Chr$(126)
        .MakeGuild = Chr$(127)
        .GuildMember = Chr$(128)
        .GuildTrainee = Chr$(129)
        .EmoteMsg = Chr$(130)
        .EditMain = Chr$(131)
        .SaveMain = Chr$(132)
        .UseItem = Chr$(133)
        .PlayerMoveMouse = Chr$(134)
        .Warp = Chr$(135)
        .UseStatPoint = Chr$(136)
        .PlayerInfoRequest = Chr$(137)
        .SetSprite = Chr$(138)
        .SetPlayerSprite = Chr$(139)
        .GetStats = Chr$(140)
        .CanonShoot = Chr$(141)
        .RequestNewMap = Chr$(142)
        .NeedMap = Chr$(143)
        '//!! .NeedMapNum2 = Chr$(144)
        .MapGetItem = Chr$(145)
        .MapDropItem = Chr$(146)
        .MapRespawn = Chr$(147)
        .KickPlayer = Chr$(148)
        .Banlist = Chr$(149)
        .BanDestroy = Chr$(150)
        .BanPlayer = Chr$(151)
        .RequestEditMap = Chr$(152)
        .RequestEditHouse = Chr$(153)
        .RequestEditItem = Chr$(154)
        .SaveItem = Chr$(155)
        '//!! .EnableDayNight = Chr$(156)
        .DayNight = Chr$(157)
        .RequestEditNPC = Chr$(158)
        .SaveNPC = Chr$(159)
        .RequestEditShop = Chr$(160)
        .SaveShop = Chr$(161)
        .RequestEditSpell = Chr$(162)
        .SaveSpell = Chr$(163)
        .ForgetSpell = Chr$(164)
        .Key = Chr$(165)
        .SetAccess = Chr$(166)
        .WhosOnline = Chr$(167)
        .SetMOTD = Chr$(168)
        .Buy = Chr$(169)
        .SellItem = Chr$(170)
        .FixItem = Chr$(171)
        .Search = Chr$(172)
        .PlayerChat = Chr$(173)
        .AChat = Chr$(174)
        .DChat = Chr$(175)
        .SendChat = Chr$(176)
        .PPTrade = Chr$(177)
        .ATrade = Chr$(178)
        .DTrade = Chr$(179)
        .UpdateTradeInv = Chr$(180)
        .SwapItems = Chr$(181)
        .Party = Chr$(182)
        .JoinParty = Chr$(183)
        .LeaveParty = Chr$(184)
        '//!! .PartyChat = Chr$(185)
        .HotScript1 = Chr$(186)
        .HotScript2 = Chr$(187)
        .HotScript3 = Chr$(188)
        .HotScript4 = Chr$(189)
        .Cast = Chr$(190)
        .RequestLocation = Chr$(191)
        .Refresh = Chr$(192)
        .BuySprite = Chr$(193)
        .ClearOwner = Chr$(194)
        .BuyHouse = Chr$(195)
        .CheckCommands = Chr$(196)
        .RequestEditArrow = Chr$(197)
        .SaveArrow = Chr$(198)
        .RequestEditEmoticon = Chr$(199)
        .RequestEditElement = Chr$(200)
        .RequestEditSkill = Chr$(201)
        .RequestEditQuest = Chr$(202)
        .SaveEmoticon = Chr$(203)
        .SaveSkill = Chr$(204)
        .SaveQuest = Chr$(205)
        .SaveElement = Chr$(206)
        .GMTime = Chr$(207)
        .WarpTo = Chr$(208)
        .ArrowHit = Chr$(209)
        .BankDeposit = Chr$(210)
        .BankWithdraw = Chr$(211)
        .ReloadScripts = Chr$(212)
        .CustomMenuClick = Chr$(213)
        .ReturningCustomBoxMsg = Chr$(214)

        .PlayerMove = Chr$(215)
        .PlayerDir = Chr$(216)
        .Attack = Chr$(217)
        .MapData = Chr$(218)
        .SayMsg = Chr$(219)
        .BroadcastMsg = Chr$(220)
        .GlobalMsg = Chr$(221)
        .PlayerMsg = Chr$(223)
        .AdminMsg = Chr$(224)
        .EditItem = Chr$(225)
        .EditNPC = Chr$(226)
        .EditShop = Chr$(227)
        .EditSpell = Chr$(228)
        .Spells = Chr$(229)
        .Weather = Chr$(230)
        .OnlineList = Chr$(231)
        .QTrade = Chr$(232)
        .QChat = Chr$(233)
        .Prompt = Chr$(234)
        .QueryBox = Chr$(235)
        .EditQuest = Chr$(236)
        'Chr$(237) reserved for END_CHAR
        .EditSkill = Chr$(238)
        .EditElement = Chr$(239)
        .EditEmoticon = Chr$(240)
        .EditArrow = Chr$(241)
        .CheckArrows = Chr$(242)
        .MapReport = Chr$(243)
        .CheckEmoticons = Chr$(244)
        .ScriptTile = Chr$(245)
        .EndShot = Chr$(246)
        .BankMsg = Chr$(247)
    End With

End Sub


