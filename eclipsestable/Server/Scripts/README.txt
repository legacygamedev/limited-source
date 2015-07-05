Installation:
 * Eclipse Stable and branches.
   1. Move the files and folders in the ES folder to your Scripts folder.
   2. Run the server and turn on scripting.

 * Eclipse Evolution and branches.
   1. Move Main.txt, Main.ess, Events, Functions and Modules to your Scripts folder.
   2. Run the server and turn on scripting.

F.A.Q.
 Q. What are those ESS-files?
   They're Eclipse SadScript modules or files and can be edited by just using notepad.

 Q. What's FSMain.ess for?
   FSMain.ess is the full support main and can be used when the version of Eclipse you're using support
   wildcards for #include. If you don't know, then use Main.ess or Main.txt instead.

 Q. What's the difference between Main.ess and Main.txt?
   The extension is the only difference. Some versions may be conform with GM's standards and may load
   Main.ess instead of Main.txt.

 Q. Why's there an alternative version for Eclipse Stable?
   I'm trying to get a (new) Sadscript standard official. Eclipse Stable may already include some of
   these changes already.

Changelog:
 * GM 1.0.0
   GM 1.0.0 is the initial release of my set of scripts. Several features are:

   * Optimised as for performance.
   * More checks added to prevent crashes and to clarify something when the user did something wrong.
   * Split up every function in ESS-modules.
   * Added comments to explain each module and to categorise the modules in the Main.txt.
   * Replaced the nested if-statements with a list of if-statements which exit when something is wrong to gain more performance.
   * Constants are heavily used to clarify the code.
   * Added more globals so people can directly access every setting in Data.ini in their scripts.

 * GM 1.0.1
   This version contains Main.txt and Main.ess, the first for backward compatibility. Newer versions of Eclipse might want to switch to .ess as the official extension for Eclipse SadScript.

 * GM 1.0.2
   I fixed a few issues which were:

   * Constants and Global modules failed, removed them and added all the constants and globals to the Main module.
   * Added a check to /warpmeto and /warptome to prevent yourself from warping yourself.

 * GM 1.0.3
   New version with a few fixes.

   * It has default values now for all the globals when it can't find the setting for it. So MAX_STAT is no longer required.
   * Renamed TestMain to TestScripts.

 * GM 1.1.0
   A new release with some new features:

   * AccountVar.ess has been added and is a new module which adds GetAccountVar, GetCharVar, PutAccountVar and PutCharVar to Sadscript. These functions allow you to store account data and such.
   * Attributes.ess has been added and is a new module which supports the SetAttribute functions from the original Main. It is still unfinished though, but it already supports all the old functions from the original Main.
   * Added new TILE_TYPE constants for use with SetAttribute and GetAttribute.
   * FSMain.ess has been added and this supports the include wild cards.

 * GM 1.1.1
   * AccountVar.ess and Attributes.ess are now relocated to /Modules/.
   * Added Profile.ess and the (removable) plug in Commands.ess.
   * Changed FSMain.ess, Main.ess and Main.txt to support the Modules folder.
   * Added the "The scripts have been reloaded..." message to OnScriptReload.ess.

 * GM 1.1.2
   * Added another (removable) plug for Profile.ess in MenuScripts.ess to prevent the "Unknown menu type" message.
   * Fixed a position issue, because I removed the Job entry in the custom menu.

 * GM 1.1.3
   * Fixed the STAT-constants in Main.txt, Main.ess and FSMain.ess.
   * Fixed the /warpto-command in Commands.ess.
   * Added a /settime-command to Commands.ess (Parameters: hours minutes seconds).
   * Fixed PlayerLevelUp.ess.
   * Removed the "The scripts have been reloaded..."-message, as the server message seems to work again.

 * GM 1.1.4
   * Negativity bug in PlayerLevelUp.ess should be fixed now.

 * GM 1.1.5
   * Critical update.

 * GM 1.2.0
   * The Profile custom menu now also supports editing and also contains additional information like the IP-address for administrators.
   * Inventory.ess added. It contains the following functions:
     * Function AddPlayerInvItem(Index, Item, Durability)
     * Function AddPlayerInvStackableItem(Index, Item, Amount)
     * Function CountPlayerInvItem(Index, Item)
     * Sub RemovePlayerInvItem(Index, Item, Amount)
     * Function GetPlayerInvSlots(Index)
     * Sub ClearPlayerInv(Index)
   * BCInventory.ess added for backwards compatibility with Godlord's Old Inventory Script. It contains the following functions:
     * Function GetFreeSlots(Index)
     * Sub GiveItem(Index, Number, Durability)
     * Sub GiveCurrency(Index, Number, Amount)
     * Function CanTake(Index, Number, Amount)
     * Sub TakeItem(Index, Number, Amount)
   * Bank.ess added. It contains the following functions:
     * Low-level (Do not use these):
       * Function GetPlayerBnkItemNum(Index, Slot)
       * Function GetPlayerBnkItemValue(Index, Slot)
       * Function GetPlayerBnkItemDur(Index, Slot)
       * Sub SetPlayerBnkItemNum(Index, Slot, Num)
       * Sub SetPlayerBnkItemValue(Index, Slot, Value)
       * Sub SetPlayerBnkItemDur(Index, Slot, Dur)
     * High-level:
       * Function AddPlayerBankItem(Index, Item, Durability)
       * Function AddPlayerBankStackableItem(Index, Item, Amount)
       * Function CountPlayerBankItem(Index, Item)
       * Sub RemovePlayerBankItem(Index, Item, Amount)
       * Function GetPlayerBankSlots(Index)
       * Sub ClearPlayerBank(Index)