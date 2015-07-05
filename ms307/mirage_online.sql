# MySQL-Front Dump 2.5
#
# Host: localhost   Database: mirage_online
# --------------------------------------------------------
# Server version 4.0.15-nt

USE mirage_online;


#
# Table structure for table 'accounts'
#

DROP TABLE IF EXISTS `accounts`;
CREATE TABLE `accounts` (
  `FKey` int(11) NOT NULL auto_increment,
  `Login` tinytext NOT NULL,
  `Password` tinytext NOT NULL,
  `Suspended` tinyint(1) unsigned default '0',
  `Banned` tinyint(1) unsigned default '0',
  `HDModel` text,
  `HDSerial` text,
  `PriChar` int(11) unsigned default '0',
  PRIMARY KEY  (`FKey`)
) TYPE=MyISAM;



#
# Table structure for table 'bans'
#

DROP TABLE IF EXISTS `bans`;
CREATE TABLE `bans` (
  `Fkey` int(11) NOT NULL auto_increment,
  `nDate` tinytext,
  `nTime` tinytext,
  `BannedIP` tinytext NOT NULL,
  `BannedBY` tinytext NOT NULL,
  PRIMARY KEY  (`Fkey`)
) TYPE=MyISAM;



#
# Table structure for table 'characters'
#

DROP TABLE IF EXISTS `characters`;
CREATE TABLE `characters` (
  `FKey` int(11) NOT NULL auto_increment,
  `Account` tinyint(11) unsigned default '0',
  `Name` tinytext NOT NULL,
  `Class` int(11) NOT NULL default '0',
  `Sex` int(11) NOT NULL default '0',
  `Sprite` int(11) NOT NULL default '0',
  `Level` int(11) NOT NULL default '0',
  `Exp` int(11) NOT NULL default '0',
  `Access` int(11) NOT NULL default '0',
  `PK` int(11) NOT NULL default '0',
  `Guild` int(11) NOT NULL default '0',
  `HP` int(11) NOT NULL default '0',
  `MP` int(11) NOT NULL default '0',
  `SP` int(11) NOT NULL default '0',
  `STR` int(11) NOT NULL default '0',
  `DEF` int(11) NOT NULL default '0',
  `SPEED` int(11) NOT NULL default '0',
  `MAGI` int(11) NOT NULL default '0',
  `POINTS` int(11) NOT NULL default '0',
  `ArmorSlot` int(11) NOT NULL default '0',
  `WeaponSlot` int(11) NOT NULL default '0',
  `HelmetSlot` int(11) NOT NULL default '0',
  `ShieldSlot` int(11) NOT NULL default '0',
  `Map` int(11) NOT NULL default '0',
  `X` int(11) NOT NULL default '0',
  `Y` int(11) NOT NULL default '0',
  `Dir` int(11) NOT NULL default '0',
  `Inventory` text NOT NULL,
  `Spells` text NOT NULL,
  `Suspended` tinyint(1) unsigned default '0',
  `Jailed` tinyint(1) unsigned default '0',
  PRIMARY KEY  (`FKey`)
) TYPE=MyISAM;



#
# Table structure for table 'classes'
#

DROP TABLE IF EXISTS `classes`;
CREATE TABLE `classes` (
  `FKey` int(11) NOT NULL auto_increment,
  `Name` tinytext NOT NULL,
  `Sprite` int(11) NOT NULL default '1',
  `STR` int(11) NOT NULL default '1',
  `DEF` int(11) NOT NULL default '1',
  `SPEED` int(11) NOT NULL default '1',
  `MAGI` int(11) NOT NULL default '1',
  PRIMARY KEY  (`FKey`)
) TYPE=MyISAM;



#
# Table structure for table 'items'
#

DROP TABLE IF EXISTS `items`;
CREATE TABLE `items` (
  `FKey` int(11) NOT NULL auto_increment,
  `Name` tinytext NOT NULL,
  `Pic` int(11) NOT NULL default '0',
  `Type` int(11) NOT NULL default '0',
  `Data1` int(11) NOT NULL default '0',
  `Data2` int(11) NOT NULL default '0',
  `Data3` int(11) NOT NULL default '0',
  `UnBreakable` tinyint(1) unsigned NOT NULL default '0',
  `Locked` tinyint(1) unsigned NOT NULL default '0',
  `Disabled` tinyint(1) unsigned NOT NULL default '0',
  `Assigned` tinyint(11) unsigned NOT NULL default '0',
  PRIMARY KEY  (`FKey`)
) TYPE=MyISAM;



#
# Table structure for table 'logs'
#

DROP TABLE IF EXISTS `logs`;
CREATE TABLE `logs` (
  `FKey` int(11) NOT NULL auto_increment,
  `nDate` tinytext NOT NULL,
  `nTime` tinytext NOT NULL,
  `nType` tinytext NOT NULL,
  `Entry` text NOT NULL,
  PRIMARY KEY  (`FKey`)
) TYPE=MyISAM;



#
# Table structure for table 'maps'
#

DROP TABLE IF EXISTS `maps`;
CREATE TABLE `maps` (
  `FKey` int(11) NOT NULL auto_increment,
  `Name` text NOT NULL,
  `Revision` int(11) NOT NULL default '0',
  `Moral` int(11) NOT NULL default '0',
  `Up` int(11) NOT NULL default '0',
  `Down` int(11) NOT NULL default '0',
  `mLeft` int(11) NOT NULL default '0',
  `mRight` int(11) NOT NULL default '0',
  `Music` int(11) NOT NULL default '0',
  `BootMap` int(11) NOT NULL default '0',
  `BootX` int(11) NOT NULL default '0',
  `BootY` int(11) NOT NULL default '0',
  `Shop` int(11) NOT NULL default '0',
  `Indoors` int(11) NOT NULL default '0',
  `Tiles` text NOT NULL,
  `NPCs` text NOT NULL,
  PRIMARY KEY  (`FKey`)
) TYPE=MyISAM;



#
# Table structure for table 'npcs'
#

DROP TABLE IF EXISTS `npcs`;
CREATE TABLE `npcs` (
  `FKey` int(11) NOT NULL auto_increment,
  `Name` tinytext NOT NULL,
  `Sprite` int(11) NOT NULL default '0',
  `Behavior` int(11) NOT NULL default '0',
  `Range` int(11) NOT NULL default '0',
  `DropChance` int(11) NOT NULL default '0',
  `DropItem` int(11) NOT NULL default '0',
  `DropItemValue` int(11) NOT NULL default '0',
  `STR` int(11) NOT NULL default '0',
  `DEF` int(11) NOT NULL default '0',
  `SPEED` int(11) NOT NULL default '0',
  `MAGI` int(11) NOT NULL default '0',
  `AttackSay` tinytext NOT NULL,
  `SpawnSecs` int(11) NOT NULL default '0',
  PRIMARY KEY  (`FKey`)
) TYPE=MyISAM;



#
# Table structure for table 'settings'
#

DROP TABLE IF EXISTS `settings`;
CREATE TABLE `settings` (
  `FKey` int(11) NOT NULL auto_increment,
  `MOTD` text NOT NULL,
  `Players` int(11) NOT NULL default '0',
  `Online` int(11) NOT NULL default '0',
  PRIMARY KEY  (`FKey`)
) TYPE=MyISAM;



#
# Table structure for table 'shops'
#

DROP TABLE IF EXISTS `shops`;
CREATE TABLE `shops` (
  `FKey` int(11) NOT NULL auto_increment,
  `Name` tinytext NOT NULL,
  `JoinSay` tinytext NOT NULL,
  `LeaveSay` tinytext NOT NULL,
  `GiveItem1` int(11) NOT NULL default '0',
  `GiveValue1` int(11) NOT NULL default '0',
  `GetItem1` int(11) NOT NULL default '0',
  `GetValue1` int(11) NOT NULL default '0',
  `GiveItem2` int(11) NOT NULL default '0',
  `GiveValue2` int(11) NOT NULL default '0',
  `GetItem2` int(11) NOT NULL default '0',
  `GetValue2` int(11) NOT NULL default '0',
  `GiveItem3` int(11) NOT NULL default '0',
  `GiveValue3` int(11) NOT NULL default '0',
  `GetItem3` int(11) NOT NULL default '0',
  `GetValue3` int(11) NOT NULL default '0',
  `GiveItem4` int(11) NOT NULL default '0',
  `GiveValue4` int(11) NOT NULL default '0',
  `GetItem4` int(11) NOT NULL default '0',
  `GetValue4` int(11) NOT NULL default '0',
  `GiveItem5` int(11) NOT NULL default '0',
  `GiveValue5` int(11) NOT NULL default '0',
  `GetItem5` int(11) NOT NULL default '0',
  `GetValue5` int(11) NOT NULL default '0',
  `GiveItem6` int(11) NOT NULL default '0',
  `GiveValue6` int(11) NOT NULL default '0',
  `GetItem6` int(11) NOT NULL default '0',
  `GetValue6` int(11) NOT NULL default '0',
  `GiveItem7` int(11) NOT NULL default '0',
  `GiveValue7` int(11) NOT NULL default '0',
  `GetItem7` int(11) NOT NULL default '0',
  `GetValue7` int(11) NOT NULL default '0',
  `GiveItem8` int(11) NOT NULL default '0',
  `GiveValue8` int(11) NOT NULL default '0',
  `GetItem8` int(11) NOT NULL default '0',
  `GetValue8` int(11) NOT NULL default '0',
  `FixesItems` tinyint(11) unsigned default '0',
  PRIMARY KEY  (`FKey`)
) TYPE=MyISAM;



#
# Table structure for table 'spells'
#

DROP TABLE IF EXISTS `spells`;
CREATE TABLE `spells` (
  `FKey` int(11) NOT NULL auto_increment,
  `Name` tinytext NOT NULL,
  `ClassReq` int(11) NOT NULL default '0',
  `Type` int(11) NOT NULL default '0',
  `Data1` int(11) NOT NULL default '0',
  `Data2` int(11) NOT NULL default '0',
  `Data3` int(11) NOT NULL default '0',
  PRIMARY KEY  (`FKey`)
) TYPE=MyISAM;

