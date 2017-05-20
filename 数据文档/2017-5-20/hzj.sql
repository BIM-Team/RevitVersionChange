/*
Navicat MySQL Data Transfer

Source Server         : localhost_3306
Source Server Version : 50713
Source Host           : localhost:3306
Source Database       : hzj

Target Server Type    : MYSQL
Target Server Version : 50713
File Encoding         : 65001

Date: 2017-05-20 18:22:46
*/

SET FOREIGN_KEY_CHECKS=0;

-- ----------------------------
-- Table structure for `drawtable`
-- ----------------------------
DROP TABLE IF EXISTS `drawtable`;
CREATE TABLE `drawtable` (
  `ID` int(11) NOT NULL AUTO_INCREMENT,
  `Entity_ID` int(11) NOT NULL,
  `Date` datetime NOT NULL,
  `EntityMaxValue` float NOT NULL,
  `EntityMidValue` float NOT NULL,
  `EntityMinValue` float NOT NULL,
  `Detail` text NOT NULL,
  `Version` varchar(20) NOT NULL,
  PRIMARY KEY (`ID`),
  KEY `dt_EntityId` (`Entity_ID`)
) ENGINE=InnoDB AUTO_INCREMENT=502 DEFAULT CHARSET=utf8;

-- ----------------------------
-- Records of drawtable
-- ----------------------------

-- ----------------------------
-- Table structure for `entitytable`
-- ----------------------------
DROP TABLE IF EXISTS `entitytable`;
CREATE TABLE `entitytable` (
  `ID` int(11) NOT NULL AUTO_INCREMENT,
  `ExcelSignal` varchar(20) NOT NULL,
  `EntityName` varchar(20) NOT NULL,
  `Remark` varchar(30) DEFAULT NULL,
  `Version` varchar(20) NOT NULL,
  PRIMARY KEY (`ID`),
  KEY `et_ExcelId` (`ExcelSignal`)
) ENGINE=InnoDB AUTO_INCREMENT=27 DEFAULT CHARSET=utf8;

-- ----------------------------
-- Records of entitytable
-- ----------------------------

-- ----------------------------
-- Table structure for `exceltable`
-- ----------------------------
DROP TABLE IF EXISTS `exceltable`;
CREATE TABLE `exceltable` (
  `ID` int(11) NOT NULL AUTO_INCREMENT,
  `CurrentFile` varchar(30) NOT NULL,
  `ExcelSignal` varchar(20) NOT NULL,
  `IsInfo` tinyint(1) NOT NULL,
  `Total_hold` varchar(20) NOT NULL DEFAULT '0',
  `Total_operator` enum('>=','<=','<','IN','OUT','>') DEFAULT '>',
  `Diff_hold` varchar(20) NOT NULL DEFAULT '0',
  `Diff_operator` enum('>=','<=','>','IN','OUT','<') DEFAULT '>',
  `History` varchar(50) NOT NULL,
  `Version` varchar(20) NOT NULL,
  PRIMARY KEY (`ID`),
  UNIQUE KEY `ExcelSignal` (`ExcelSignal`)
) ENGINE=InnoDB AUTO_INCREMENT=2 DEFAULT CHARSET=utf8;

-- ----------------------------
-- Records of exceltable
-- ----------------------------

-- ----------------------------
-- Table structure for `grouptable`
-- ----------------------------
DROP TABLE IF EXISTS `grouptable`;
CREATE TABLE `grouptable` (
  `ID` int(11) NOT NULL AUTO_INCREMENT,
  `ExcelSignal` varchar(20) NOT NULL,
  `GroupName` varchar(20) NOT NULL,
  PRIMARY KEY (`ID`),
  KEY `gt_ExcelId` (`ExcelSignal`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- ----------------------------
-- Records of grouptable
-- ----------------------------

-- ----------------------------
-- Table structure for `infotable`
-- ----------------------------
DROP TABLE IF EXISTS `infotable`;
CREATE TABLE `infotable` (
  `ID` int(11) NOT NULL AUTO_INCREMENT,
  `Key_ID` int(11) NOT NULL,
  `Entity_ID` int(11) NOT NULL,
  `Value` varchar(20) NOT NULL,
  `Version` varchar(20) NOT NULL,
  PRIMARY KEY (`ID`),
  KEY `it_keyId` (`Key_ID`),
  KEY `it_EntityId` (`Entity_ID`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- ----------------------------
-- Records of infotable
-- ----------------------------

-- ----------------------------
-- Table structure for `keytable`
-- ----------------------------
DROP TABLE IF EXISTS `keytable`;
CREATE TABLE `keytable` (
  `ID` int(11) NOT NULL AUTO_INCREMENT,
  `ExcelSignal` varchar(20) NOT NULL,
  `Group_ID` int(11) DEFAULT NULL,
  `KeyName` varchar(20) NOT NULL,
  `Odr` int(11) NOT NULL DEFAULT '0',
  `Version` varchar(20) NOT NULL,
  PRIMARY KEY (`ID`),
  KEY `kt_ExcelId` (`ExcelSignal`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- ----------------------------
-- Records of keytable
-- ----------------------------
