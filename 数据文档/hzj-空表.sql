/*
Navicat MySQL Data Transfer

Source Server         : localhost_3306
Source Server Version : 50713
Source Host           : localhost:3306
Source Database       : hzj

Target Server Type    : MYSQL
Target Server Version : 50713
File Encoding         : 65001

Date: 2016-12-15 10:52:51
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
) ENGINE=InnoDB AUTO_INCREMENT=14368 DEFAULT CHARSET=utf8;

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
) ENGINE=InnoDB AUTO_INCREMENT=329 DEFAULT CHARSET=utf8;

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
  `Total_hold` float NOT NULL DEFAULT '0',
  `Diff_hold` float NOT NULL DEFAULT '0',
  `History` varchar(50) NOT NULL,
  `Version` varchar(20) NOT NULL,
  PRIMARY KEY (`ID`),
  UNIQUE KEY `ExcelSignal` (`ExcelSignal`)
) ENGINE=InnoDB AUTO_INCREMENT=5 DEFAULT CHARSET=utf8;

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
) ENGINE=InnoDB AUTO_INCREMENT=10 DEFAULT CHARSET=utf8;

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
) ENGINE=InnoDB AUTO_INCREMENT=1639 DEFAULT CHARSET=utf8;

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
) ENGINE=InnoDB AUTO_INCREMENT=19 DEFAULT CHARSET=utf8;

-- ----------------------------
-- Records of keytable
-- ----------------------------
