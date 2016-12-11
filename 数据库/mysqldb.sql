drop database if exists hzj;
create database hzj;
use hzj;
CREATE table ExcelTable(ID INT PRIMARY KEY AUTO_INCREMENT,
                   TableDesc VARCHAR(30) NOT NULL,
				   ExcelSignal VARCHAR(10) UNIQUE,
				   IsInfo Boolean,
				   Total_hold double,
				   Diff_hold double,
                   Remark VARCHAR(100)
                   ) DEFAULT CHARSET = utf8;
CREATE TABLE KeyTable(ID INT PRIMARY KEY AUTO_INCREMENT,
                      Excel_ID INT NOT NULL,
					  Group_ID INT ,
                      KeyName VARCHAR(20) NOT NULL,
					  CONSTRAINT kt_ExcelId FOREIGN KEY (Excel_ID) REFERENCES ExcelTable(ID) on delete cascade					  
                      ) DEFAULT CHARSET = utf8;
CREATE TABLE EntityTable(ID INT PRIMARY KEY AUTO_INCREMENT,
                       Excel_ID INT NOT NULL,
                       EntityName VARCHAR(20) NOT NULL,
                       Remark VARCHAR(100),
                       CONSTRAINT et_ExcelId FOREIGN KEY (Excel_ID) REFERENCES ExcelTable(ID) on delete cascade
                       ) DEFAULT CHARSET = utf8;
CREATE TABLE GroupTable(ID INT PRIMARY KEY AUTO_INCREMENT,
                        Excel_ID INT NOT NULL,
                        GroupName VARCHAR(20) NOT NULL,
						Remark VARCHAR(100),
                        CONSTRAINT gt_ExcelId FOREIGN KEY (Excel_ID) REFERENCES ExcelTable(ID)                        
                        ) DEFAULT CHARSET = utf8;
CREATE TABLE InfoTable(ID INT PRIMARY KEY AUTO_INCREMENT,
                       Key_ID INT NOT NULL,                                             
                       Entity_ID INT NOT NULL,
                       Value VARCHAR(20),
                       CONSTRAINT it_keyId FOREIGN KEY (Key_ID) REFERENCES KeyTable(ID) on delete cascade,
                       CONSTRAINT it_EntityId FOREIGN KEY (Entity_ID) REFERENCES EntityTable(ID) on delete cascade
                       )DEFAULT CHARSET = utf8;
CREATE TABLE DrawDataTable(ID INT PRIMARY KEY AUTO_INCREMENT,
                         Excel_ID INT NOT NULL,
                         Entity_ID INT NOT NULL,
                         Date DATETIME,
						 EntityMaxValue double,
						 EntityMidValue double,
						 EntityMinValue double,
						 Detail MEDIUMTEXT,
                         CONSTRAINT dt_ExcelId FOREIGN KEY (Excel_ID) REFERENCES ExcelTable(ID) on delete cascade, 
                         CONSTRAINT dt_EntityId FOREIGN KEY (Entity_ID) REFERENCES EntityTable(ID) on delete cascade
                       )DEFAULT CHARSET = utf8;
