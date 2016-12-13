﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.SQLite;
using MySql.Data.MySqlClient;
using Revit.Addin.RevitTooltip.Intface;
using Revit.Addin.RevitTooltip.Dto;
using System.Windows;
using System.IO;
using System.Text;

namespace Revit.Addin.RevitTooltip.Impl
{
    public class SQLiteHelper : ISQLiteHelper, IDisposable
    {
        /// <summary>
        /// 保存当前的一个实例
        /// </summary>
        private static SQLiteHelper _sqliteHelper = null;
        /// <summary>
        /// 获取实例
        /// </summary>
        public static SQLiteHelper Instance { get { return _sqliteHelper; } }

        //连接
        private SQLiteConnection conn;


        private string dbName = null;
        private string dbPath = null;

        /// <summary>
        /// 构造函数
        /// </summary>
        public SQLiteHelper()
        {
            this.dbName = "SqliteDB.db";
            this.dbPath = "E:\\revit数据文档";
            if (!Directory.Exists(this.dbPath))
            {
                Directory.CreateDirectory(this.dbPath);
            }
            Directory.SetCurrentDirectory(this.dbPath);
            this.conn = new SQLiteConnection("Data Source = " + dbName);
            _sqliteHelper = this;
        }
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="settings"></param>
        internal SQLiteHelper(RevitTooltip settings)
        {
            this.dbName = settings.SqliteFileName;
            this.dbPath = settings.SqliteFilePath;
            if (!Directory.Exists(this.dbPath))
            {
                Directory.CreateDirectory(this.dbPath);
            }
            Directory.SetCurrentDirectory(this.dbPath);
            this.conn = new SQLiteConnection("Data Source=" + dbName);
            _sqliteHelper = this;
        }
        /// <summary>
        /// 销毁当前对象
        /// </summary>
        public void Dispose()
        {
            if (conn.State != ConnectionState.Closed)
            {
                conn.Close();
            }
            conn.Dispose();
            GC.SuppressFinalize(this);
        }
        /// <summary>
        /// 从Mysql数据库中导入数据
        /// 如果本地有同名的Sqlite数据库文件，先删除，再创建
        /// </summary>
        public bool LoadDataToSqlite()
        {
            bool result = false;
            if (conn.State != ConnectionState.Closed)
            {
                conn.Close();
            }
            string file_path = Path.Combine(this.dbPath, this.dbName);
            if (File.Exists(file_path))
            {
                File.Delete(file_path);
            }
            if (conn.State != ConnectionState.Open)
            {
                conn.Open();
            }
            using (DbTransaction transaction = conn.BeginTransaction())
            {
                try
                {
                    CreateDB();
                    InsertDBfromMysql();
                    transaction.Commit();
                    result = true;
                }
                catch (Exception)
                {
                    throw;
                }
                finally
                {
                    conn.Close();
                }
            }
            return result;
        }
        //创建SQLite数据库文件
        public void CreateDB()
        {
            //如果连接没有打开则打开连接
            if (conn.State != ConnectionState.Open)
            {
                conn.Open();
            }
            using (SQLiteCommand command = new SQLiteCommand(conn))
            {
                if (!isExist("ExcelTable", "table"))
                {
                    command.CommandText = "CREATE TABLE ExcelTable(ID integer NOT NULL PRIMARY KEY AUTOINCREMENT,CurrentFile VARCHAR(30) NOT NULL,ExcelSignal VARCHAR(20) UNIQUE, IsInfo BOOLEAN NOT NULL, Total_hold float NOT NULL default 0, Diff_hold float NOT NULL default 0, History VARCHAR(100) NOT NULL )";
                    command.ExecuteNonQuery();
                }
                if (!isExist("KeyTable", "table"))
                {
                    command.CommandText = "CREATE TABLE KeyTable(ID integer PRIMARY KEY AUTOINCREMENT, ExcelSignal VARCHAR(20) NOT NULL, Group_ID integer , KeyName VARCHAR(20) NOT NULL,Odr integer NOT NULL default 0)";
                    command.ExecuteNonQuery();
                }
                if (!isExist("EntityTable", "table"))
                {
                    command.CommandText = "CREATE TABLE EntityTable(ID integer PRIMARY KEY AUTOINCREMENT, ExcelSignal VARCHAR(20) NOT NULL, EntityName VARCHAR(20) NOT NULL, Remark VARCHAR(100))";
                    command.ExecuteNonQuery();
                }
                if (!isExist("GroupTable", "table"))
                {
                    command.CommandText = "CREATE TABLE GroupTable(ID integer PRIMARY KEY AUTOINCREMENT, ExcelSignal VARCHAR(20) NOT NULL, GroupName VARCHAR(20) NOT NULL)";
                    command.ExecuteNonQuery();
                }
                if (!isExist("InfoTable", "table"))
                {
                    command.CommandText = "CREATE TABLE InfoTable(ID integer PRIMARY KEY AUTOINCREMENT, Key_ID integer NOT NULL REFERENCES KeyTable(ID) ON DELETE CASCADE ON UPDATE NO ACTION, Entity_ID integer NOT NULL REFERENCES EntityTable(ID) ON DELETE CASCADE ON UPDATE NO ACTION, Value VARCHAR(20) NOT NULL)";
                    command.ExecuteNonQuery();
                }
                if (!isExist("DrawTable", "table"))
                {
                    command.CommandText = "CREATE TABLE DrawTable(ID integer PRIMARY KEY AUTOINCREMENT,  Entity_ID integer NOT NULL REFERENCES EntityTable(ID) ON DELETE CASCADE ON UPDATE NO ACTION, Date VARCHAR(20) NOT NULL, EntityMaxValue float NOT NULL, EntityMidValue float NOT NULL, EntityMinValue float NOT NULL, Detail TEXT NOT NULL )";
                    command.ExecuteNonQuery();
                }
            }
        }

        //插入数据
        private void InsertDBfromMysql()
        {
            if (!MysqlUtil.Instance.IsReady)
            {
                return;
            }
            string conn_string = MysqlUtil.Instance.ConnectionMessage;
            MySqlConnection mysql_conn = new MySqlConnection(conn_string);
            MySqlDataReader mysql_reader = null;
            try
            {
                //sqlite
                if (conn.State != ConnectionState.Open)
                {
                    conn.Open();
                }
                SQLiteCommand sqlite_command = new SQLiteCommand(conn);
                //mysql
                mysql_conn.Open();
                string select_sql = "Select ID,CurrentFile,ExcelSignal,IsInfo,Total_hold,Diff_hold,History From ExcelTable";
                MySqlCommand mysql_command = new MySqlCommand(select_sql, mysql_conn);
                //ExcelTable
                mysql_reader = mysql_command.ExecuteReader();
                if (mysql_reader.HasRows)
                {
                    while (mysql_reader.Read())
                    {
                        sqlite_command.CommandText = string.Format("INSERT OR IGNORE INTO ExcelTable(ID, CurrentFile, ExcelSignal, IsInfo, Total_hold, Diff_hold, History)values({0},'{1}','{2}',{3},{4},{5},'{6}')",
                            mysql_reader.GetInt32(0), mysql_reader.GetString(1), mysql_reader.GetString(2), mysql_reader.GetInt32(3), mysql_reader.GetFloat(4), mysql_reader.GetFloat(5), mysql_reader.GetString(6));
                        sqlite_command.ExecuteNonQuery();
                    }
                }
                mysql_reader.Close();
                //KeyTable
                mysql_command.CommandText = "Select ID,ExcelSignal,Group_ID,KeyName,Odr From KeyTable";
                mysql_reader = mysql_command.ExecuteReader();
                if (mysql_reader.HasRows)
                {
                    while (mysql_reader.Read())
                    {
                        sqlite_command.CommandText = string.Format("INSERT OR IGNORE INTO KeyTable(ID,ExcelSignal,Group_ID,KeyName,Odr) Values({0},'{1}',{2},'{3}',{4})",
                           mysql_reader.GetInt32(0), mysql_reader.GetString(1), mysql_reader.IsDBNull(2) ? "null" : mysql_reader.GetInt32(2).ToString(), mysql_reader.GetString(3), mysql_reader.GetInt32(4));
                        sqlite_command.ExecuteNonQuery();
                    }
                }
                mysql_reader.Close();
                //EntityTable
                mysql_command.CommandText = "Select ID, ExcelSignal,EntityName,Remark From EntityTable";
                mysql_reader = mysql_command.ExecuteReader();
                if (mysql_reader.HasRows)
                {
                    while (mysql_reader.Read())
                    {
                        sqlite_command.CommandText = string.Format("INSERT OR IGNORE INTO EntityTable(ID, ExcelSignal,EntityName,Remark) Values({0},'{1}','{2}','{3}')",
                           mysql_reader.GetInt32(0), mysql_reader.GetString(1), mysql_reader.GetString(2), mysql_reader.IsDBNull(3) ? "null" : mysql_reader.GetString(3));
                        sqlite_command.ExecuteNonQuery();
                    }
                }
                mysql_reader.Close();
                //GroupTable
                mysql_command.CommandText = "Select ID,ExcelSignal,GroupName From GroupTable";
                mysql_reader = mysql_command.ExecuteReader();
                if (mysql_reader.HasRows)
                {
                    while (mysql_reader.Read())
                    {
                        sqlite_command.CommandText = string.Format("INSERT OR IGNORE INTO GroupTable(ID,ExcelSignal,GroupName) Values({0},'{1}','{2}')",
                           mysql_reader.GetInt32(0), mysql_reader.GetString(1), mysql_reader.GetString(2));
                        sqlite_command.ExecuteNonQuery();
                    }
                }
                mysql_reader.Close();
                //InfoTable
                mysql_command.CommandText = "Select ID, Key_ID,Entity_ID,Value From InfoTable";
                mysql_reader = mysql_command.ExecuteReader();
                if (mysql_reader.HasRows)
                {
                    while (mysql_reader.Read())
                    {
                        sqlite_command.CommandText = string.Format("INSERT OR IGNORE INTO InfoTable(ID, Key_ID,Entity_ID,Value) Values({0},{1},{2},'{3}')",
                           mysql_reader.GetInt32(0), mysql_reader.GetInt32(1), mysql_reader.GetInt32(2), mysql_reader.GetString(3));
                        sqlite_command.ExecuteNonQuery();
                    }
                }
                mysql_reader.Close();
                //DrawTable
                mysql_command.CommandText = "Select ID,Entity_ID, Date ,EntityMaxValue,EntityMidValue,EntityMinValue,Detail From DrawTable ";
                mysql_reader = mysql_command.ExecuteReader();
                if (mysql_reader.HasRows)
                {
                    while (mysql_reader.Read())
                    {
                        sqlite_command.CommandText = string.Format("INSERT OR IGNORE INTO DrawTable(ID,Entity_ID,Date ,EntityMaxValue,EntityMidValue,EntityMinValue,Detail) Values({0},{1},'{2}',{3},{4},{5},'{6}')",
                            mysql_reader.GetInt32(0), mysql_reader.GetInt32(1), mysql_reader.GetDateTime(2).ToString("yyyy-MM-dd HH:mm:ss"), mysql_reader.GetFloat(3), mysql_reader.GetFloat(4), mysql_reader.GetFloat(5), mysql_reader.GetString(6));
                        sqlite_command.ExecuteNonQuery();
                    }
                }
                mysql_reader.Close();
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                if (!mysql_reader.IsClosed)
                {
                    mysql_reader.Close();
                }
                mysql_conn.Close();
                mysql_conn.Dispose();
            }
        }

        /// <summary>
        /// 判断Sqlite数据库中表是否存在
        /// </summary>
        /// <param name="table"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        private bool isExist(string table, string type)
        {
            bool flag = false;
            //连接没有打开
            if (this.conn.State != ConnectionState.Open)
            {
                this.conn.Open();
            }
            string sql = String.Format("select count(*) from sqlite_master where type='{0}' and name = '{1}' ", type, table);
            using (SQLiteCommand command = new SQLiteCommand(sql, this.conn))
            {
                if (Convert.ToInt32(command.ExecuteScalar()) > 0)
                    flag = true;
            }
            return flag;
        }

        //*****************************插入Excel数据***************************

        /// <summary>
        /// 插入SheetInfo
        /// </summary>
        public void InsertSheetInfo(SheetInfo sheetInfo)
        {
            if (sheetInfo == null)
            {
                throw new Exception("无效的传入参数");
            }
            if (!File.Exists(Path.Combine(dbPath, dbName)))
            {
                CreateDB();
            }
            if (sheetInfo.Tag)
            {
                InsertInfoData(sheetInfo);
            }
            else
            {
                InsertDrawData(sheetInfo);
            }
        }
        /// <summary>
        /// 插入基础数据表
        /// </summary>
        /// <param name="sheetInfo"></param>
        private void InsertInfoData(SheetInfo sheetInfo)
        {
            //是否已经处理过
            bool hasDone = false;
            string signal = sheetInfo.ExcelTableData.Signal;
            string reset_auto_increment = "DELETE FROM sqlite_sequence";
            SQLiteTransaction tran = null;
            try
            {
                if (conn.State != ConnectionState.Open)
                {
                    conn.Open();
                }
                //事务开始
                tran = conn.BeginTransaction();
                //重置自增
                new SQLiteCommand(reset_auto_increment, conn, tran).ExecuteNonQuery();
                //插入到ExcelTable表,并返回ID
                if (sheetInfo.ExcelTableData == null)
                {
                    throw new Exception("无效的插入数据");
                }
                //判断是否该Signal已存在
                ExcelTable exist = SelectExcelTable(signal);
                if (exist == null)
                {
                    //插入ExcelTable
                    new SQLiteCommand(string.Format("insert into ExcelTable (CurrentFile,ExcelSignal,IsInfo,History) values ('{0}', '{1}', {2},'{3}')",
                         sheetInfo.ExcelTableData.CurrentFile, signal, sheetInfo.Tag?1:0, sheetInfo.ExcelTableData.CurrentFile), tran.Connection, tran).ExecuteNonQuery();
                    exist = SelectExcelTable(signal);
                    //插入表结构KeyNames
                    new SQLiteCommand(InsertIntoKeyTable(sheetInfo.KeyNames, sheetInfo.ExcelTableData.Signal), tran.Connection, tran).ExecuteNonQuery();
                }
                else
                {
                    string[] his = exist.History.Split(';');
                    string currentFile = sheetInfo.ExcelTableData.CurrentFile;
                    //判断是否已经做过处理
                    foreach (string s in his)
                    {
                        if (currentFile.Equals(s))
                        {
                            hasDone = true;
                            break;
                        }
                    }
                    //对于没有处理过的，添加到History中
                    if (!hasDone)
                    {
                        string history = exist.History + ";" + sheetInfo.ExcelTableData.CurrentFile;
                        //更新已有的数据表
                        new SQLiteCommand(string.Format("update ExcelTable set CurrentFile='{0}',History='{1}' where ExcelSignal='{2}'",
                           sheetInfo.ExcelTableData.CurrentFile, history, signal), tran.Connection, tran).ExecuteNonQuery();
                    }
                }
                if (!hasDone)
                {
                    List<CKeyName> inMysqls = SelectKeyNames(signal, conn);
                    if (inMysqls == null || inMysqls.Count == 0)
                    {
                        throw new Exception("数据异常");
                    }
                    //构造Map
                    Dictionary<string, int> KeyMap = new Dictionary<string, int>();
                    foreach (CKeyName one in inMysqls)
                    {
                        KeyMap.Add(one.KeyName, one.Id);
                    }
                    InsertInfoTable(sheetInfo, KeyMap, tran);
                }
                tran.Commit();    //事务提交
            }
            catch (Exception e)
            {
                tran.Rollback();    // 事务回滚
                throw new Exception("事务操作出错，系统信息：" + e.Message);
            }
            finally
            {
                conn.Close();
            }
        }

        /// <summary>
        /// 插入EntityInfo数据
        /// </summary>
        /// <param name="sheetInfo"></param>
        /// <param name="KeyMap"></param>
        /// <param name="command"></param>
        private void InsertInfoTable(SheetInfo sheetInfo, Dictionary<string, int> KeyMap, SQLiteTransaction tran)
        {
            if (sheetInfo.InfoRows == null || sheetInfo.InfoRows.Count == 0)
            {

                throw new Exception("无效的数据");
            }
            List<InfoEntityData> rows = sheetInfo.InfoRows;
            string signal = sheetInfo.ExcelTableData.Signal;

            foreach (InfoEntityData one in rows)
            {
                string sql = string.Format("insert into EntityTable(ExcelSignal,EntityName) values ('{0}','{1}')", signal, one.EntityName);
                //插入Entity
                new SQLiteCommand(sql, tran.Connection, tran).ExecuteNonQuery();
                CEntityName entity = selectEntity(one.EntityName);
                Dictionary<string, string> data = one.Data;
                StringBuilder buider = null;
                if (data != null && data.Count != 0)
                {
                    buider = new StringBuilder("insert into InfoTable(Key_ID,Entity_ID,Value) values");
                }
                foreach (string s in data.Keys)
                {
                    buider.AppendFormat(" ({0},{1},'{2}'),", KeyMap[s], entity.Id, data[s]);
                }
                buider.Remove(buider.Length - 1, 1);
                new SQLiteCommand(buider.ToString(), tran.Connection, tran).ExecuteNonQuery();
            }
        }

        /// <summary>
        /// 查询CEntityName
        /// </summary>
        /// <param name="entityName"></param>
        /// <param name="command"></param>
        /// <param name="err"></param>
        /// <returns></returns>
        private CEntityName selectEntity(string entityName)
        {
            CEntityName result = null;
                string sql = string.Format("select ID,EntityName from EntityTable where EntityName='{0}'", entityName);
                if (conn.State != ConnectionState.Open)
                {
                conn.Open();
                }
                SQLiteCommand command = new SQLiteCommand(sql, conn);
                SQLiteDataReader reader = command.ExecuteReader();
                try
                {
                    while (reader.Read())
                    {
                        result = new CEntityName();
                        result.Id = reader.GetInt32(0);
                        result.EntityName = reader.GetString(1);
                    }

                }
                catch (Exception)
                {
                    throw;
                }
                finally
                {
                    reader.Close();
                }
            
            return result;
        }
        /// <summary>
        /// 查询某signal的所有KeyName
        /// </summary>
        /// <param name="Signal"></param>
        /// <param name="mycom"></param>
        /// <returns></returns>
        private List<CKeyName> SelectKeyNames(string Signal, SQLiteConnection OpenedConn)
        {
            string sql = string.Format("select ID,KeyName,Odr from KeyTable where ExcelSignal='{0}' order by Odr", Signal);
            List<CKeyName> keyNames = null;
            //需要关闭
            SQLiteDataReader reader = new SQLiteCommand(sql, OpenedConn).ExecuteReader();
            try
            {
                if (reader.HasRows)
                {
                    keyNames = new List<CKeyName>();
                    while (reader.Read())
                    {
                        CKeyName one = new CKeyName();
                        one.Id = reader.GetInt32(0);
                        one.KeyName = reader.GetString(1);
                        one.Odr = reader.GetInt32(2);
                        keyNames.Add(one);
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                reader.Close();
            }
            return keyNames;
        }


        /// <summary>
        /// 获取KeyTable的SQL
        /// </summary>
        /// <param name="KeyNames"></param>
        /// <param name="Signal"></param>
        /// <returns></returns>
        private string InsertIntoKeyTable(List<string> KeyNames, string Signal)
        {
            StringBuilder sql = new StringBuilder("insert into KeyTable(ExcelSignal,KeyName) values");
            foreach (string one in KeyNames)
            {
                sql.AppendFormat(" ('{0}','{1}'),", Signal, one);
            }
            sql.Remove(sql.Length - 1, 1);
            return sql.ToString();
        }
        /// <summary>
        /// 查询现有的ExcelTable
        /// </summary>
        /// <param name="signal"></param>
        /// <param name="OpenedConn"></param>
        /// <returns></returns>
        private ExcelTable SelectExcelTable(String signal)
        {
            if (conn.State != ConnectionState.Open)
            {
                conn.Open();
            }
            string sql = String.Format("select ID,CurrentFile,ExcelSignal,Total_hold,Diff_hold,History from ExcelTable where ExcelSignal = '{0}'", signal);
            SQLiteDataReader reader = new SQLiteCommand(sql, conn).ExecuteReader();
            ExcelTable result = null;
            try
            {
                while (reader.Read())
                {
                    result = new ExcelTable();
                    result.Id = reader.GetInt32(0);
                    result.CurrentFile = reader.GetString(1);
                    result.Signal = reader.GetString(2);
                    result.Total_hold = reader.GetFloat(3);
                    result.Diff_hold = reader.GetFloat(4);
                    result.History = reader.GetString(5);
                }
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                reader.Close();
            }
            return result;
        }
        /// <summary>
        /// 插入绘图数据表
        /// </summary>
        /// <param name="sheetInfo"></param>
        private void InsertDrawData(SheetInfo sheetInfo)
        {
            //是否已经处理过
            bool hasDone = false;
            string signal = sheetInfo.ExcelTableData.Signal;
            //该连接仅用于修改数据库
            SQLiteTransaction tran = null;
            try
            {
                if (conn.State != ConnectionState.Open)
                {
                    conn.Open();
                }
                //事务开始
                tran = conn.BeginTransaction();
                string reset_auto_increment = "DELETE FROM sqlite_sequence";
                //主命令
                SQLiteCommand command = new SQLiteCommand(reset_auto_increment, conn, tran);
                //重置自增
                command.ExecuteNonQuery();
                //插入到ExcelTable表,并返回ID
                if (sheetInfo.ExcelTableData == null)
                {
                    throw new Exception("无效的插入数据");
                }
                //判断是否该Signal已存在
                ExcelTable exist = SelectExcelTable(signal);
                if (exist == null)
                {
                    command.CommandText = string.Format("insert into ExcelTable (CurrentFile,ExcelSignal,IsInfo,History) values ('{0}', '{1}', {2},'{3}')",
                          sheetInfo.ExcelTableData.CurrentFile, signal, sheetInfo.Tag?1:0, sheetInfo.ExcelTableData.CurrentFile);
                    //插入新的数据表
                    command.ExecuteNonQuery();
                }
                else
                {
                    string[] his = exist.History.Split(';');
                    string currentFile = sheetInfo.ExcelTableData.CurrentFile;
                    foreach (string s in his)
                    {
                        if (currentFile.Equals(s))
                        {
                            hasDone = true;
                            break;
                        }
                    }
                    if (!hasDone)
                    {
                        string history = exist.History + ";" + sheetInfo.ExcelTableData.CurrentFile;
                        command.CommandText = string.Format("update ExcelTable set CurrentFile='{0}',History='{1}' where ExcelSignal='{3}'",
                        sheetInfo.ExcelTableData.CurrentFile, history, signal);
                        //更新已有的数据表
                        command.ExecuteNonQuery();
                    }
                }
                if (!hasDone)
                {
                    InsertDrawDataTable(sheetInfo, tran);
                }
                tran.Commit();    //事务提交
            }
            catch (Exception e)
            {
                tran.Rollback();    // 事务回滚
                throw new Exception("事务操作出错，系统信息：" + e.Message);
            }
            finally
            {
                conn.Close();
            }
        }
        /// <summary>
        /// 插入DrawDataTable
        /// </summary>
        /// <param name="sheetInfo"></param>
        /// <param name="command"></param>
        private void InsertDrawDataTable(SheetInfo sheetInfo, SQLiteTransaction tran)
        {
            List<DrawEntityData> rows = sheetInfo.DrawRows;
            string signal = sheetInfo.ExcelTableData.Signal;
            foreach (DrawEntityData one in rows)
            {
                CEntityName entity = selectEntity(one.EntityName);
                DateTime? maxDate = null;
                if (entity == null)
                {
                    string sql = string.Format("insert into EntityTable(ExcelSignal,EntityName) values ('{0}','{1}')", signal, one.EntityName);
                    //插入Entity
                    new SQLiteCommand(sql, tran.Connection, tran).ExecuteNonQuery();
                    entity = selectEntity(one.EntityName);
                }
                else
                {
                    string sql = string.Format("select Max(Date) from DrawTable where Entity_ID={0}", entity.Id);
                    maxDate = Convert.ToDateTime(new SQLiteCommand(sql, tran.Connection).ExecuteScalar());
                }
                List<DrawData> data = one.Data;
                StringBuilder buider = null;
                if (data != null && data.Count != 0)
                {
                    buider = new StringBuilder("insert into DrawTable(Entity_ID,Date,EntityMaxValue,EntityMidValue,EntityMinValue,Detail) values");
                }
                bool hasValue = false;
                foreach (DrawData p in data)
                {
                    if (maxDate == null || p.Date > maxDate)
                    {
                        hasValue = true;
                        buider.AppendFormat(" ({0},'{1}','{2}','{3}','{4}','{5}'),", entity.Id, p.Date.ToString("yyyy-MM-dd HH:mm:ss"), p.MaxValue, p.MidValue, p.MinValue, p.Detail);
                    }
                }
                if (hasValue) {
                buider.Remove(buider.Length - 1, 1);
                new SQLiteCommand(buider.ToString(), tran.Connection, tran).ExecuteNonQuery();
                }
            }

        }

        //*****************************查询功能*******************************

        //查询单个CX某日期的测斜数据，返回键值对<属性，数据>
        public Dictionary<string, float> SelectOneDateData(string Entity, DateTime date)
        {
            Dictionary<string, float> data = new Dictionary<string, float>();

            String sql = String.Format("select it.DATE,cast(ft.COLUMNNAME as DECIMAL(4,2)),cast(it.VALUE as DECIMAL(5,2)) from InclinationTable it,FrameTable ft,EntityTable et "
                          + " where ft.ID_TYPE=et.ID_TYPE and it.ID_ENTITY=et.ID and it.ID_FRAME=ft.ID and et.ENTITY= '{0}' and it.DATE = '{1}' ", Entity, date);

            if (conn.State != ConnectionState.Open)
            {
                conn.Open();
            }
            //判断是否创建该查询的表（一定要先打开数据库）
            if (!isExist("EntityTable", "table"))
            {
                MessageBox.Show("本地数据库不存在，建议更新本地数据库！");
                return data;
            }
            using (SQLiteCommand command = new SQLiteCommand(sql, conn))  //建立执行命令语句对象
            {
                SQLiteDataReader reader = command.ExecuteReader();
                try
                {
                    while (reader.Read())
                    {
                        if (reader.HasRows)
                        {
                            data.Add(reader.GetString(1), reader.GetFloat(2));
                        }
                    }
                    return data;
                }
                catch (Exception e)
                {
                    throw e;
                }
                finally
                {
                    reader.Close();  //关闭
                }
            }
        }

        //查询单个CX连续的测量结果（每个测量日期最大值），返回键值对<日期，数据>
        public Dictionary<string, float> SelectOneCXData(string entity)
        {
            Dictionary<string, float> data = new Dictionary<string, float>();

            string sql = String.Format(" select ID,DATE,VALUE from CXTable where ENTITY = '{0}' order by ID ", entity);
            // string sql = String.Format(" select ID,DATE,MAX(VALUE) from CXTable where ENTITY = '{0}' group by DATE order by ID ", entity);

            if (conn.State != ConnectionState.Open)
            {
                conn.Open();
            }
            //判断是否创建该查询的表（一定要先打开数据库）
            if (!isExist("CXTable", "table"))
            {
                MessageBox.Show("本地数据库不存在，建议更新本地数据库！");
                return data;
            }
            using (SQLiteCommand command = new SQLiteCommand(sql, conn))  //建立执行命令语句对象
            {
                SQLiteDataReader reader = command.ExecuteReader();
                try
                {
                    while (reader.Read())
                    {
                        if (reader.HasRows)
                        {
                            data.Add(reader.GetString(1), reader.GetFloat(2));

                        }
                    }
                    return data;
                }
                catch (Exception e)
                {
                    throw e;
                }
                finally
                {
                    reader.Close();  //关闭
                }
            }
        }

        /// <summary>
        /// 查询InfoTable
        ///返回的数据是已分组数据
        /// </summary>
        public InfoEntityData SelectInfoData(string EntityName)
        {
            //判断是否创建该查询的表（一定要先打开数据库）
            if (!isExist("InfoTable", "table"))
            {
                MessageBox.Show("本地数据库不存在，建议更新本地数据库！");
                return null;
            }
            Dictionary<string, string> Data = new Dictionary<string, string>();
            Dictionary<string, List<string>> GroupMsg = new Dictionary<string, List<string>>();
            if (conn.State != ConnectionState.Open)
            {
                conn.Open();
            }
            string getDataSql = String.Format("select KeyName, Value from KeyTable kt, EntityTable et,InfoTable it where it.Key_ID = kt.ID and it.Entity_ID = et.ID and et.EntityName = '{0}'", EntityName);
            string getGroupMsgSql = String.Format("select GroupName, KeyName from GroupTable gt, KeyTable kt, EntityTable et where kt.Group_ID = gt.ID and gt.ExcelSignal = et.ExcelSignal and et.EntityName = '{0}' order by GroupName ", EntityName);
            using (SQLiteCommand command = new SQLiteCommand(getDataSql, conn))
            {
                SQLiteDataReader reader = command.ExecuteReader();
                try
                {
                    while (reader.Read())
                    {
                        Data.Add(reader.GetString(0), reader.GetString(1));
                    }
                }
                catch (Exception e)
                {
                    throw e;
                }
                finally
                {
                    reader.Close();
                }
            }
            //查询备注
            string sql_remark = string.Format("Select Remark From EntityTable Where EntityName='{0}'", EntityName);
            string remark = null;
            using (SQLiteCommand command = new SQLiteCommand(sql_remark, conn))
            {
                SQLiteDataReader reader = command.ExecuteReader();
                try
                {
                    while (reader.Read())
                    {
                        if (!reader.IsDBNull(0))
                        {
                            remark = reader.GetString(0);
                        }
                    }
                }
                catch (Exception e)
                {
                    throw e;
                }
                finally
                {
                    reader.Close();
                }
            }
            using (SQLiteCommand command = new SQLiteCommand(getGroupMsgSql, conn))
            {
                SQLiteDataReader reader = command.ExecuteReader();
                try
                {
                    string groupName = "";
                    while (reader.Read())
                    {
                        if (!groupName.Equals(reader.GetString(0)))
                        {
                            groupName = reader.GetString(0);
                            GroupMsg.Add(groupName, new List<string>());
                            GroupMsg[groupName].Add(reader.GetString(1));
                        }
                        else
                        {
                            GroupMsg[groupName].Add(reader.GetString(1));
                        }
                    }
                }
                catch (Exception e)
                {
                    throw e;
                }
                finally
                {
                    reader.Close();
                }
            }
            InfoEntityData infoData = new InfoEntityData();
            infoData.EntityName = EntityName;
            infoData.Data = Data;
            infoData.GroupMsg = GroupMsg;
            infoData.Remark = remark;
            return infoData;
        }

        /// <summary>
        /// 查询DrawDataTable
        ///查询Entity时间序列数据
        ///根据传入的起始时间查询
        /// </summary>
        public DrawEntityData SelectDrawEntityData(string EntityName, DateTime? StartDate, DateTime? EndDate)
        {
            //判断是否创建该查询的表（一定要先打开数据库）
            if (!isExist("DrawTable", "table"))
            {
                MessageBox.Show("本地数据库不存在，建议更新本地数据库！");
                return null;
            }
            DrawEntityData drawEntityData = new DrawEntityData();
            drawEntityData.EntityName = EntityName;
            drawEntityData.Data = new List<DrawData>();
            string sql = null;
            string start = null;
            string end = null;
            if (StartDate == null && EndDate == null)
            {
                sql = string.Format("select Date,EntityMaxValue,EntityMidValue,EntityMinValue,Detail from DrawTable dt, EntityTable et where dt.Entity_ID = et.ID and et.EntityName = '{0}' order by Date", EntityName);
            }
            else if (StartDate == null)
            {
                end = ((DateTime)EndDate).ToString("yyyy-MM-dd HH:mm:ss");
                sql = String.Format("select Date,EntityMaxValue,EntityMidValue,EntityMinValue,Detail from DrawTable dt, EntityTable et where dt.Entity_ID = et.ID and et.EntityName = '{0}' and Date <= '{1}' order by Date", EntityName, end);
            }
            else if (EndDate == null)
            {
                start = ((DateTime)StartDate).ToString("yyyy-MM-dd HH:mm:ss");
                sql = String.Format("select Date,EntityMaxValue,EntityMidValue,EntityMinValue,Detail from DrawTable dt, EntityTable et where dt.Entity_ID = et.ID and et.EntityName = '{0}' and Date >= '{1}' order by Date", EntityName, start);
            }
            else
            {
                end = ((DateTime)EndDate).ToString("yyyy-MM-dd HH:mm:ss");
                start = ((DateTime)StartDate).ToString("yyyy-MM-dd HH:mm:ss");
                sql = String.Format("select Date,EntityMaxValue,EntityMidValue,EntityMinValue,Detail from DrawTable dt, EntityTable et where dt.Entity_ID = et.ID and et.EntityName = '{0}' and Date between '{1}' and '{2}' order by Date", EntityName, start, end);
            }
            if (conn.State != ConnectionState.Open)
            {
                conn.Open();
            }
            using (SQLiteCommand command = new SQLiteCommand(sql, conn))  //建立执行命令语句对象
            {
                SQLiteDataReader reader = command.ExecuteReader();
                try
                {
                    while (reader.Read())
                    {
                        DrawData drawData = new DrawData();

                        drawData.Date = Convert.ToDateTime(reader.GetString(0));
                        drawData.MaxValue = reader.GetFloat(1);
                        drawData.MidValue = reader.GetFloat(2);
                        drawData.MinValue = reader.GetFloat(3);
                        drawData.Detail = reader.GetString(4);
                        drawEntityData.Data.Add(drawData);
                    }
                    reader.Close();
                    command.CommandText = string.Format("Select Total_hold,Diff_hold From ExcelTable,EntityTable Where EntityTable.EntityName='{0}' and EntityTable.ExcelSignal=ExcelTable.ExcelSignal", EntityName);
                    reader = command.ExecuteReader();
                    if (reader.Read())
                    {
                        drawEntityData.Total_hold = reader.GetFloat(0);
                        drawEntityData.Diff_hold = reader.GetFloat(1);
                    }
                }
                catch (Exception e)
                {
                    throw e;
                }
                finally
                {
                    if (!reader.IsClosed)
                    {
                        reader.Close();
                    }
                    conn.Close();
                }
            }
            return drawEntityData;
        }


        /// <summary>
        /// 查询DrawDataTable
        ///查询Entity某日期的数据
        /// </summary>
        public DrawData SelectDrawData(string EntityName, DateTime Date)
        {
            //判断是否创建该查询的表（一定要先打开数据库）
            if (!isExist("DrawDataTable", "table"))
            {
                MessageBox.Show("本地数据库不存在，建议更新本地数据库！");
                return null;
            }
            DrawData drawData = new DrawData();
            drawData.Date = Date;
            string datestr = Date.ToString("yyyy-MM-dd");
            string sql = String.Format("select EntityMaxValue,EntityMidValue,EntityMinValue,Detail from DrawDataTable dt, EntityTable et where dt.Entity_ID = et.ID and et.EntityName = '{0}' and dt.date = '{1}'", EntityName, datestr);

            if (conn.State != ConnectionState.Open)
            {
                conn.Open();
            }
            using (SQLiteCommand command = new SQLiteCommand(sql, conn))  //建立执行命令语句对象
            {
                SQLiteDataReader reader = command.ExecuteReader();
                try
                {
                    while (reader.Read())
                    {

                        drawData.MaxValue = reader.GetFloat(0);
                        drawData.MidValue = reader.GetFloat(1);
                        drawData.MinValue = reader.GetFloat(2);
                        drawData.Detail = reader.GetString(3);

                    }

                    return drawData;
                }
                catch (Exception e)
                {
                    throw e;
                }
                finally
                {
                    reader.Close();
                }
            }
        }
        /// <summary>
        ///通过传入的Signal，查询与之对应的所有的测点
        ///传入的Signal应该是测量数据的signal
        ///ErrMsg:Total,TotalDiff,No,NoDiff
        /// </summary>
        public List<CEntityName> SelectAllEntitiesAndErr(string ExcelSignal)
        {
            //判断是否创建该查询的表（一定要先打开数据库）
            if (!isExist("EntityTable", "table"))
            {
                MessageBox.Show("本地数据库不存在，建议更新本地数据库！");
                return null;
            }
            if (conn.State != ConnectionState.Open)
            {
                conn.Open();
            }
            string select_Threshold = string.Format("Select Total_hold,Diff_hold From ExcelTable Where ExcelSignal='{0}'", ExcelSignal);
            List<CEntityName> Entities = new List<CEntityName>();
            Dictionary<string, CEntityName> maps = new Dictionary<string, CEntityName>();

            using (SQLiteCommand command = new SQLiteCommand(select_Threshold, conn))  //建立执行命令语句对象
            {
                SQLiteDataReader reader = null;
                try
                {
                    reader = command.ExecuteReader();
                    float? Total_hold = null;
                    float? Diff_hold = null;
                    if (reader.Read())
                    {
                        Total_hold = reader.GetFloat(0);
                        Diff_hold = reader.GetFloat(1);
                    }
                    reader.Close();
                    if (Total_hold == null || Diff_hold == null)
                    {
                        throw new Exception("无效的阈值");
                    }
                    string sql_Total = null;
                    if (Total_hold >= 0)
                    {
                        sql_Total = String.Format("select EntityTable.ID,EntityTable.EntityName,Max(DrawTable.EntityMaxValue)>={0} From  EntityTable,DrawTable where DrawTable.Entity_ID=EntityTable.ID and EntityTable.ExcelSignal = '{1}' GROUP BY EntityTable.EntityName ORDER BY EntityTable.ID", Total_hold, ExcelSignal);
                    }
                    else
                    {
                        sql_Total = String.Format("select EntityTable.ID,EntityTable.EntityName,Min(DrawTable.EntityMaxValue)<={0} From  EntityTable,DrawTable where DrawTable.Entity_ID=EntityTable.ID and EntityTable.ExcelSignal = '{1}' GROUP BY EntityTable.EntityName ORDER BY EntityTable.ID", Total_hold, ExcelSignal);
                    }
                    command.CommandText = sql_Total;
                    reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        CEntityName one = new CEntityName();
                        one.Id = reader.GetInt32(0);
                        one.EntityName = reader.GetString(1);
                        one.ErrMsg = reader.GetBoolean(2) ? "Total" : "No";
                        Entities.Add(one);
                        maps.Add(one.EntityName, one);
                    }
                    reader.Close();
                    string sql_Diff = string.Format("SELECT DrawTable.EntityMaxValue,EntityTable.EntityName From DrawTable ,EntityTable WHERE DrawTable.Entity_ID = EntityTable.ID and EntityTable.ExcelSignal='{0}' Order BY EntityTable.ID,DrawTable.Date", ExcelSignal);
                    command.CommandText = sql_Diff;
                    reader = command.ExecuteReader();
                    Diff_hold = Math.Abs((float)Diff_hold);
                    float first = 0;
                    float next = 0;
                    float diff = 0;
                    bool isErr = false;
                    string entityName = null;
                    if (reader.Read())
                    {
                        first = reader.GetFloat(0);
                        entityName = reader.GetString(1);
                    }
                    while (reader.Read())
                    {
                        next = reader.GetFloat(0);
                        diff = Math.Abs((float)(next - first));
                        if (entityName.Equals(reader.GetString(1)))
                        {
                            if (diff >= Diff_hold)
                            {
                                isErr = true;
                            }
                        }
                        else if (isErr)
                        {
                            maps[entityName].ErrMsg += "Diff";
                            isErr = false;
                            entityName = reader.GetString(1);
                        }
                        else
                        {
                            entityName = reader.GetString(1);
                        }
                        first = next;
                    }
                    //最后一个
                    if (isErr) {
                        maps[entityName].ErrMsg += "Diff";
                    }
                }
                catch (Exception e)
                {
                    throw e;
                }
                finally
                {
                    if (!reader.IsClosed)
                    {
                        reader.Close();
                    }
                    conn.Close();
                }
            }
            return Entities;
        }
        public List<ExcelTable> SelectDrawTypes()
        {
            List<ExcelTable> result = null;
            if (conn.State != ConnectionState.Open)
            {
                conn.Open();
            }
            string sql = "Select ID,CurrentFile,ExcelSignal,Total_hold,Diff_hold,History From ExcelTable Where IsInfo=0";
            using (SQLiteCommand command = new SQLiteCommand(sql, conn))
            {
                SQLiteDataReader reader = null;
                try
                {
                    reader = command.ExecuteReader();
                    if (reader.HasRows)
                    {
                        result = new List<ExcelTable>();
                    }
                    while (reader.Read())
                    {
                        ExcelTable one = new ExcelTable();
                        one.Id = reader.GetInt32(0);
                        one.CurrentFile = reader.GetString(1);
                        one.Signal = reader.GetString(2);
                        one.Total_hold = reader.GetFloat(3);
                        one.Diff_hold = reader.GetFloat(4);
                        one.History = reader.GetString(5);
                        result.Add(one);
                    }
                }
                catch (Exception e)
                {
                    throw e;
                }
                finally
                {
                    reader.Close();
                    conn.Close();
                }
            }
            return result;
        }

        public bool ModifyEntityRemark(string EntityName, string Remark)
        {
            bool result = false;
            if (string.IsNullOrWhiteSpace(EntityName) || string.IsNullOrWhiteSpace(Remark))
            {
                return false;
            }
            string sql = string.Format("Update EntityTable Set Remark='{0}' Where EntityName='{1}'", Remark, EntityName);
            try
            {
                if (conn.State != ConnectionState.Open)
                {
                    conn.Open();
                }
                SQLiteCommand command = new SQLiteCommand(sql, conn);
                if (command.ExecuteNonQuery() > 0)
                {
                    result = true;
                }
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                conn.Close();
            }
            return result;
        }


    }
}
