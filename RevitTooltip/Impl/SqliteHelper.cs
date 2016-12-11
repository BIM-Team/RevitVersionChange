using System;
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
                    command.CommandText = "CREATE TABLE ExcelTable(ID integer NOT NULL PRIMARY KEY AUTOINCREMENT,CurrentFile VARCHAR(30) NOT NULL,ExcelSignal VARCHAR(20) UNIQUE, IsInfo BOOLEAN, Total_hold float, Diff_hold float, History VARCHAR(100) )";
                    command.ExecuteNonQuery();
                }

                if (!isExist("KeyTable", "table"))
                {
                    command.CommandText = "CREATE TABLE KeyTable(ID integer PRIMARY KEY AUTOINCREMENT, ExcelSignal VARCHAR(20) NOT NULL, Group_ID integer, KeyName VARCHAR(20) NOT NULL,Odr integer )";
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
                    command.CommandText = "CREATE TABLE InfoTable(ID integer PRIMARY KEY AUTOINCREMENT, Key_ID integer NOT NULL REFERENCES KeyTable(ID) ON DELETE CASCADE ON UPDATE NO ACTION, Entity_ID integer NOT NULL REFERENCES EntityTable(ID) ON DELETE CASCADE ON UPDATE NO ACTION, Value VARCHAR(20))";
                    command.ExecuteNonQuery();
                }

                if (!isExist("DrawTable", "table"))
                {
                    command.CommandText = "CREATE TABLE DrawTable(ID integer PRIMARY KEY AUTOINCREMENT,  Entity_ID integer NOT NULL REFERENCES EntityTable(ID) ON DELETE CASCADE ON UPDATE NO ACTION, Date VARCHAR(20), EntityMaxValue float, EntityMidValue float, EntityMinValue float, Detail TEXT )";
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
                            mysql_reader.GetInt32(0), mysql_reader.GetInt32(1), mysql_reader.GetDateTime(2), mysql_reader.GetFloat(3), mysql_reader.GetFloat(4), mysql_reader.GetFloat(5), mysql_reader.GetString(6));
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
            if (conn.State != ConnectionState.Open)
            {
                conn.Open();
            }
            using (SQLiteTransaction tran = conn.BeginTransaction())
            {
                try
                {
                    if (!File.Exists(Path.Combine(dbName, dbPath)))
                    {
                        CreateDB();
                    }
                    tran.Commit();
                }
                catch (Exception)
                {
                    tran.Rollback();
                    throw;
                }
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
        ///// <summary>
        ///// 当插入一个sheet数据回滚时，删除当前sheet表的entity
        ///// </summary>
        ///// <param name="sheetInfo"></param>
        ///// <param name="IdExcel"></param>
        //private void DeleteCurrentEntity(SheetInfo sheetInfo, int IdExcel)
        //{
        //    Dictionary<string, int> PeriousEntities = SelectEntities(IdExcel);
        //    string sql = "delete from entitytable where EntityName in (";
        //    int n = 0;
        //    foreach (string entityName in sheetInfo.EntityNames)
        //    {
        //        if (PeriousEntities.Keys.Contains(entityName))
        //        {
        //            if (n != 0)
        //                sql += ",";

        //            sql += "'" + entityName + "'";
        //            n++;
        //        }
        //    }

        //    if (n == 0)
        //        sql += "'')";
        //    else
        //        sql += ")";

        //    ExecuteOneSql(sql);
        //}

        /// <summary>
        /// 插入基础数据表
        /// </summary>
        /// <param name="sheetInfo"></param>
        private void InsertInfoData(SheetInfo sheetInfo)
        {
            //是否已经处理过
            bool hasDone = false;
            string signal = sheetInfo.ExcelTableData.Signal;
            string reset_auto_increment = "alter table ExcelTable auto_increment =1;"
                         + "alter table KeyTable auto_increment =1;"
                         + "alter table EntityTable auto_increment =1;"
                         + "alter table GroupTable auto_increment =1;"
                         + "alter table InfoTable auto_increment =1;"
                         + "alter table DrawTable auto_increment =1";
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
                ExcelTable exist = SelectExcelTable(signal, conn);
                if (exist == null)
                {
                    //插入ExcelTable
                    new SQLiteCommand(string.Format("insert into ExcelTable (CurrentFile,ExcelSignal,IsInfo,History) values ('{0}', '{1}', {2},'{3}')",
                         sheetInfo.ExcelTableData.CurrentFile, signal, sheetInfo.Tag, sheetInfo.ExcelTableData.CurrentFile), tran.Connection, tran).ExecuteNonQuery();
                    exist = SelectExcelTable(signal, conn);
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
                CEntityName entity = selectEntity(one.EntityName, tran.Connection);
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
        private CEntityName selectEntity(string entityName, SQLiteConnection OpenedConn, bool err = false)
        {
            CEntityName result = null;
            //不需要Err信息
            if (!err)
            {
                string sql = string.Format("select ID,EntityName from EntityTable where EntityName='{0}'", entityName);
                if (OpenedConn.State != ConnectionState.Open)
                {
                    OpenedConn.Open();
                }
                SQLiteCommand command = new SQLiteCommand(sql, OpenedConn);
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
        private ExcelTable SelectExcelTable(String signal, SQLiteConnection OpenedConn)
        {
            string sql = String.Format("select ID,CurrentFile,ExcelSignal,Total_hold,Diff_hold,History from ExcelTable where ExcelSignal = '{0}'", signal);
            SQLiteDataReader reader = new SQLiteCommand(sql, OpenedConn).ExecuteReader();
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
                string reset_auto_increment = "alter table ExcelTable auto_increment =1;"
                             + "alter table KeyTable auto_increment =1;"
                             + "alter table EntityTable auto_increment =1;"
                             + "alter table GroupTable auto_increment =1;"
                             + "alter table InfoTable auto_increment =1;"
                             + "alter table DrawTable auto_increment =1";
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
                ExcelTable exist = SelectExcelTable(signal, tran.Connection);
                if (exist == null)
                {
                    command.CommandText = string.Format("insert into ExcelTable (CurrentFile,ExcelSignal,IsInfo,History) values ('{0}', '{1}', {2},'{3}')",
                          sheetInfo.ExcelTableData.CurrentFile, signal, sheetInfo.Tag, sheetInfo.ExcelTableData.CurrentFile);
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
                CEntityName entity = selectEntity(one.EntityName, tran.Connection);
                DateTime? maxDate = null;
                if (entity == null)
                {
                    string sql = string.Format("insert into EntityTable(ExcelSignal,EntityName) values ('{0}','{1}')", signal, one.EntityName);
                    //插入Entity
                    new SQLiteCommand(sql, tran.Connection, tran).ExecuteNonQuery();
                    entity = selectEntity(one.EntityName, tran.Connection);
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
                foreach (DrawData p in data)
                {
                    if (maxDate == null || p.Date > maxDate)
                    {
                        buider.AppendFormat(" ({0},'{1}','{2}','{3}','{4}','{5}'),", entity.Id, p.Date, p.MaxValue, p.MidValue, p.MinValue, p.Detail);
                    }
                }
                buider.Remove(buider.Length - 1, 1);
                new SQLiteCommand(buider.ToString(), tran.Connection, tran).ExecuteNonQuery();
            }

        }
        ///// <summary>
        ///// 插入到ExcelTable表
        ///// 并返回IdExcel和IdGroup
        ///// </summary>
        ///// <param name="sheetInfo"></param>
        ///// <returns></returns>
        //private Dictionary<string, int> InsertIntoExcelTable(SheetInfo sheetInfo)
        //{
        //    Dictionary<string, int> id = new Dictionary<string, int>();
        //    bool isInfo = sheetInfo.Tag;
        //    string signal = sheetInfo.ExcelTableData.Signal;
        //    string tableDesc = sheetInfo.ExcelTableData.TableDesc;

        //    int IdExcel = getIdExcel(signal);
        //    int IdGroup = getIdGroup(IdExcel);
        //    string currentTableDesc = GetTableDesc(signal);
        //    bool hasTableDesc = false;

        //    if (IdExcel == 0)
        //    {
        //        string sql = GetInsertIntoExcelTableSql(isInfo, signal, tableDesc);
        //        ExecuteOneSql(sql);
        //        IdExcel = getIdExcel(signal);

        //        //插入到GroupTable
        //        IdGroup = InsertIntoGroupTable(IdExcel);
        //    }
        //    else
        //    {
        //        string[] tableDesclist = currentTableDesc.Split(';');
        //        foreach (string t in tableDesclist)
        //        {
        //            if (t.Equals(tableDesc))
        //            {
        //                hasTableDesc = true;
        //            }
        //        }
        //        if (!hasTableDesc)
        //        {
        //            string updateTableDesc = currentTableDesc + tableDesc;
        //            UpdateTableDesc(signal, updateTableDesc);
        //        }
        //    }
        //    id.Add("IdExcel", IdExcel);
        //    id.Add("IdGroup", IdGroup);

        //    return id;
        //}
        ///// <summary>
        ///// 插入“未分组”Group
        ///// </summary>
        ///// <param name="idExcel"></param>
        ///// <returns></returns>
        //private int InsertIntoGroupTable(int idExcel)
        //{
        //    int idGroup = 0;
        //    string sql = GetInsertIntoGroupTableSql(idExcel) + ";select @@IDENTITY ";

        //    SQLiteCommand mycom = new SQLiteCommand();
        //    mycom.Connection = this.conn;
        //    mycom.CommandText = sql;
        //    if (isBegin)
        //        mycom.Transaction = myTran;

        //    using (mycom)
        //    {
        //        try
        //        {
        //            idGroup = Convert.ToInt32(mycom.ExecuteScalar());
        //            return idGroup;
        //        }
        //        catch { throw; }
        //    }

        //}
        ///// <summary>
        ///// 插入属性到KeyTable表
        ///// 返回当前表的属性及对应ID，<KeyName,ID>
        ///// </summary>
        ///// <param name="KeyNames"></param>
        ///// <param name="IdExcel"></param>
        ///// <returns></returns>
        //private Dictionary<string, int> InsertIntoKeyTable(List<string> KeyNames, Dictionary<string, int> id)
        //{
        //    int IdExcel = id["IdExcel"];
        //    int IdGroup = id["IdGroup"];
        //    Dictionary<string, int> PeriousKeys = SelectKeyNames(IdExcel);
        //    String sql = "";
        //    int n = 0;
        //    foreach (string name in KeyNames)
        //    {
        //        if (name.Equals("备注") || PeriousKeys.Keys.Contains("备注"))
        //        {
        //            hasEntityRemark = true;
        //            continue;
        //        }

        //        //如果没有在原表中匹配到，则添加到插入语句当中
        //        if (!PeriousKeys.Keys.Contains(name))
        //        {
        //            if (n == 0)
        //                sql = GetInsertIntoKeyTableSql(IdExcel, IdGroup, name);
        //            else
        //                sql = sql + ",('" + IdExcel + "','" + IdGroup + "','" + name + "')";
        //            n++;
        //        }
        //    }
        //    if (!hasEntityRemark)
        //    {
        //        sql = sql + ",('" + IdExcel + "','" + IdGroup + "','备注')";
        //    }
        //    ExecuteOneSql(sql);
        //    //获取更新后的<属性，ID>
        //    Dictionary<string, int> CurrentKeys = SelectKeyNames(IdExcel);
        //    return CurrentKeys;
        //}

        ///// <summary>
        ///// 插入实体到EntityTable表
        ///// 返回当前插入的EntityNames,<EntityName,ID> 
        ///// </summary>
        ///// <param name="EntityNames"></param>
        ///// <param name="IdExcel"></param>
        ///// <returns></returns>
        //private Dictionary<string, int> InsertIntoEntityTable(List<string> EntityNames, int IdExcel)
        //{
        //    Dictionary<string, int> UpdateEntities = new Dictionary<string, int>();
        //    List<string> UpdateEntityNames = new List<string>();
        //    Dictionary<string, int> PeriousEntities = SelectEntities(IdExcel);
        //    String sql = "";
        //    int n = 0;
        //    foreach (string entityName in EntityNames)
        //    {
        //        //如果没有在原表中匹配到，则添加到插入语句当中
        //        if (!PeriousEntities.Keys.Contains(entityName))
        //        {
        //            if (n == 0)
        //                sql = GetInsertIntoEntityTableSql(IdExcel, entityName);
        //            else
        //                sql = sql + ",('" + IdExcel + "','" + entityName + "','" + "')";

        //            UpdateEntityNames.Add(entityName);
        //            n++;
        //        }
        //    }
        //    ExecuteOneSql(sql);

        //    //获取当前表所有的Entities
        //    foreach (string s in UpdateEntityNames)
        //    {
        //        int id = getIdEntity(IdExcel, s);
        //        UpdateEntities.Add(s, id);
        //    }

        //    return UpdateEntities;
        //}
        ///// <summary>
        ///// 插入到InfoTable
        ///// </summary>
        ///// <param name="sheetInfo"></param>
        ///// <param name="CurrentKeys"></param>
        ///// <param name="UpdateEntities"></param>
        //private void InsertIntoInfoTable(SheetInfo sheetInfo, Dictionary<string, int> CurrentKeys, Dictionary<string, int> UpdateEntities)
        //{
        //    int IdKey;
        //    int IdEntity;
        //    List<string> SqlStringList = new List<string>();  // 用来存放多条SQL语句
        //    string sql = "";
        //    int n = 0;
        //    foreach (InfoEntityData infoRow in sheetInfo.InfoRows)
        //    {
        //        if (UpdateEntities.Keys.Contains(infoRow.EntityName))
        //        {
        //            IdEntity = UpdateEntities[infoRow.EntityName];
        //            foreach (string key in infoRow.Data.Keys)
        //            {
        //                IdKey = CurrentKeys[key];
        //                if (n % 500 == 0)
        //                {
        //                    if (n > 0)
        //                        SqlStringList.Add(sql);
        //                    sql = GetInsertIntoInfoTableSql(IdKey, IdEntity, infoRow.Data[key]);
        //                }
        //                else
        //                    sql = sql + ",('" + IdKey + "','" + IdEntity + "','" + infoRow.Data[key] + "')";

        //                n++;
        //            }
        //        }
        //    }
        //    if (n % 500 != 0)
        //        SqlStringList.Add(sql);  //不够500values的SQL语言也要添加进去
        //    InsertSqlStringList(SqlStringList);

        //}
        ///// <summary>
        ///// 插入到DrawDataTabel
        ///// </summary>
        ///// <param name="sheetInfo"></param>
        ///// <param name="UpdateEntities"></param>
        ///// <param name="IdExcel"></param>
        //private void InsertIntoDrawDataTable(SheetInfo sheetInfo, Dictionary<string, int> UpdateEntities, int IdExcel)
        //{
        //    int IdEntity;
        //    List<string> SqlStringList = new List<string>();  // 用来存放多条SQL语句
        //    String sql = "";
        //    int n = 0;

        //    foreach (DrawEntityData drawRow in sheetInfo.DrawRows)
        //    {
        //        if (UpdateEntities.Keys.Contains(drawRow.EntityName))
        //        {
        //            IdEntity = UpdateEntities[drawRow.EntityName];
        //            foreach (DrawData data in drawRow.Data)
        //            {
        //                string datestr = data.Date.ToString("yyyy-MM-dd HH:mm:ss");
        //                if (n % 500 == 0)
        //                {
        //                    if (n > 0)
        //                        SqlStringList.Add(sql);
        //                    sql = GetInsertIntoDrawDataTableSql(IdExcel, IdEntity, datestr, data.MaxValue, data.MidValue, data.MinValue, data.Detail);
        //                }
        //                else
        //                    sql = sql + ",('" + IdExcel + "','" + IdEntity + "','" + datestr + "','" + data.MaxValue + "','" + data.MidValue + "','" + data.MinValue + "','" + data.Detail + "')";

        //                n++;
        //            }
        //        }
        //        else  //实体表如果没有更新，要检查每个实体表中，时间是否增加
        //        {
        //            IdEntity = getIdEntity(IdExcel, drawRow.EntityName);
        //            List<string> PeriousDateTimes = GetDrawDataTableDateTimes(IdEntity);
        //            foreach (DrawData data in drawRow.Data)
        //            {
        //                string datestr = data.Date.ToString("yyyy-MM-dd HH:mm:ss");
        //                if (!PeriousDateTimes.Contains(datestr))
        //                {
        //                    if (n % 500 == 0)
        //                    {
        //                        if (n > 0)
        //                            SqlStringList.Add(sql);
        //                        sql = GetInsertIntoDrawDataTableSql(IdExcel, IdEntity, datestr, data.MaxValue, data.MidValue, data.MinValue, data.Detail);
        //                    }
        //                    else
        //                        sql = sql + ",('" + IdExcel + "','" + IdEntity + "','" + datestr + "','" + data.MaxValue + "','" + data.MidValue + "','" + data.MinValue + "','" + data.Detail + "')";

        //                    n++;
        //                }
        //            }
        //        }
        //    }
        //    if (n % 500 != 0)
        //        SqlStringList.Add(sql);  //不够500values的SQL语言也要添加进去            
        //    InsertSqlStringList(SqlStringList);
        //}


        ///// <summary>
        ///// 获取KeyTable表的Key和ID
        ///// </summary>
        ///// <param name="IdExcel"></param>
        ///// <returns></returns>
        //private Dictionary<string, int> SelectKeyNames(int IdExcel)
        //{
        //    Dictionary<string, int> KeyNames = new Dictionary<string, int>();
        //    String sql = "select KeyName, ID from KeyTable where Excel_ID = " + IdExcel;

        //    SQLiteCommand mycom = new SQLiteCommand(sql, this.conn, myTran);  //建立执行命令语句对象
        //    SQLiteDataReader reader = mycom.ExecuteReader();    //需要关闭
        //    try
        //    {
        //        while (reader.Read())
        //        {
        //            if (reader.HasRows)
        //            {
        //                KeyNames.Add(reader.GetString(0), reader.GetInt32(1));
        //            }
        //        }
        //        return KeyNames;
        //    }
        //    catch (Exception e)
        //    {
        //        throw e;
        //    }
        //    finally
        //    {
        //        reader.Close();
        //    }
        //}
        ///// <summary>
        ///// 获取EntityTable表的EntityName和ID
        ///// </summary>
        ///// <param name="IdExcel"></param>
        ///// <returns></returns>
        //private Dictionary<string, int> SelectEntities(int IdExcel)
        //{
        //    Dictionary<string, int> Entities = new Dictionary<string, int>();
        //    string sql = "select distinct EntityName, ID from EntityTable where Excel_ID = " + IdExcel;
        //    SQLiteCommand mycom = new SQLiteCommand();
        //    mycom.Connection = this.conn;
        //    mycom.CommandText = sql;
        //    if (isBegin)
        //        mycom.Transaction = myTran;
        //    SQLiteDataReader reader = mycom.ExecuteReader();    //需要关闭
        //    try
        //    {
        //        while (reader.Read())
        //        {
        //            if (reader.HasRows)
        //            {
        //                //这里如果添加了相同的项会出错，所以entity一定要唯一
        //                Entities.Add(reader.GetString(0), reader.GetInt32(1));
        //            }
        //        }
        //        return Entities;
        //    }
        //    catch (Exception e)
        //    {
        //        throw e;
        //    }
        //    finally
        //    {
        //        reader.Close();
        //    }
        //}
        ///// <summary>
        ///// 从DrawDataTable表中，获取该实体数据中，当前存储的DATETIME项
        ///// </summary>
        ///// <param name="IdEntity"></param>
        ///// <returns></returns>
        //private List<string> GetDrawDataTableDateTimes(int IdEntity)
        //{
        //    List<string> PeriousDateTimes = new List<string>();
        //    String sql = "select Date from DrawDataTable where Entity_ID = " + IdEntity;

        //    SQLiteCommand mycom = new SQLiteCommand(sql, this.conn, myTran);  //建立执行命令语句对象
        //    SQLiteDataReader reader = mycom.ExecuteReader();    //需要关闭
        //    try
        //    {
        //        while (reader.Read())
        //        {
        //            if (reader.HasRows)
        //            {
        //                PeriousDateTimes.Add(reader.GetString(0));
        //            }
        //        }
        //        return PeriousDateTimes;
        //    }
        //    catch (Exception e)
        //    {
        //        throw e;
        //    }
        //    finally
        //    {
        //        reader.Close();
        //    }
        //}

        ///// <summary>
        ///// 获取当前Excel的TableDesc
        ///// </summary>
        ///// <param name="signal"></param>
        ///// <returns></returns>
        //private string GetTableDesc(string signal)
        //{
        //    string currentTableDesc = "";
        //    string sql = String.Format("select tableDesc from ExcelTable where ExcelSignal = '{0}' ", signal);

        //    SQLiteCommand mycom = new SQLiteCommand();
        //    mycom.Connection = this.conn;
        //    mycom.CommandText = sql;
        //    if (isBegin)
        //        mycom.Transaction = myTran;

        //    SQLiteDataReader reader = mycom.ExecuteReader();    //需要关闭
        //    try
        //    {
        //        while (reader.Read())
        //        {
        //            if (reader.HasRows)
        //            {
        //                currentTableDesc = reader.GetString(0);
        //            }
        //        }
        //        return currentTableDesc;
        //    }
        //    catch (Exception e)
        //    {
        //        throw e;
        //    }
        //    finally
        //    {
        //        reader.Close();
        //    }
        //}
        ///// <summary>
        ///// 更新TableDesc
        ///// </summary>
        ///// <param name="signal"></param>
        ///// <param name="updateTableDesc"></param>
        //private void UpdateTableDesc(string signal, string updateTableDesc)
        //{
        //    string sql = String.Format("update ExcelTable set tableDesc = '{0}' where ExcelSignal = '{1}' ", updateTableDesc, signal);
        //    ExecuteOneSql(sql);
        //}
        //private Boolean HasCX(string CX)
        //{
        //    string sql = String.Format("select * from CXTable where ENTITY = '{0}'", CX);
        //    SQLiteCommand mycom = new SQLiteCommand(sql, this.conn, myTran);  //建立执行命令语句对象
        //    SQLiteDataReader reader = mycom.ExecuteReader();
        //    try
        //    {
        //        reader.Read();
        //        if (reader.HasRows)
        //            return true;
        //        else
        //            return false;
        //    }
        //    finally
        //    {
        //        reader.Close();
        //    }

        //}
        ////Test############################################################################
        //public List<string> SelectCXTable()
        //{
        //    Dictionary<string, string> cxTable = new Dictionary<string, string>();
        //    List<string> list = new List<string>();
        //    //String sql = "Select e.Entity,c.Date,c.Value from CXView c, EntityTable e where c.ID_Entity = e.ID "
        //    //           + " group by c.ID_Entity,c.Date order by c.ID_Entity,c.Date ";

        //    // String sql = " Select ENTITY,DATE,VALUE from CXTable group by ENTITY,DATE order by ENTITY,DATE";
        //    String sql = " Select ENTITY,DATE,VALUE from CXTable order by ENTITY,DATE";

        //    //判断数据库是否打开
        //    if (conn.State != ConnectionState.Open)
        //    {
        //        OpenConnect();
        //    }
        //    SQLiteCommand mycom = new SQLiteCommand(sql, this.conn);  //建立执行命令语句对象
        //    SQLiteDataReader reader = mycom.ExecuteReader();
        //    try
        //    {
        //        while (reader.Read())
        //        {
        //            if (reader.HasRows)
        //            {

        //                //cxTable.Add(reader.GetString(1), reader.GetFloat(2));
        //                //DateTime dateTime = Convert.ToDateTime(reader.GetString(1));
        //                //string datestr = dateTime.ToString("yyyy-MM-dd");
        //                //if (dateTime.Hour >= 12)
        //                //{
        //                //    datestr += "pm";
        //                //}
        //                list.Add(reader.GetString(1));
        //            }
        //        }
        //        return list;
        //    }
        //    catch (Exception e)
        //    {
        //        throw e;
        //    }
        //    finally
        //    {
        //        reader.Close();
        //    }
        //}



        ///// <summary>
        ///// 执行批量插入
        ///// </summary>
        //private void InsertSqlStringList(List<string> SqlStringList)
        //{

        //    SQLiteCommand command = new SQLiteCommand();
        //    command.Connection = this.conn;
        //    command.Transaction = myTran;
        //    try
        //    {
        //        for (int i = 0; i < SqlStringList.Count; i++)
        //        {
        //            string sql = SqlStringList[i].ToString();
        //            if (sql.Equals(""))
        //                return;
        //            command.CommandText = sql;
        //            command.ExecuteNonQuery();
        //        }
        //    }
        //    catch (Exception e)
        //    {
        //        throw e;
        //    }
        //}
        ///// <summary>
        ///// 执行单个sql语句，返回插入结果
        ///// </summary>
        ///// <param name="sql"></param>
        //private void ExecuteOneSql(string sql)
        //{
        //    if (sql.Equals(""))
        //        return;
        //    SQLiteCommand mycom = new SQLiteCommand();
        //    mycom.Connection = this.conn;
        //    mycom.CommandText = sql;
        //    if (isBegin)
        //        mycom.Transaction = myTran;

        //    try
        //    {
        //        mycom.ExecuteNonQuery();
        //    }
        //    catch (Exception e)
        //    {
        //        throw e;
        //    }
        //}
        ////执行单个插入语句，返回插入结果
        //private void InsertOne(String sql)
        //{
        //    if (sql.Equals(""))
        //        return;
        //    SQLiteCommand mycom = new SQLiteCommand(sql, this.conn, myTran);  //建立执行命令语句对象，其中myTran为事务

        //    try
        //    {
        //        mycom.ExecuteNonQuery();
        //    }
        //    catch (Exception e)
        //    {
        //        throw e;
        //    }
        //}
        ////根据entity,修改备注
        //public bool ModifyEntityRemark(string entity, string entityremark)
        //{
        //    bool flag = false;
        //    string sql = String.Format("Update EntityTable set ENTITYREMARK = '{0}' where ENTITY = '{1}'", entityremark, entity);
        //    //判断数据库是否打开
        //    if (conn.State != ConnectionState.Open)
        //    {
        //        OpenConnect();
        //    }
        //    SQLiteCommand mycom;
        //    if (isBegin)
        //        mycom = new SQLiteCommand(sql, this.conn, this.myTran);  //建立执行命令语句对象
        //    else
        //        mycom = new SQLiteCommand(sql, this.conn);  //建立执行命令语句对象

        //    try
        //    {
        //        if (mycom.ExecuteNonQuery() > 0)
        //            flag = true;
        //        return flag;
        //    }
        //    catch (Exception e)
        //    {
        //        throw e;
        //    }
        //}

        //表ExcelTable的LoadSql
        private String GetLoadIntoExcelTableSql(string tableDesc, string signal, bool isInfo, int id = 0, double total_hold = 0, double diff_hold = 0, string remark = null)
        {
            String sql = String.Format("INSERT OR IGNORE INTO ExcelTable(ID,TableDesc,ExcelSignal,IsInfo,Total_hold,Diff_hold,Remark) VALUES ('{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}')", id, tableDesc, signal, isInfo, total_hold, diff_hold, remark);
            return sql;
        }
        //表KeyTable的LoadSql
        private String GetLoadIntoKeyTableSql(int idExcel, string keyName, int id = 0, int idGroup = 0)
        {
            String sql = String.Format("INSERT OR IGNORE INTO KeyTable (ID,Excel_ID,Group_ID,KeyName) VALUES ({0} ,{1}, {2}, '{3}') ", id, idExcel, idGroup, keyName);
            return sql;
        }
        //表EntityTable的LoadSql
        private String GetLoadIntoEntityTableSql(int idExcel, string entityName, int id = 0, string remark = "")
        {
            String sql = String.Format("INSERT OR IGNORE INTO EntityTable(ID,Excel_ID,EntityName,Remark) VALUES ('{0}','{1}','{2}','{3}')", id, idExcel, entityName, remark);
            return sql;
        }
        //表InfoTable的LoadSql
        private String GetLoadIntoInfoTableSql(int idKey, int idEntity, string value, int id = 0)
        {
            String sql = String.Format("INSERT OR IGNORE into InfoTable ( ID, Key_ID, Entity_ID, Value ) VALUES ('{0}','{1}','{2}','{3}')", id, idKey, idEntity, value);
            return sql;
        }

        //表DrawDataTable的LoadSql
        private String GetLoadIntoDrawDataTableSql(int idExcel, int idEntity, string date, double maxValue, double midValue, double minValue, string detail, int id = 0)
        {
            String sql = String.Format("INSERT OR IGNORE into DrawDataTable (ID, Excel_ID, Entity_ID, Date, EntityMaxValue, EntityMidValue, EntityMinValue, Detail) VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}')", id, idExcel, idEntity, date, maxValue, midValue, minValue, detail);
            return sql;
        }

        ////表GroupTable的LoadSql
        //private String GetLoadIntoGroupTableSql(int idExcel, string groupName, int id = 0, string remark = "")
        //{
        //    String sql = String.Format("INSERT OR IGNORE into GroupTable ( ID, Excel_ID, GroupName, Remark ) VALUES ('{0}','{1}','{2}','{3}')", id, idExcel, groupName, remark);
        //    return sql;
        //}

        ////表ExcelTable的InsertSql
        //private String GetInsertIntoExcelTableSql(bool isInfo, string signal, string tableDesc)
        //{
        //    String sql = String.Format("INSERT INTO ExcelTable(TableDesc,ExcelSignal,IsInfo) VALUES ('{0}', '{1}', '{2}' ) ", tableDesc, signal, isInfo);
        //    return sql;
        //}
        ////表KeyTable的InsertSql
        //private String GetInsertIntoKeyTableSql(int idExcel, int idGroup, string keyName)
        //{
        //    String sql = String.Format("INSERT INTO KeyTable ( Excel_ID, Group_ID, KeyName ) values ('{0}','{1}','{2}')", idExcel, idGroup, keyName);
        //    return sql;
        //}
        ////表EntityTable的InsertSql
        //private String GetInsertIntoEntityTableSql(int idExcel, string entityName, string entityRemark = "")
        //{
        //    String sql = String.Format("INSERT INTO EntityTable ( Excel_ID, EntityName, Remark ) VALUES ('{0}','{1}','{2}')", idExcel, entityName, entityRemark);
        //    return sql;
        //}
        //表InfoTable的InsertSql
        //private String GetInsertIntoInfoTableSql(int idKey, int idEntity, string value)
        //{
        //    String sql = String.Format("INSERT INTO InfoTable ( Key_ID, Entity_ID, Value ) VALUES ('{0}','{1}','{2}')", idKey, idEntity, value);
        //    return sql;
        //}

        ////表DrawDataTable的InsertSql
        //private String GetInsertIntoDrawDataTableSql(int idExcel, int idEntity, string date, double maxValue, double midValue, double minValue, string detail)
        //{
        //    String sql = String.Format("INSERT INTO DrawDataTable ( Excel_ID, Entity_ID, Date, EntityMaxValue, EntityMidValue, EntityMinValue, Detail ) VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", idExcel, idEntity, date, maxValue, midValue, minValue, detail);
        //    return sql;
        //}

        ////表GroupTable的InsertSql
        //private String GetInsertIntoGroupTableSql(int idExcel, string groupName = "未分组", string remark = "")
        //{
        //    String sql = String.Format("INSERT INTO GroupTable ( Excel_ID, GroupName, Remark ) VALUES ('{0}','{1}','{2}')", idExcel, groupName, remark);
        //    return sql;
        //}

        ///// <summary>
        ///// 根据signal获取Excel_ID
        ///// </summary>
        ///// <param name="signal"></param>
        ///// <returns></returns>
        //private int getIdExcel(String signal)
        //{
        //    string sql = String.Format("select ID from ExcelTable where ExcelSignal = '{0}'", signal);
        //    int id = getID(sql);
        //    return id;
        //}
        ///// <summary>
        ///// 获取“未分组”Group的ID
        ///// </summary>
        ///// <param name="idExcel"></param>
        ///// <returns></returns>
        //private int getIdGroup(int idExcel)
        //{
        //    string sql = String.Format("select ID from GroupTable where Excel_ID = '{0}' and GroupName = '未分组' ", idExcel);
        //    int id = getID(sql);
        //    return id;
        //}
        ///// <summary>
        ///// 
        ///// 获取Key_ID
        ///// </summary>
        ///// <param name="idExcel"></param>
        ///// <param name="keyName"></param>
        ///// <returns></returns>
        //private int getIdKey(int idExcel, String keyName)
        //{
        //    string sql = String.Format("select ID from KeyTable where Excel_ID = {0} and KeyName = '{1}'", idExcel, keyName);
        //    int id = getID(sql);
        //    return id;
        //}
        ///// <summary>
        ///// 获取Entity_ID
        ///// </summary>
        ///// <param name="idExcel"></param>
        ///// <param name="entityName"></param>
        ///// <returns></returns>
        //private int getIdEntity(int idExcel, string entityName)
        //{
        //    string sql = String.Format("select ID from EntityTable where Excel_ID = {0} and EntityName = '{1}'", idExcel, entityName);
        //    int id = getID(sql);
        //    return id;
        //}
        //private int getID(string sql)
        //{
        //    int id = 0;

        //    SQLiteCommand mycom = new SQLiteCommand();
        //    mycom.Connection = this.conn;
        //    mycom.CommandText = sql;
        //    if (isBegin)
        //        mycom.Transaction = myTran;

        //    using (mycom)
        //    {
        //        try
        //        {
        //            id = Convert.ToInt32(mycom.ExecuteScalar());   //查询返回一个值的时候，用ExecuteScalar()更节约资源，快捷                    
        //        }
        //        catch { throw; }
        //    }
        //    return id;
        //}



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

        ///// <summary>
        ///// 
        /////查询CX中的异常点，返回异常的entity值
        ///// </summary>
        ///// <param name="threshold">累计预警值threshold</param>
        ///// <returns>返回list<string></returns>
        //public List<ParameterData> SelectExceptionalCX1(double threshold)
        //{
        //    List<ParameterData> ExceptionalCX = new List<ParameterData>();
        //    string sql = String.Format("select distinct b.ENTITY from  CXTable b where   b.VALUE > {0} ", threshold);

        //    //判断数据库是否打开
        //    if (conn.State != ConnectionState.Open)
        //    {
        //        OpenConnect();
        //    }
        //    //判断是否创建该查询的表（一定要先打开数据库）
        //    if (!isExist("CXTable", "table"))
        //    {
        //        MessageBox.Show("本地数据库不存在，建议更新本地数据库！");
        //        return ExceptionalCX;
        //    }

        //    using (SQLiteCommand command = new SQLiteCommand(sql, conn)) //建立执行命令语句对象
        //    {
        //        SQLiteDataReader reader = command.ExecuteReader();
        //        try
        //        {
        //            while (reader.Read())
        //            {
        //                if (reader.HasRows)
        //                {
        //                    ExceptionalCX.Add(new ParameterData(reader.GetString(0), reader.GetString(0)));
        //                }
        //            }
        //            return ExceptionalCX;

        //        }
        //        catch (Exception e)
        //        {
        //            throw e;
        //        }
        //        finally
        //        {
        //            reader.Close();  //关闭
        //        }
        //    }
        //}

        ///// <summary>
        ///// 
        /////查询CX中的异常点，返回异常的entity值
        ///// </summary>
        ///// <param name="D_value">相邻差值预警D_value</param>
        ///// <returns>返回list<string></returns>
        //public List<ParameterData> SelectExceptionalCX2(double D_value)
        //{
        //    List<ParameterData> ExceptionalCX = new List<ParameterData>();
        //    string sql = String.Format("select distinct b.ENTITY from CXTable a, CXTable b where a.ENTITY = b.ENTITY and b.ID - a.ID = 1 and b.VALUE - a.VALUE > {0} ", D_value);

        //    //判断数据库是否打开
        //    if (conn.State != ConnectionState.Open)
        //    {
        //        OpenConnect();
        //    }
        //    //判断是否创建该查询的表（一定要先打开数据库）
        //    if (!isExist("CXTable", "table"))
        //    {
        //        MessageBox.Show("本地数据库不存在，建议更新本地数据库！");
        //        return ExceptionalCX;
        //    }

        //    using (SQLiteCommand command = new SQLiteCommand(sql, conn)) //建立执行命令语句对象
        //    {
        //        SQLiteDataReader reader = command.ExecuteReader();
        //        try
        //        {
        //            while (reader.Read())
        //            {
        //                if (reader.HasRows)
        //                {
        //                    ExceptionalCX.Add(new ParameterData(reader.GetString(0), reader.GetString(0)));
        //                }
        //            }
        //            return ExceptionalCX;

        //        }
        //        catch (Exception e)
        //        {
        //            throw e;
        //        }
        //        finally
        //        {
        //            reader.Close();  //关闭
        //        }
        //    }
        //}




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
                        if (!reader.IsDBNull(0)) {
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
            if (!isExist("DrawDataTable", "table"))
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
                sql = string.Format("select Date,EntityMaxValue,EntityMidValue,EntityMinValue,Detail from DrawDataTable dt, EntityTable et where dt.Entity_ID = et.ID and et.EntityName = '{0}'order by Date", EntityName);
            }
            else if (StartDate == null)
            {
                end = ((DateTime)EndDate).ToString("yyyy-MM-dd");
                if (((DateTime)EndDate).Hour >= 12)
                {
                    end += "pm";
                }
                sql = String.Format("select Date,EntityMaxValue,EntityMidValue,EntityMinValue,Detail from DrawDataTable dt, EntityTable et where dt.Entity_ID = et.ID and et.EntityName = '{0}' and dt.date <= '{1}' order by Date", EntityName, end);
            }
            else if (EndDate == null)
            {
                start = ((DateTime)StartDate).ToString("yyyy-MM-dd");
                if (((DateTime)StartDate).Hour >= 12)
                {
                    start += "pm";
                }
                sql = String.Format("select Date,EntityMaxValue,EntityMidValue,EntityMinValue,Detail from DrawDataTable dt, EntityTable et where dt.Entity_ID = et.ID and et.EntityName = '{0}' and dt.date >= '{1}' order by Date", EntityName, start);
            }
            else
            {
                end = ((DateTime)EndDate).ToString("yyyy-MM-dd");
                if (((DateTime)EndDate).Hour >= 12)
                {
                    end += "pm";
                }
                start = ((DateTime)StartDate).ToString("yyyy-MM-dd");
                if (((DateTime)StartDate).Hour >= 12)
                {
                    start += "pm";
                }
                sql = String.Format("select Date,EntityMaxValue,EntityMidValue,EntityMinValue,Detail from DrawDataTable dt, EntityTable et where dt.Entity_ID = et.ID and et.EntityName = '{0}' and dt.date between '{1}' and '{2}' order by Date", EntityName, start, end);
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
                        if (reader.HasRows)
                        {
                            drawData.Date = Convert.ToDateTime(reader.GetString(0));
                            drawData.MaxValue = reader.GetFloat(1);
                            drawData.MidValue = reader.GetFloat(2);
                            drawData.MinValue = reader.GetFloat(3);
                            drawData.Detail = reader.GetString(4);
                            drawEntityData.Data.Add(drawData);
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
            if (Date.Hour >= 12)
            {
                datestr += "pm";
            }

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
                        if (reader.HasRows)
                        {
                            drawData.MaxValue = reader.GetFloat(0);
                            drawData.MidValue = reader.GetFloat(1);
                            drawData.MinValue = reader.GetFloat(2);
                            drawData.Detail = reader.GetString(3);
                        }
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
        /// 查询Total_hold异常点
        ///返回该类型的所有异常点
        /// </summary>
        public List<string> SelectTotalThresholdEntity(string ExcelSignal, float TotalThreshold)
        {
            ////判断是否创建该查询的表（一定要先打开数据库）
            //if (!isExist("DrawDataTable", "table"))
            //{
            //    MessageBox.Show("本地数据库不存在，建议更新本地数据库！");
            //    return null;
            //}
            //List<string> totalThresholdEntity = new List<string>();

            ////判断数据库是否打开
            //if (conn.State != ConnectionState.Open)
            //{
            //    conn.Open();
            //}
            //string sql = String.Format("select distinct EntityName from  DrawDataTable dt, EntityTable et where dt.Entity_ID = et.ID and dt.Excel_ID = '{0}' and dt.EntityMaxValue > {1} ", idExcel, TotalThreshold);

            //using (SQLiteCommand command = new SQLiteCommand(sql, conn)) //建立执行命令语句对象
            //{
            //    SQLiteDataReader reader = command.ExecuteReader();
            //    try
            //    {
            //        while (reader.Read())
            //        {
            //            if (reader.HasRows)
            //            {
            //                totalThresholdEntity.Add(reader.GetString(0));
            //            }
            //        }
            //        return totalThresholdEntity;

            //    }
            //    catch (Exception e)
            //    {
            //        throw e;
            //    }
            //    finally
            //    {
            //        reader.Close();  //关闭
            //    }
            //}

            ///




            throw new NotImplementedException();
        }

        /// <summary>
        /// 查询Diff_hold异常点
        ///返回该类型的所有异常点
        /// </summary>
        public List<string> SelectDiffThresholdEntity(string ExcelSignal, float DiffThreshold)
        {
            ////判断是否创建该查询的表（一定要先打开数据库）
            //if (!isExist("DrawDataTable", "table"))
            //{
            //    MessageBox.Show("本地数据库不存在，建议更新本地数据库！");
            //    return null;
            //}
            //List<string> diffThresholdEntity = new List<string>();
            ////判断数据库是否打开
            //if (conn.State != ConnectionState.Open)
            //{
            //    conn.Open();
            //}

            //string sql = String.Format("select EntityName from EntityTable where ID IN (select distinct a.Entity_ID from  DrawDataTable a, DrawDataTable b where a.Entity_ID = b.Entity_ID and a.Excel_ID = '{0}' and b.Excel_ID = '{1}' and b.ID - a.ID = 1 and b.EntityMaxValue - a.EntityMaxValue > {2} )", idExcel, idExcel, DiffThreshold);

            //using (SQLiteCommand command = new SQLiteCommand(sql, conn)) //建立执行命令语句对象
            //{
            //    SQLiteDataReader reader = command.ExecuteReader();
            //    try
            //    {
            //        while (reader.Read())
            //        {
            //            if (reader.HasRows)
            //            {
            //                diffThresholdEntity.Add(reader.GetString(0));
            //            }
            //        }
            //        return diffThresholdEntity;

            //    }
            //    catch (Exception e)
            //    {
            //        throw e;
            //    }
            //    finally
            //    {
            //        reader.Close();  //关闭
            //    }
            //}


            //


            throw new NotImplementedException();
        }

        /// <summary>
        /// 通过传入的Signal，查询与之对应的所有的测点
        ///传入的Signal应该是测量数据的signal
        /// </summary>
        public List<CEntityName> SelectAllEntities(string ExcelSignal)
        {
            //判断是否创建该查询的表（一定要先打开数据库）
            if (!isExist("EntityTable", "table"))
            {
                MessageBox.Show("本地数据库不存在，建议更新本地数据库！");
                return null;
            }
            List<CEntityName> Entities = new List<CEntityName>();
            string sql = String.Format("select ID,EntityName from  EntityTable where ExcelSignal = '{0}'", ExcelSignal);

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
                        CEntityName one = new CEntityName();
                        one.Id = reader.GetInt32(0);
                        one.EntityName = reader.GetString(1);
                        Entities.Add(one);
                    }
                    return Entities;
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
        public List<ExcelTable> SelectDrawTypes()
        {
            throw new NotImplementedException();
        }
    }
}
