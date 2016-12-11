using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Data.Common;
using System.Data.SQLite;
using MySql.Data.MySqlClient;
using Revit.Addin.RevitTooltip.Intface;
using Revit.Addin.RevitTooltip.Dto;
using static Revit.Addin.RevitTooltip.App;
using System.Windows;
using System.IO;

namespace Revit.Addin.RevitTooltip.Impl
{
    public class SQLiteHelper : ISQLiteHelper
    {
        //单例模式
        private static SQLiteHelper sqliteHelper;
        private MysqlUtil mysql;

        //连接
        private SQLiteConnection conn;
        private string connectionString = string.Empty;
        //private static string dbName = App.settings.SqliteFileName;
        //private static string dbPath = App.settings.SqliteFilePath;

        private static string dbName = "SqliteDB";
        private static string dbPath = "C:/workspace/慧之建项目/项目文档/数据库";


        //事务
        private SQLiteTransaction myTran;
        //是否打开事务
        private bool isBegin = false;
        //是否有备注
        private bool hasEntityRemark = false;

        /// <summary>
        /// 构造函数
        /// </summary>
        private SQLiteHelper()
        {
            setSQLiteDbDirectory(dbPath);
            this.connectionString = "Data Source=" + dbName;
            this.conn = new SQLiteConnection(this.connectionString);
        }
        /// <summary>
        /// 单例模式
        /// </summary>
        /// <returns></returns>
        public static SQLiteHelper CreateInstance()
        {
            if (null == sqliteHelper)
            {
                sqliteHelper = new SQLiteHelper();
            }
            //else if (!App.settings.SqliteFileName.Equals(SQLiteHelper.dbName) ||
            //   !App.settings.SqliteFilePath.Equals(SQLiteHelper.dbPath))
            //{
            //    SQLiteHelper.dbName = App.settings.SqliteFileName;
            //    SQLiteHelper.dbPath = App.settings.SqliteFilePath;
            //    sqliteHelper.Dispose();
            //    sqliteHelper = new SQLiteHelper();
            //}

            return sqliteHelper;
        }
        /// <summary>
        /// 建立SQLite数据库连接
        /// </summary>
        public void OpenConnect()
        {
            if (null == conn)
            {
                this.connectionString = "Data Source=" + dbName;
                this.conn = new SQLiteConnection(this.connectionString);
            }
            conn.Open();
        }
        /// <summary>
        /// 关闭数据库
        /// </summary>
        public void Close()
        {
            conn.Close();
        }
        /// <summary>
        /// 销毁当前对象
        /// </summary>
        public void Dispose()
        {
            conn.Close();
            conn.Dispose();
            GC.SuppressFinalize(this);     //垃圾回收机制跳过this           
        }
        /// <summary>
        /// 设置SQLiteDB目录
        /// </summary>
        /// <param name="path"></param>
        public void setSQLiteDbDirectory(string path)
        {
            System.IO.Directory.SetCurrentDirectory(path);
        }

        /// <summary>
        /// 从Mysql数据库中导入数据
        /// 如果本地有同名的Sqlite数据库文件，先删除，再创建
        /// </summary>
        public void UpdateDB()
        {
            //先关闭数据库连接
            if (conn.State == ConnectionState.Open)
                Close();

            //删除原有本地数据库
            //if (System.IO.File.Exists(Path.Combine(App.settings.SqliteFilePath, App.settings.SqliteFileName)))
            //{
            //    System.IO.File.Delete(Path.Combine(App.settings.SqliteFilePath, App.settings.SqliteFileName));
            //}
            if (System.IO.File.Exists(Path.Combine("C:/workspace/慧之建项目/项目文档/数据库","SqliteDB")))
            {
                System.IO.File.Delete(Path.Combine("C:/workspace/慧之建项目/项目文档/数据库", "SqliteDB"));
            }

            mysql = MysqlUtil.CreateInstance();
            //打开数据库连接       
            if (conn.State != ConnectionState.Open)
            {
                OpenConnect();
            }

            using (DbTransaction transaction = conn.BeginTransaction())
            {
                CreateDB();
                InsertDBfromMysql();
                transaction.Commit();
            }
        }
        //创建SQLite数据库文件
        public void CreateDB()
        {
            if (conn.State != ConnectionState.Open)
            {
                OpenConnect();
            }
            using (SQLiteCommand command = new SQLiteCommand(this.conn))
            {
                if (!isExist("ExcelTable", "table"))
                {
                    command.CommandText = "CREATE TABLE ExcelTable(ID integer NOT NULL PRIMARY KEY AUTOINCREMENT,TableDesc VARCHAR(30) NOT NULL,ExcelSignal VARCHAR(10) UNIQUE, IsInfo BOOLEAN, Total_hold double, Diff_hold double, Remark VARCHAR(100) )";
                    command.ExecuteNonQuery();
                }

                if (!isExist("KeyTable", "table"))
                {
                    command.CommandText = "CREATE TABLE KeyTable(ID integer PRIMARY KEY AUTOINCREMENT, Excel_ID integer NOT NULL REFERENCES ExcelTable(ID) ON DELETE CASCADE ON UPDATE NO ACTION, Group_ID integer, KeyName VARCHAR(20) NOT NULL )";
                    command.ExecuteNonQuery();
                }

                if (!isExist("EntityTable", "table"))
                {
                    command.CommandText = "CREATE TABLE EntityTable(ID integer PRIMARY KEY AUTOINCREMENT, Excel_ID integer NOT NULL REFERENCES ExcelTable(ID) ON DELETE CASCADE ON UPDATE NO ACTION, EntityName VARCHAR(20) NOT NULL, Remark VARCHAR(100))";
                    command.ExecuteNonQuery();
                }

                if (!isExist("GroupTable", "table"))
                {
                    command.CommandText = "CREATE TABLE GroupTable(ID integer PRIMARY KEY AUTOINCREMENT, Excel_ID integer NOT NULL REFERENCES ExcelTable(ID) ON UPDATE NO ACTION, GroupName VARCHAR(20) NOT NULL, Remark VARCHAR(100))";
                    command.ExecuteNonQuery();
                }

                if (!isExist("InfoTable", "table"))
                {
                    command.CommandText = "CREATE TABLE InfoTable(ID integer PRIMARY KEY AUTOINCREMENT, Key_ID integer NOT NULL REFERENCES KeyTable(ID) ON DELETE CASCADE ON UPDATE NO ACTION, Entity_ID integer NOT NULL REFERENCES EntityTable(ID) ON DELETE CASCADE ON UPDATE NO ACTION, Value VARCHAR(20))";
                    command.ExecuteNonQuery();
                }

                if (!isExist("DrawDataTable", "table"))
                {
                    command.CommandText = "CREATE TABLE DrawDataTable(ID integer PRIMARY KEY AUTOINCREMENT, Excel_ID integer NOT NULL REFERENCES ExcelTable(ID) ON DELETE CASCADE ON UPDATE NO ACTION, Entity_ID integer NOT NULL REFERENCES EntityTable(ID) ON DELETE CASCADE ON UPDATE NO ACTION, Date VARCHAR(20), EntityMaxValue double, EntityMidValue double, EntityMinValue double, Detail TEXT )";
                    command.ExecuteNonQuery();
                }

                //if (!isExist("CXTable", "table"))
                //{
                //    command.CommandText = "CREATE TABLE CXTable (ID integer PRIMARY KEY AUTOINCREMENT, ENTITY VARCHAR(20), DATE VARCHAR(20), VALUE FLOAT)";
                //    command.ExecuteNonQuery();
                //}

                //if (!isExist("BaseWallView", "view"))
                //{
                //    command.CommandText = "Create view BaseWallView (Entity, ColumnName, Value, FrameID) AS Select et.ENTITY, ft.COLUMNNAME, bt.VALUE, ft.ID from  EntityTable et, FrameTable ft, BaseComponentTable bt Where et.ID_TYPE = ft.ID_TYPE and bt.ID_ENTITY = et.ID and bt.ID_FRAME = ft.ID "
                //                          + " UNION Select et.ENTITY, ft.COLUMNNAME, wt.VALUE, ft.ID from  EntityTable et, FrameTable ft, wylWallTable wt Where et.ID_TYPE = ft.ID_TYPE and wt.ID_ENTITY = et.ID and wt.ID_FRAME = ft.ID ";
                //    command.ExecuteNonQuery();
                //}

                //if (!isExist("CXView", "view"))
                //{
                //    command.CommandText = "Create view CXView(ID_Entity, Date, Value) AS Select ID_Entity, DATE, MAX(cast(VALUE as DECIMAL(5, 2))) from InclinationTable group by ID_Entity,DATE order by ID_Entity, DATE ";
                //    command.ExecuteNonQuery();
                //}
            }

        }

        //插入数据
        private void InsertDBfromMysql()
        {
            LoadToExcelTable();
            LoadToKeyTable();
            LoadToEntityTable();
            LoadToGroupTable();
            LoadToInfoTable();
            LoadToDrawDataTable();
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
            string sql = String.Format("select count(*) from sqlite_master where type='{0}' and name = '{1}' ", type, table);

            using (SQLiteCommand command = new SQLiteCommand(sql, this.conn))
            {
                if (Convert.ToInt32(command.ExecuteScalar()) > 0)
                    flag = true;
            }
            return flag;
        }
        /// <summary>
        /// 加载Mysql中的ExcelTable数据到Sqlite
        /// </summary>
        private void LoadToExcelTable()
        {
            MySqlDataReader reader = this.mysql.LoadTableData("ExcelTable");

            List<string> SqlStringList = new List<string>();  // 用来存放多条SQL语句
            string sql = "";
            try
            {
                while (reader.Read())
                {
                    if (reader.HasRows)
                    {
                        sql = GetLoadIntoExcelTableSql(reader.GetString(1), reader.GetString(2), reader.GetBoolean(3), reader.GetInt32(0), reader.IsDBNull(4) ? 0 : Convert.ToDouble(reader.GetValue(4)), reader.IsDBNull(5) ? 0 : Convert.ToDouble(reader.GetValue(5)), reader.IsDBNull(5) ? null : Convert.ToString(reader.GetValue(6)));
                        SqlStringList.Add(sql);                        
                    }
                }
                InsertSqlStringList(SqlStringList);
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
        /// <summary>
        /// 加载Mysql中的KeyTable数据到Sqlite
        /// </summary>
        private void LoadToKeyTable()
        {
            MySqlDataReader reader = this.mysql.LoadTableData("KeyTable");
            //string sql = "INSERT OR IGNORE INTO KeyTable(ID,Excel_ID,Group_ID,KeyName)values(@ID,@Excel_ID,@Group_ID,@KeyName)";

            List<string> SqlStringList = new List<string>();  
            string sql = "";
            try
            {
                while (reader.Read())
                {
                    if (reader.HasRows)
                    {
                        sql = GetLoadIntoKeyTableSql(reader.GetInt32(1), reader.GetString(3), reader.GetInt32(0),reader.IsDBNull(2) ? 0 : Convert.ToInt32(reader.GetValue(2)));
                        SqlStringList.Add(sql);

                        //SQLiteParameter[] paramenters = new SQLiteParameter[]
                        //{
                        //    new SQLiteParameter("@ID", reader.GetInt64(0)),
                        //    new SQLiteParameter("@Excel_ID", reader.GetInt64(1)),
                        //    new SQLiteParameter("@Group_ID", reader.GetInt64(2)),
                        //    new SQLiteParameter("@KeyName", reader.GetString(3))
                        //};
                        //ExecuteNonQuery(sql, paramenters);
                    }
                }
                InsertSqlStringList(SqlStringList);
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
        /// <summary>
        /// 加载Mysql中的EntityTable数据到Sqlite
        /// </summary>
        private void LoadToEntityTable()
        {
            MySqlDataReader reader = this.mysql.LoadTableData("EntityTable");
            //string sql = "INSERT OR IGNORE INTO EntityTable(ID,Excel_ID,EntityName,Remark)values(@ID,@Excel_ID,@EntityName,@Remark)";

            List<string> SqlStringList = new List<string>();
            string sql = "";
            try
            {
                while (reader.Read())
                {
                    if (reader.HasRows)
                    {
                        sql = GetLoadIntoEntityTableSql(reader.GetInt32(1), reader.GetString(2), reader.GetInt32(0), reader.IsDBNull(3) ? null : Convert.ToString(reader.GetValue(3)));
                        SqlStringList.Add(sql);
                        
                        //SQLiteParameter[] paramenters = new SQLiteParameter[]
                        //{
                        //    new SQLiteParameter("@ID", reader.GetInt64(0)),
                        //    new SQLiteParameter("@Excel_ID", reader.GetInt64(1)),
                        //    new SQLiteParameter("@EntityName", reader.GetString(2)),
                        //    new SQLiteParameter("@Remark", reader.GetString(3))
                        //};
                        //ExecuteNonQuery(sql, paramenters);
                    }
                }
                InsertSqlStringList(SqlStringList);
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
        /// <summary>
        /// 加载Mysql中的GroupTable数据到Sqlite
        /// </summary>
        private void LoadToGroupTable()
        {
            MySqlDataReader reader = this.mysql.LoadTableData("GroupTable");
            //string sql = "INSERT OR IGNORE INTO GroupTable(ID,Excel_ID,GroupName,Remark)values(@ID,@Excel_ID,@GroupName,@Remark)";

            List<string> SqlStringList = new List<string>();
            string sql = "";
            try
            {
                while (reader.Read())
                {
                    if (reader.HasRows)
                    {
                        sql = GetLoadIntoGroupTableSql(reader.GetInt32(1), reader.GetString(2), reader.GetInt32(0), reader.IsDBNull(3) ? null : Convert.ToString(reader.GetValue(3)));
                        SqlStringList.Add(sql);

                        //SQLiteParameter[] paramenters = new SQLiteParameter[]
                        //{
                        //    new SQLiteParameter("@ID", reader.GetInt64(0)),
                        //    new SQLiteParameter("@Excel_ID", reader.GetInt64(1)),
                        //    new SQLiteParameter("@GroupName", reader.GetString(2)),
                        //    new SQLiteParameter("@Remark", reader.GetString(3))
                        //};
                        //ExecuteNonQuery(sql, paramenters);
                    }
                }
                InsertSqlStringList(SqlStringList);
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
        /// <summary>
        /// 加载Mysql中的InfoTable数据到Sqlite
        /// </summary>
        private void LoadToInfoTable()
        {
            MySqlDataReader reader = this.mysql.LoadTableData("InfoTable");
            //string sql = "INSERT OR IGNORE INTO InfoTable(ID,Key_ID,Entity_ID,Value)values(@ID,@Key_ID,@Entity_ID,@Value)";

            List<string> SqlStringList = new List<string>();
            string sql = "";
            try
            {
                while (reader.Read())
                {
                    if (reader.HasRows)
                    {
                        sql = GetLoadIntoInfoTableSql(reader.GetInt32(1), reader.GetInt32(2), reader.GetString(3), reader.GetInt32(0));
                        SqlStringList.Add(sql);

                        //SQLiteParameter[] paramenters = new SQLiteParameter[]
                        //{
                        //    new SQLiteParameter("@ID", reader.GetInt64(0)),
                        //    new SQLiteParameter("@Key_ID", reader.GetInt64(1)),
                        //    new SQLiteParameter("@Entity_ID", reader.GetInt64(2)),
                        //    new SQLiteParameter("@Value", reader.GetString(3))
                        //};
                        //ExecuteNonQuery(sql, paramenters);
                    }
                }
                InsertSqlStringList(SqlStringList);
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
        /// <summary>
        /// 加载Mysql中的DrawDataTable数据到Sqlite
        /// </summary>
        private void LoadToDrawDataTable()
        {
            MySqlDataReader reader = this.mysql.LoadTableData("DrawDataTable");
            //string sql = "INSERT OR IGNORE INTO DrawDataTable(ID,Excel_ID,Entity_ID,Date,EntityMaxValue,EntityMidValue,EntityMinValue,Detail)values(@ID,@Excel_ID,@Entity_ID,@Date,@EntityMaxValue,@EntityMidValue,@EntityMinValue,@Detail)";

            List<string> SqlStringList = new List<string>();
            string sql = "";
            try
            {
                while (reader.Read())
                {
                    if (reader.HasRows)
                    {
                        DateTime dateTime = reader.GetDateTime(3);
                        string datestr = dateTime.ToString("yyyy-MM-dd");
                        if (dateTime.Hour >= 12)
                        {
                            datestr += "pm";
                        }

                        sql = GetLoadIntoDrawDataTableSql(reader.GetInt32(1), reader.GetInt32(2), datestr, reader.GetDouble(4), reader.GetDouble(5), reader.GetDouble(6),reader.GetString(7), reader.GetInt32(0));
                        SqlStringList.Add(sql);

                        //SQLiteParameter[] paramenters = new SQLiteParameter[]
                        //{
                        //    new SQLiteParameter("@ID", reader.GetInt64(0)),
                        //    new SQLiteParameter("@Excel_ID", reader.GetInt64(1)),
                        //    new SQLiteParameter("@Entity_ID", reader.GetInt64(2)),
                        //    new SQLiteParameter("@Date", datestr),  //reader.GetDateTime(3).ToString("yyyy-MM-dd HH:mm:ss")
                        //    new SQLiteParameter("@EntityMaxValue", reader.GetDouble(4)),
                        //    new SQLiteParameter("@EntityMidValue", reader.GetDouble(5)),
                        //    new SQLiteParameter("@EntityMinValue", reader.GetDouble(6)),
                        //    new SQLiteParameter("@Detail", reader.GetString(7))
                        //};
                        //ExecuteNonQuery(sql, paramenters);
                    }
                }
                InsertSqlStringList(SqlStringList);
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
        ///// <summary>
        ///// 对SQLite数据库执行增删改操作
        ///// </summary>
        ///// <param name="sql"></param>
        ///// <param name="parameters"></param>
        ///// <returns></returns>
        //public int ExecuteNonQuery(string sql, SQLiteParameter[] parameters)
        //{
        //    int affectedRows = 0;
        //    if (conn.State != ConnectionState.Open)
        //    {
        //        OpenConnect();
        //    }

        //    using (SQLiteCommand command = new SQLiteCommand(conn))
        //    {
        //        command.CommandText = sql;
        //        if (parameters != null)
        //        {
        //            command.Parameters.AddRange(parameters);
        //        }
        //        affectedRows = command.ExecuteNonQuery();
        //    }
        //    return affectedRows;
        //}


        //*****************************插入Excel数据***************************

        /// <summary>
        /// 插入SheetInfo
        /// </summary>
        public int InsertSheetInfo(SheetInfo sheetInfo)
        {
            //判断数据库是否打开
            if (conn.State != ConnectionState.Open)
            {
                OpenConnect();
            }
            //if (!System.IO.File.Exists(Path.Combine(App.settings.SqliteFilePath, App.settings.SqliteFileName)))
            if (!System.IO.File.Exists(Path.Combine(dbName, dbPath)))
            {
                CreateDB();
            }
            //if (!isExist("TypeTable", "table"))
            //{
            //    CreateDB();                             
            //}
            //事务开始               
            myTran = conn.BeginTransaction();
            isBegin = true;
            try
            {
                if (null == sheetInfo)
                {
                    return 0;
                }

                if (sheetInfo.Tag)
                {
                    InsertInfoData(sheetInfo);
                }
                else
                {
                    InsertDrawData(sheetInfo);
                }

                myTran.Commit();    //事务提交
                isBegin = false;
                hasEntityRemark = false;
                return 1;
            }
            catch (Exception e)
            {
                myTran.Rollback();    // 事务回滚
                hasEntityRemark = false;
                isBegin = false;
                //删除entitytable表中当前sheet的entity
                DeleteCurrentEntity(sheetInfo, InsertIntoExcelTable(sheetInfo)["IdExcel"]);

                throw new Exception("事务操作出错，系统信息：" + e.Message);
            }
        }
        /// <summary>
        /// 当插入一个sheet数据回滚时，删除当前sheet表的entity
        /// </summary>
        /// <param name="sheetInfo"></param>
        /// <param name="IdExcel"></param>
        private void DeleteCurrentEntity(SheetInfo sheetInfo, int IdExcel)
        {
            Dictionary<string, int> PeriousEntities = SelectEntities(IdExcel);
            string sql = "delete from entitytable where EntityName in (";
            int n = 0;
            foreach (string entityName in sheetInfo.EntityNames)
            {
                if (PeriousEntities.Keys.Contains(entityName))
                {
                    if (n != 0)
                        sql += ",";

                    sql += "'" + entityName + "'";
                    n++;
                }
            }

            if (n == 0)
                sql += "'')";
            else
                sql += ")";

            ExecuteOneSql(sql);
        }

        /// <summary>
        /// 插入基础数据表
        /// </summary>
        /// <param name="sheetInfo"></param>
        private void InsertInfoData(SheetInfo sheetInfo)
        {
            //插入到ExcelTable表,并返回IdExcel和IdGroup
            Dictionary<string, int> id = InsertIntoExcelTable(sheetInfo);

            //插入到KeyTable表，并返回当前表的键值对<属性名，ID>
            Dictionary<string, int> CurrentKeys = InsertIntoKeyTable(sheetInfo.KeyNames, id);

            //插入到EntityTable表，并返回新插入的实体数据的键值对<实体名,ID>
            Dictionary<string, int> UpdateEntities = InsertIntoEntityTable(sheetInfo.EntityNames, id["IdExcel"]);

            //插入到InfoTable表
            InsertIntoInfoTable(sheetInfo, CurrentKeys, UpdateEntities);
        }
        /// <summary>
        /// 插入绘图数据表
        /// </summary>
        /// <param name="sheetInfo"></param>
        private void InsertDrawData(SheetInfo sheetInfo)
        {
            Dictionary<string, int> id = InsertIntoExcelTable(sheetInfo);

            Dictionary<string, int> UpdateEntities = InsertIntoEntityTable(sheetInfo.EntityNames, id["IdExcel"]);

            InsertIntoDrawDataTable(sheetInfo, UpdateEntities, id["IdExcel"]);
        }

        /// <summary>
        /// 插入到ExcelTable表
        /// 并返回IdExcel和IdGroup
        /// </summary>
        /// <param name="sheetInfo"></param>
        /// <returns></returns>
        private Dictionary<string, int> InsertIntoExcelTable(SheetInfo sheetInfo)
        {
            Dictionary<string, int> id = new Dictionary<string, int>();
            bool isInfo = sheetInfo.Tag;
            string signal = sheetInfo.ExcelTableData.Signal;
            string tableDesc = sheetInfo.ExcelTableData.TableDesc;

            int IdExcel = getIdExcel(signal);
            int IdGroup = getIdGroup(IdExcel);
            string currentTableDesc = GetTableDesc(signal);
            bool hasTableDesc = false;

            if (IdExcel == 0)
            {
                string sql = GetInsertIntoExcelTableSql(isInfo, signal, tableDesc);
                ExecuteOneSql(sql);
                IdExcel = getIdExcel(signal);

                //插入到GroupTable
                IdGroup = InsertIntoGroupTable(IdExcel);
            }
            else
            {
                string[] tableDesclist = currentTableDesc.Split(';');
                foreach (string t in tableDesclist)
                {
                    if (t.Equals(tableDesc))
                    {
                        hasTableDesc = true;
                    }
                }
                if (!hasTableDesc)
                {
                    string updateTableDesc = currentTableDesc + tableDesc;
                    UpdateTableDesc(signal, updateTableDesc);
                }
            }
            id.Add("IdExcel", IdExcel);
            id.Add("IdGroup", IdGroup);

            return id;
        }
        /// <summary>
        /// 插入“未分组”Group
        /// </summary>
        /// <param name="idExcel"></param>
        /// <returns></returns>
        private int InsertIntoGroupTable(int idExcel)
        {
            int idGroup = 0;
            string sql = GetInsertIntoGroupTableSql(idExcel) + ";select @@IDENTITY ";

            SQLiteCommand mycom = new SQLiteCommand();
            mycom.Connection = this.conn;
            mycom.CommandText = sql;
            if (isBegin)
                mycom.Transaction = myTran;

            using ( mycom )
            {
                try
                {
                    idGroup = Convert.ToInt32(mycom.ExecuteScalar());
                    return idGroup;
                }
                catch { throw; }
            }

        }
        /// <summary>
        /// 插入属性到KeyTable表
        /// 返回当前表的属性及对应ID，<KeyName,ID>
        /// </summary>
        /// <param name="KeyNames"></param>
        /// <param name="IdExcel"></param>
        /// <returns></returns>
        private Dictionary<string, int> InsertIntoKeyTable(List<string> KeyNames, Dictionary<string, int> id)
        {
            int IdExcel = id["IdExcel"];
            int IdGroup = id["IdGroup"];
            Dictionary<string, int> PeriousKeys = SelectKeyNames(IdExcel);
            String sql = "";
            int n = 0;
            foreach (string name in KeyNames)
            {
                if (name.Equals("备注") || PeriousKeys.Keys.Contains("备注"))
                {
                    hasEntityRemark = true;
                    continue;
                }

                //如果没有在原表中匹配到，则添加到插入语句当中
                if (!PeriousKeys.Keys.Contains(name))
                {
                    if (n == 0)
                        sql = GetInsertIntoKeyTableSql(IdExcel, IdGroup, name);
                    else
                        sql = sql + ",('" + IdExcel + "','" + IdGroup + "','" + name + "')";
                    n++;
                }
            }
            if (!hasEntityRemark)
            {
                sql = sql + ",('" + IdExcel + "','" + IdGroup + "','备注')";
            }
            ExecuteOneSql(sql);
            //获取更新后的<属性，ID>
            Dictionary<string, int> CurrentKeys = SelectKeyNames(IdExcel);
            return CurrentKeys;
        }

        /// <summary>
        /// 插入实体到EntityTable表
        /// 返回当前插入的EntityNames,<EntityName,ID> 
        /// </summary>
        /// <param name="EntityNames"></param>
        /// <param name="IdExcel"></param>
        /// <returns></returns>
        private Dictionary<string, int> InsertIntoEntityTable(List<string> EntityNames, int IdExcel)
        {
            Dictionary<string, int> UpdateEntities = new Dictionary<string, int>();
            List<string> UpdateEntityNames = new List<string>();
            Dictionary<string, int> PeriousEntities = SelectEntities(IdExcel);
            String sql = "";
            int n = 0;
            foreach (string entityName in EntityNames)
            {
                //如果没有在原表中匹配到，则添加到插入语句当中
                if (!PeriousEntities.Keys.Contains(entityName))
                {
                    if (n == 0)
                        sql = GetInsertIntoEntityTableSql(IdExcel, entityName);
                    else
                        sql = sql + ",('" + IdExcel + "','" + entityName + "','" + "')";

                    UpdateEntityNames.Add(entityName);
                    n++;
                }
            }
            ExecuteOneSql(sql);

            //获取当前表所有的Entities
            foreach (string s in UpdateEntityNames)
            {
                int id = getIdEntity(IdExcel, s);
                UpdateEntities.Add(s, id);
            }

            return UpdateEntities;
        }
        /// <summary>
        /// 插入到InfoTable
        /// </summary>
        /// <param name="sheetInfo"></param>
        /// <param name="CurrentKeys"></param>
        /// <param name="UpdateEntities"></param>
        private void InsertIntoInfoTable(SheetInfo sheetInfo, Dictionary<string, int> CurrentKeys, Dictionary<string, int> UpdateEntities)
        {
            int IdKey;
            int IdEntity;
            List<string> SqlStringList = new List<string>();  // 用来存放多条SQL语句
            string sql = "";
            int n = 0;
            foreach (InfoEntityData infoRow in sheetInfo.InfoRows)
            {
                if (UpdateEntities.Keys.Contains(infoRow.EntityName))
                {
                    IdEntity = UpdateEntities[infoRow.EntityName];
                    foreach (string key in infoRow.Data.Keys)
                    {
                        IdKey = CurrentKeys[key];
                        if (n % 500 == 0)
                        {
                            if (n > 0)
                                SqlStringList.Add(sql);
                            sql = GetInsertIntoInfoTableSql(IdKey, IdEntity, infoRow.Data[key]);
                        }
                        else
                            sql = sql + ",('" + IdKey + "','" + IdEntity + "','" + infoRow.Data[key] + "')";

                        n++;
                    }
                }
            }
            if (n % 500 != 0)
                SqlStringList.Add(sql);  //不够500values的SQL语言也要添加进去
            InsertSqlStringList(SqlStringList);

        }
        /// <summary>
        /// 插入到DrawDataTabel
        /// </summary>
        /// <param name="sheetInfo"></param>
        /// <param name="UpdateEntities"></param>
        /// <param name="IdExcel"></param>
        private void InsertIntoDrawDataTable(SheetInfo sheetInfo, Dictionary<string, int> UpdateEntities, int IdExcel)
        {
            int IdEntity;
            List<string> SqlStringList = new List<string>();  // 用来存放多条SQL语句
            String sql = "";
            int n = 0;

            foreach (DrawEntityData drawRow in sheetInfo.DrawRows)
            {
                if (UpdateEntities.Keys.Contains(drawRow.EntityName))
                {
                    IdEntity = UpdateEntities[drawRow.EntityName];
                    foreach (DrawData data in drawRow.Data)
                    {
                        string datestr = data.Date.ToString("yyyy-MM-dd HH:mm:ss");
                        if (n % 500 == 0)
                        {
                            if (n > 0)
                                SqlStringList.Add(sql);
                            sql = GetInsertIntoDrawDataTableSql(IdExcel, IdEntity, datestr, data.MaxValue, data.MidValue, data.MinValue, data.Detail);
                        }
                        else
                            sql = sql + ",('" + IdExcel + "','" + IdEntity + "','" + datestr + "','" + data.MaxValue + "','" + data.MidValue + "','" + data.MinValue + "','" + data.Detail + "')";

                        n++;
                    }
                }
                else  //实体表如果没有更新，要检查每个实体表中，时间是否增加
                {
                    IdEntity = getIdEntity(IdExcel, drawRow.EntityName);
                    List<string> PeriousDateTimes = GetDrawDataTableDateTimes(IdEntity);
                    foreach (DrawData data in drawRow.Data)
                    {
                        string datestr = data.Date.ToString("yyyy-MM-dd HH:mm:ss");
                        if (!PeriousDateTimes.Contains(datestr))
                        {                           
                            if (n % 500 == 0)
                            {
                                if (n > 0)
                                    SqlStringList.Add(sql);
                                sql = GetInsertIntoDrawDataTableSql(IdExcel, IdEntity, datestr, data.MaxValue, data.MidValue, data.MinValue, data.Detail);
                            }
                            else
                                sql = sql + ",('" + IdExcel + "','" + IdEntity + "','" + datestr + "','" + data.MaxValue + "','" + data.MidValue + "','" + data.MinValue + "','" + data.Detail + "')";

                            n++;
                        }
                    }
                }
            }
            if (n % 500 != 0)
                SqlStringList.Add(sql);  //不够500values的SQL语言也要添加进去            
            InsertSqlStringList(SqlStringList);
        }

        //private void InsertIntoInclinationTable(SheetInfo sheetInfo, Dictionary<string, int> CurrentTableColumns, Dictionary<string, int> UpdateEntities, int IdType)
        //{
        //    int IdFrame;
        //    int IdEntity;
        //    int count = sheetInfo.Names.Count;
        //    List<string> SqlStringList = new List<string>();  // 用来存放多条SQL语句
        //    String sql = "";
        //    int n = 0;

        //    foreach (String key in sheetInfo.Data.Keys)
        //    {
        //        //对于测斜数据，一个实体数据是一个表，实体表如果有更新，则在InsertIntoInclinationTable插入该实体对应的一个表的数据
        //        if (UpdateEntities.Keys.Contains(key))
        //        {
        //            IdEntity = UpdateEntities[key];
        //            foreach (String[] sts in sheetInfo.Data[key])
        //            {
        //                for (int i = 0; i < count; i++)
        //                {
        //                    IdFrame = CurrentTableColumns[sheetInfo.Names.ElementAt(i)];
        //                    string datestr = DateTime.Parse(sts[0]).ToString("yyyy-MM-dd HH:mm:ss");

        //                    if (n % 500 == 0)
        //                    {
        //                        if (n > 0)
        //                            SqlStringList.Add(sql);
        //                        sql = GetInsertIntoInclinationTableSql(IdFrame, IdEntity, datestr, string.IsNullOrEmpty(sts[i + 1]) ? "0" : Convert.ToDouble(sts[i + 1]).ToString("f2"));
        //                    }
        //                    else
        //                        sql = sql + ",('" + IdFrame + "','" + IdEntity + "','" + datestr + "','" + (string.IsNullOrEmpty(sts[i + 1]) ? "0" : Convert.ToDouble(sts[i + 1]).ToString("f2")) + "')";

        //                    n++;

        //                }
        //            }
        //        }
        //        else  //实体表如果没有更新，要检查每个实体表中，时间是否增加
        //        {
        //            IdEntity = getIdEntity(IdType, key);
        //            List<DateTime> PeriousDateTimes = GetInclinationTableDateTimes(IdEntity);

        //            foreach (String[] sts in sheetInfo.Data[key])
        //            {
        //                if (!PeriousDateTimes.Contains(DateTime.Parse(sts[0])))
        //                {
        //                    for (int i = 0; i < count; i++)
        //                    {
        //                        IdFrame = CurrentTableColumns[sheetInfo.Names.ElementAt(i)];
        //                        string datestr = DateTime.Parse(sts[0]).ToString("yyyy-MM-dd HH:mm:ss");
        //                        if (n % 500 == 0)
        //                        {
        //                            if (n > 0)
        //                                SqlStringList.Add(sql);
        //                            sql = GetInsertIntoInclinationTableSql(IdFrame, IdEntity, datestr, string.IsNullOrEmpty(sts[i + 1]) ? "0" : Convert.ToDouble(sts[i + 1]).ToString("f2"));  // Convert.ToDouble(sts[i + 1]).ToString("f2")
        //                        }
        //                        else
        //                            sql = sql + ",('" + IdFrame + "','" + IdEntity + "','" + datestr + "','" + (string.IsNullOrEmpty(sts[i + 1]) ? "0" : Convert.ToDouble(sts[i + 1]).ToString("f2")) + "')";

        //                        n++;

        //                    }
        //                }
        //            }
        //        }
        //    }

        //    if (n % 500 != 0)
        //        SqlStringList.Add(sql);  //不够500values的SQL语言也要添加进去            
        //    InsertSqlStringList(SqlStringList);
        //}
        /// <summary>
        /// 获取KeyTable表的Key和ID
        /// </summary>
        /// <param name="IdExcel"></param>
        /// <returns></returns>
        private Dictionary<string, int> SelectKeyNames(int IdExcel)
        {
            Dictionary<string, int> KeyNames = new Dictionary<string, int>();
            String sql = "select KeyName, ID from KeyTable where Excel_ID = " + IdExcel;

            SQLiteCommand mycom = new SQLiteCommand(sql, this.conn, myTran);  //建立执行命令语句对象
            SQLiteDataReader reader = mycom.ExecuteReader();    //需要关闭
            try
            {
                while (reader.Read())
                {
                    if (reader.HasRows)
                    {
                        KeyNames.Add(reader.GetString(0), reader.GetInt32(1));
                    }
                }
                return KeyNames;
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
        /// <summary>
        /// 获取EntityTable表的EntityName和ID
        /// </summary>
        /// <param name="IdExcel"></param>
        /// <returns></returns>
        private Dictionary<string, int> SelectEntities(int IdExcel)
        {
            Dictionary<string, int> Entities = new Dictionary<string, int>();
            string sql = "select distinct EntityName, ID from EntityTable where Excel_ID = " + IdExcel;

            SQLiteCommand mycom = new SQLiteCommand();
            mycom.Connection = this.conn;
            mycom.CommandText = sql;
            if (isBegin)
                mycom.Transaction = myTran;
            SQLiteDataReader reader = mycom.ExecuteReader();    //需要关闭
            try
            {
                while (reader.Read())
                {
                    if (reader.HasRows)
                    {
                        //这里如果添加了相同的项会出错，所以entity一定要唯一
                        Entities.Add(reader.GetString(0), reader.GetInt32(1));
                    }
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
        /// <summary>
        /// 从DrawDataTable表中，获取该实体数据中，当前存储的DATETIME项
        /// </summary>
        /// <param name="IdEntity"></param>
        /// <returns></returns>
        private List<string> GetDrawDataTableDateTimes(int IdEntity)
        {
            List<string> PeriousDateTimes = new List<string>();
            String sql = "select Date from DrawDataTable where Entity_ID = " + IdEntity;

            SQLiteCommand mycom = new SQLiteCommand(sql, this.conn, myTran);  //建立执行命令语句对象
            SQLiteDataReader reader = mycom.ExecuteReader();    //需要关闭
            try
            {
                while (reader.Read())
                {
                    if (reader.HasRows)
                    {
                        PeriousDateTimes.Add(reader.GetString(0));
                    }
                }
                return PeriousDateTimes;
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

        /// <summary>
        /// 获取当前Excel的TableDesc
        /// </summary>
        /// <param name="signal"></param>
        /// <returns></returns>
        private string GetTableDesc(string signal)
        {
            string currentTableDesc = "";
            string sql = String.Format("select tableDesc from ExcelTable where ExcelSignal = '{0}' ", signal);

            SQLiteCommand mycom = new SQLiteCommand();
            mycom.Connection = this.conn;
            mycom.CommandText = sql;
            if (isBegin)
                mycom.Transaction = myTran;
            
            SQLiteDataReader reader = mycom.ExecuteReader();    //需要关闭
            try
            {
                while (reader.Read())
                {
                    if (reader.HasRows)
                    {
                        currentTableDesc = reader.GetString(0);
                    }
                }
                return currentTableDesc;
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
        /// <summary>
        /// 更新TableDesc
        /// </summary>
        /// <param name="signal"></param>
        /// <param name="updateTableDesc"></param>
        private void UpdateTableDesc(string signal, string updateTableDesc)
        {
            string sql = String.Format("update ExcelTable set tableDesc = '{0}' where ExcelSignal = '{1}' ", updateTableDesc, signal);
            ExecuteOneSql(sql);
        }
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



        /// <summary>
        /// 执行批量插入
        /// </summary>
        private void InsertSqlStringList(List<string> SqlStringList)
        {
            SQLiteCommand command = new SQLiteCommand();
            command.Connection = this.conn;
            command.Transaction = myTran;
            try
            {
                for (int i = 0; i < SqlStringList.Count; i++)
                {
                    string sql = SqlStringList[i].ToString();
                    if (sql.Equals(""))
                        return;
                    command.CommandText = sql;
                    command.ExecuteNonQuery();
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        /// <summary>
        /// 执行单个sql语句，返回插入结果
        /// </summary>
        /// <param name="sql"></param>
        private void ExecuteOneSql(string sql)
        {
            if (sql.Equals(""))
                return;
            SQLiteCommand mycom = new SQLiteCommand();
            mycom.Connection = this.conn;
            mycom.CommandText = sql;
            if (isBegin)
                mycom.Transaction = myTran;
            
            try
            {
                mycom.ExecuteNonQuery();
            }
            catch (Exception e)
            {
                throw e;
            }
        }
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
        private String GetLoadIntoExcelTableSql(string tableDesc, string signal,bool isInfo, int id = 0, double total_hold = 0, double diff_hold = 0, string remark = null)
        {
            String sql = String.Format("INSERT OR IGNORE INTO ExcelTable(ID,TableDesc,ExcelSignal,IsInfo,Total_hold,Diff_hold,Remark) VALUES ('{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}')", id, tableDesc, signal, isInfo, total_hold, diff_hold, remark);
            return sql;
        }
        //表KeyTable的LoadSql
        private String GetLoadIntoKeyTableSql(int idExcel, string keyName,int id = 0, int idGroup = 0)
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

        //表GroupTable的LoadSql
        private String GetLoadIntoGroupTableSql(int idExcel, string groupName, int id = 0, string remark = "")
        {
            String sql = String.Format("INSERT OR IGNORE into GroupTable ( ID, Excel_ID, GroupName, Remark ) VALUES ('{0}','{1}','{2}','{3}')", id, idExcel, groupName, remark);
            return sql;
        }

        //表ExcelTable的InsertSql
        private String GetInsertIntoExcelTableSql(bool isInfo, string signal, string tableDesc)
        {
            String sql = String.Format("INSERT INTO ExcelTable(TableDesc,ExcelSignal,IsInfo) VALUES ('{0}', '{1}', '{2}' ) ", tableDesc, signal, isInfo);
            return sql;
        }
        //表KeyTable的InsertSql
        private String GetInsertIntoKeyTableSql(int idExcel,int idGroup, string keyName)
        {
            String sql = String.Format("INSERT INTO KeyTable ( Excel_ID, Group_ID, KeyName ) values ('{0}','{1}','{2}')", idExcel, idGroup, keyName);
            return sql;
        }
        //表EntityTable的InsertSql
        private String GetInsertIntoEntityTableSql(int idExcel, string entityName, string entityRemark = "")
        {
            String sql = String.Format("INSERT INTO EntityTable ( Excel_ID, EntityName, Remark ) VALUES ('{0}','{1}','{2}')", idExcel, entityName, entityRemark);
            return sql;
        }
        //表InfoTable的InsertSql
        private String GetInsertIntoInfoTableSql(int idKey, int idEntity, string value)
        {
            String sql = String.Format("INSERT INTO InfoTable ( Key_ID, Entity_ID, Value ) VALUES ('{0}','{1}','{2}')", idKey, idEntity, value);
            return sql;
        }

        //表DrawDataTable的InsertSql
        private String GetInsertIntoDrawDataTableSql(int idExcel, int idEntity, string date, double maxValue, double midValue, double minValue, string detail)
        {
            String sql = String.Format("INSERT INTO DrawDataTable ( Excel_ID, Entity_ID, Date, EntityMaxValue, EntityMidValue, EntityMinValue, Detail ) VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", idExcel, idEntity, date, maxValue, midValue, minValue, detail);
            return sql;
        }

        //表GroupTable的InsertSql
        private String GetInsertIntoGroupTableSql(int idExcel, string groupName = "未分组", string remark = "")
        {
            String sql = String.Format("INSERT INTO GroupTable ( Excel_ID, GroupName, Remark ) VALUES ('{0}','{1}','{2}')", idExcel, groupName, remark);
            return sql;
        }

        /// <summary>
        /// 根据signal获取Excel_ID
        /// </summary>
        /// <param name="signal"></param>
        /// <returns></returns>
        private int getIdExcel(String signal)
        {
            string sql = String.Format("select ID from ExcelTable where ExcelSignal = '{0}'", signal);
            int id = getID(sql);
            return id;
        }
        /// <summary>
        /// 获取“未分组”Group的ID
        /// </summary>
        /// <param name="idExcel"></param>
        /// <returns></returns>
        private int getIdGroup(int idExcel)
        {
            string sql = String.Format("select ID from GroupTable where Excel_ID = '{0}' and GroupName = '未分组' ", idExcel);
            int id = getID(sql);
            return id;
        }
        /// <summary>
        /// 
        /// 获取Key_ID
        /// </summary>
        /// <param name="idExcel"></param>
        /// <param name="keyName"></param>
        /// <returns></returns>
        private int getIdKey(int idExcel, String keyName)
        {
            string sql = String.Format("select ID from KeyTable where Excel_ID = {0} and KeyName = '{1}'", idExcel, keyName);
            int id = getID(sql);
            return id;
        }
        /// <summary>
        /// 获取Entity_ID
        /// </summary>
        /// <param name="idExcel"></param>
        /// <param name="entityName"></param>
        /// <returns></returns>
        private int getIdEntity(int idExcel, string entityName)
        {
            string sql = String.Format("select ID from EntityTable where Excel_ID = {0} and EntityName = '{1}'", idExcel, entityName);
            int id = getID(sql);
            return id;
        }
        private int getID(string sql)
        {
            int id = 0;

            SQLiteCommand mycom = new SQLiteCommand();
            mycom.Connection = this.conn;
            mycom.CommandText = sql;
            if (isBegin)
                mycom.Transaction = myTran;

            using (mycom)
            {               
                try
                {
                    id = Convert.ToInt32(mycom.ExecuteScalar());   //查询返回一个值的时候，用ExecuteScalar()更节约资源，快捷                    
                }
                catch { throw; }
            }
            return id;
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
                OpenConnect();
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
                OpenConnect();
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
        /// 
        ///查询CX中的异常点，返回异常的entity值
        /// </summary>
        /// <param name="threshold">累计预警值threshold</param>
        /// <returns>返回list<string></returns>
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
            InfoEntityData infoData = new InfoEntityData();
            infoData.EntityName = EntityName;
            infoData.Data = new Dictionary<string, string>();
            infoData.GroupMsg = new Dictionary<string, List<string>>();

            if (conn.State != ConnectionState.Open)
            {
                OpenConnect();
            }
            //判断是否创建该查询的表（一定要先打开数据库）
            if (!isExist("InfoTable", "table"))
            {
                MessageBox.Show("本地数据库不存在，建议更新本地数据库！");
                return infoData;
            }
      
            string getDataSql = String.Format("select KeyName, Value from KeyTable kt, EntityTable et,InfoTable it where kt.Excel_ID = et.Excel_ID and it.Key_ID = kt.ID and it.Entity_ID = et.ID and et.EntityName = '{0}'", EntityName);
            string getGroupMsgSql = String.Format("select GroupName, KeyName from GroupTable gt, KeyTable kt, EntityTable et where kt.Group_ID = gt.ID and gt.Excel_ID = kt.Excel_ID and kt.Excel_ID = et.Excel_ID and et.EntityName = '{0}' order by GroupName ", EntityName);

            using (SQLiteCommand command = new SQLiteCommand(getDataSql, conn)) 
            {
                SQLiteDataReader reader = command.ExecuteReader();
                try
                {
                    while (reader.Read())
                    {
                        if (reader.HasRows)
                        {
                            infoData.Data.Add(reader.GetString(0), reader.GetString(1));
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
                    List<string> list = new List<string>();
                    string groupName = "";
                    while (reader.Read())
                    {
                        if (reader.HasRows)
                        {                                                       
                            if (!groupName.Equals(reader.GetString(0)))
                            {                                
                                if (list.Count != 0)
                                {
                                    infoData.GroupMsg.Add(groupName, list);
                                    list = new List<string>();
                                }
                                groupName = reader.GetString(0);
                                list.Add(reader.GetString(1));
                            }
                            else
                            {
                                list.Add(reader.GetString(1));
                            }
                        }
                    }
                    if(!groupName.Equals(""))
                        infoData.GroupMsg.Add(groupName, list);
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
            return infoData;
        }

        /// <summary>
        /// 查询DrawDataTable
        ///查询Entity时间序列数据
        ///根据传入的起始时间查询
        /// </summary>
        public DrawEntityData SelectDrawEntityData(string EntityName, DateTime StartDate, DateTime EndDate)
        {
            DrawEntityData drawEntityData = new DrawEntityData();
            drawEntityData.EntityName = EntityName;
            drawEntityData.Data = new List<DrawData>();

            string startDate = StartDate.ToString("yyyy-MM-dd");
            if (StartDate.Hour >= 12)
            {
                startDate += "pm";
            }

            string endDate = EndDate.ToString("yyyy-MM-dd");
            if (EndDate.Hour >= 12)
            {
                endDate += "pm";
            }

            string sql = String.Format("select Date,EntityMaxValue,EntityMidValue,EntityMinValue,Detail from DrawDataTable dt, EntityTable et where dt.Entity_ID = et.ID and et.EntityName = '{0}' and dt.date between '{1}' and '{2}' order by Date", EntityName, startDate, endDate);

            if (conn.State != ConnectionState.Open)
            {
                OpenConnect();
            }
            //判断是否创建该查询的表（一定要先打开数据库）
            if (!isExist("DrawDataTable", "table"))
            {
                MessageBox.Show("本地数据库不存在，建议更新本地数据库！");
                return drawEntityData;
            }
            using (SQLiteCommand command = new SQLiteCommand(sql, conn)) //建立执行命令语句对象
            {
                SQLiteDataReader reader = command.ExecuteReader();
                try
                {
                    DrawData drawData;
                    while (reader.Read())
                    {                        
                        if (reader.HasRows)
                        {
                            drawData = new DrawData();
                            drawData.Date = Convert.ToDateTime(reader.GetString(0));
                            drawData.MaxValue = reader.GetDouble(1);
                            drawData.MidValue = reader.GetDouble(2);
                            drawData.MinValue = reader.GetDouble(3);
                            drawData.Detail = reader.GetString(4);
                            drawEntityData.Data.Add(drawData);
                        }
                    }
                    return drawEntityData;
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
        /// 查询DrawDataTable
        ///查询Entity某日期的数据
        /// </summary>
        public DrawData SelectDrawData(string EntityName, DateTime Date)
        {
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
                OpenConnect();
            }
            //判断是否创建该查询的表（一定要先打开数据库）
            if (!isExist("DrawDataTable", "table"))
            {
                MessageBox.Show("本地数据库不存在，建议更新本地数据库！");
                return drawData;
            }
            using (SQLiteCommand command = new SQLiteCommand(sql, conn)) //建立执行命令语句对象
            {
                SQLiteDataReader reader = command.ExecuteReader();
                try
                {
                    while (reader.Read())
                    {
                        if (reader.HasRows)
                        {
                            drawData.MaxValue = reader.GetDouble(0);
                            drawData.MidValue = reader.GetDouble(1);
                            drawData.MinValue = reader.GetDouble(2);
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
        public List<string> SelectTotalThresholdEntity(string ExcelSignal, double TotalThreshold)
        {
            List<string> totalThresholdEntity = new List<string>();

            //判断数据库是否打开
            if (conn.State != ConnectionState.Open)
            {
                OpenConnect();
            }
            //判断是否创建该查询的表（一定要先打开数据库）
            if (!isExist("DrawDataTable", "table"))
            {
                MessageBox.Show("本地数据库不存在，建议更新本地数据库！");
                return totalThresholdEntity;
            }

            int idExcel = getIdExcel(ExcelSignal);
            if(idExcel == 0)
            {
                MessageBox.Show("不存在该符号的数据表，建议更新本地数据库！");
                return totalThresholdEntity;
            }
            string sql = String.Format("select distinct EntityName from  DrawDataTable dt, EntityTable et where dt.Entity_ID = et.ID and dt.Excel_ID = '{0}' and dt.EntityMaxValue > {1} ", idExcel, TotalThreshold);

            using (SQLiteCommand command = new SQLiteCommand(sql, conn)) //建立执行命令语句对象
            {
                SQLiteDataReader reader = command.ExecuteReader();
                try
                {
                    while (reader.Read())
                    {
                        if (reader.HasRows)
                        {
                            totalThresholdEntity.Add(reader.GetString(0));
                        }
                    }
                    return totalThresholdEntity;

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
        /// 查询Diff_hold异常点
        ///返回该类型的所有异常点
        /// </summary>
        public List<string> SelectDiffThresholdEntity(string ExcelSignal, double DiffThreshold)
        {
            List<string> diffThresholdEntity = new List<string>();
            //判断数据库是否打开
            if (conn.State != ConnectionState.Open)
            {
                OpenConnect();
            }
            //判断是否创建该查询的表（一定要先打开数据库）
            if (!isExist("DrawDataTable", "table"))
            {
                MessageBox.Show("本地数据库不存在，建议更新本地数据库！");
                return diffThresholdEntity;
            }

            int idExcel = getIdExcel(ExcelSignal);
            if (idExcel == 0)
            {
                MessageBox.Show("不存在该符号的数据表，建议更新本地数据库！");
                return diffThresholdEntity;
            }
            string sql = String.Format("select EntityName from EntityTable where ID IN (select distinct a.Entity_ID from  DrawDataTable a, DrawDataTable b where a.Entity_ID = b.Entity_ID and a.Excel_ID = '{0}' and b.Excel_ID = '{1}' and b.ID - a.ID = 1 and (b.EntityMaxValue - a.EntityMaxValue > {2} or a.EntityMaxValue - b.EntityMaxValue > {3}) )", idExcel, idExcel, DiffThreshold, DiffThreshold);

            using (SQLiteCommand command = new SQLiteCommand(sql, conn)) //建立执行命令语句对象
            {
                SQLiteDataReader reader = command.ExecuteReader();
                try
                {
                    while (reader.Read())
                    {
                        if (reader.HasRows)
                        {
                            diffThresholdEntity.Add(reader.GetString(0));
                        }
                    }
                    return diffThresholdEntity;

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
        /// 通过传入的Signal，查询与之对应的所有的测点
        ///传入的Signal应该是测量数据的signal
        /// </summary>
        public List<string> SelectAllEntities(string ExcelSignal)
        {
            List<string> Entities = new List<string>();
            string sql = String.Format("select EntityName from  EntityTable, ExcelTable  where EntityTable.Excel_ID = ExcelTable.ID and ExcelTable.ExcelSignal = '{0}'", ExcelSignal);

            if (conn.State != ConnectionState.Open)
            {
                OpenConnect();
            }
            //判断是否创建该查询的表（一定要先打开数据库）
            if (!isExist("EntityTable", "table"))
            {
                MessageBox.Show("本地数据库不存在，建议更新本地数据库！");
                return Entities;
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
                            Entities.Add(reader.GetString(0));
                        }
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

        /// <summary>
        /// 复制MySQL中的数据到Sqlite中
        /// </summary>
        public bool LoadDataToSqlite()
        {
            UpdateDB();
            return true;
        }

    }
}
