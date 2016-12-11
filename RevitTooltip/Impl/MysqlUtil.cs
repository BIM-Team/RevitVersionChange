using System;
using System.Collections.Generic;
using System.Linq;
using Revit.Addin.RevitTooltip.Dto;
using Revit.Addin.RevitTooltip.Intface;
using MySql.Data.MySqlClient;
using System.Windows;

namespace Revit.Addin.RevitTooltip.Impl
{
    public class MysqlUtil : IDisposable, IMysqlUtil
    {
        //单例模式
        private static MysqlUtil mysqlUtil;
        //用户名
        // private static String user;
        //密码
        // private static String password;
        //settings
        private static RevitTooltip settings;
        //连接
        private MySqlConnection conn;
        //是否连接的标志符
        private bool isOpen = false;
        //事务
        private MySqlTransaction myTran;

        //是否打开事务
        private bool isBegin = false;

        //是否有备注
        private bool hasEntityRemark = false;

        /// <summary>
        /// 单例模式
        /// </summary>
        /// <returns></returns>
        public static MysqlUtil CreateInstance()
        {
            //if (null == mysqlUtil)
            //{
            //    mysqlUtil = new MysqlUtil(App.settings);
            //}
            //else if (!MysqlUtil.settings.Equals(App.settings))
            //{
            //    mysqlUtil.Dispose();
            //    MysqlUtil.settings = App.settings;
            //    mysqlUtil = new MysqlUtil(settings);
            //}
            if( mysqlUtil == null)
            {
                mysqlUtil = new MysqlUtil();
            }           

            return mysqlUtil;
        }

        /// <summary>
        /// 初始化
        /// </summary>
        /// <param name="settings"></param>
        private MysqlUtil(RevitTooltip settings)
        {
            MysqlUtil.settings = settings;

        }
        /// <summary>
        /// 初始化
        /// </summary>
        private MysqlUtil()
        {
            if(conn == null)
            {
                //conn = new MySqlConnection("server=" + settings.DfServer + ";user=" + settings.DfUser + ";database=" + settings.DfDB + ";port=" + settings.DfPort + ";password=" + settings.DfPassword + ";charset=" + settings.DfCharset);  //实例化连接
                conn = new MySqlConnection("server= 127.0.0.1 ;user= root; database= hzj ;port= 3306;password= root;charset= utf8");  //实例化连接          
            }
        }
        /// <summary>
        /// 查看数据库能否链接
        /// </summary>
        /// <returns></returns>
        public bool isReady()
        {
            if (conn == null)
            {
                //conn = new MySqlConnection("server=" + settings.DfServer + ";user=" + settings.DfUser + ";database=" + settings.DfDB + ";port=" + settings.DfPort + ";password=" + settings.DfPassword + ";charset=" + settings.DfCharset);  //实例化连接
                conn = new MySqlConnection("server= 127.0.0.1 ;user= root; database= hzj ;port= 3306;password= root;charset= utf8");  //实例化连接          
            }
            conn.Open();
            if(conn.State == System.Data.ConnectionState.Open)
            {
                return true;
            }else
            {
                return false;
            }
        }
        /// <summary>
        /// 建立mysql数据库连接
        /// </summary>
        public void OpenConnect()
        {
            if(isReady() == true)
				this.isOpen = true;
        }
        /// <summary>
        /// 关闭数据库
        /// </summary>
        public void Close()
        {
            conn.Close();
            isOpen = false;
        }
        /// <summary>
        /// 销毁当前对象
        /// </summary>
        public void Dispose()
        {
            conn.Close();
            conn.Dispose();
            GC.SuppressFinalize(this);
            this.isOpen = false;
        }
        /// <summary>
        /// 插入一个SheetInfo，只插入以前表中不存在的数据
        /// </summary>
        /// <param name="sheetInfo"></param>
        /// <returns></returns>
        public int InsertSheetInfo(SheetInfo sheetInfo)
        {
            //判断数据库是否打开
            if (!this.isOpen)
            {
                OpenConnect();
            }
            //事务开始                       
            myTran = conn.BeginTransaction();
            isBegin = true;
            try
            {
                if (null == sheetInfo)
                {
                    return 0;
                }
                //插入数据前，将所有表中ID自增字段设置为从当前表中最大ID开始插入
                TableIDUpdate();
				
				if(sheetInfo.Tag)
				{
					InsertInfoData(sheetInfo);
				}else
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
                isBegin = false;
                hasEntityRemark = false;
                //删除entitytable表中当前sheet的entity
                DeleteCurrentEntity(sheetInfo, InsertIntoExcelTable(sheetInfo)["IdExcel"]);

                throw new Exception("事务操作出错，系统信息：" + e.Message);
            }

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
        /// 将所有表中ID自增字段设置为从当前表中最大ID开始插入数据
        /// </summary>
        private void TableIDUpdate()
        {
            string sql = "alter table ExcelTable auto_increment =1;"
                         + "alter table KeyTable auto_increment =1;"
                         + "alter table EntityTable auto_increment =1;"
                         + "alter table GroupTable auto_increment =1;"
                         + "alter table InfoTable auto_increment =1;"
                         + "alter table DrawDataTable auto_increment =1";
            MySqlCommand mycom = new MySqlCommand(sql, this.conn, myTran);  //建立执行命令语句对象，其中myTran为事务
            try
            {
                mycom.ExecuteNonQuery();
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        /// <summary>
        /// 删除所有表的数据
        /// </summary>
        public void DeleteAllData()
        {
            //关联表外键是级联删除
            string sql = "delete from Drawdatatable; delete from InfoTable; delete from EntityTable; delete from GroupTable; delete from KeyTable; delete from ExcelTable ";
            MySqlCommand mycom = new MySqlCommand(sql, this.conn, myTran);  //建立执行命令语句对象，其中myTran为事务
            try
            {
                mycom.ExecuteNonQuery();
            }
            catch (Exception e)
            {
                throw e;
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
                string sql = GetInsertIntoExcelTableSql(isInfo,signal,tableDesc);
                ExecuteOneSql(sql);
                IdExcel = getIdExcel(signal);

                //InfoTable插入到GroupTable
                if(isInfo)
                    IdGroup = InsertIntoGroupTable(IdExcel);
            }
            else
            {
                string[] tableDesclist = currentTableDesc.Split(';');
                foreach(string t in tableDesclist)
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
            if (isInfo)
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
            using (MySqlCommand mycom = new MySqlCommand(sql, this.conn, myTran))
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
                        sql = sql + ",('" + IdExcel + "','" + IdGroup + "','"+ name + "')";
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
            string sql = "";
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
						if (n % 500 == 0)
                        {
                            if (n > 0)
                                SqlStringList.Add(sql);
                            sql = GetInsertIntoDrawDataTableSql(IdExcel, IdEntity, data.Date, data.MaxValue, data.MidValue, data.MinValue, data.Detail);
                        }
                        else
                            sql = sql + ",('" + IdExcel + "','" + IdEntity + "','" + data.Date + "','" + data.MaxValue + "','" + data.MidValue + "','" + data.MinValue + "','" + data.Detail + "')";

                        n++;						
					}					
				}
				else  //实体表如果没有更新，要检查每个实体表中，时间是否增加
				{
					IdEntity = getIdEntity(IdExcel, drawRow.EntityName);
					List<DateTime> PeriousDateTimes = GetDrawDataTableDateTimes(IdEntity);
					foreach (DrawData data in drawRow.Data)
					{
						if(!PeriousDateTimes.Contains(data.Date))
						{
							if (n % 500 == 0)
                            {
								if (n > 0)
									SqlStringList.Add(sql);
								sql = GetInsertIntoDrawDataTableSql(IdExcel, IdEntity, data.Date, data.MaxValue, data.MidValue, data.MinValue, data.Detail);
							}
							else
								sql = sql + ",('" + IdExcel + "','" + IdEntity + "','" + data.Date + "','" + data.MaxValue + "','" + data.MidValue + "','" + data.MinValue + "','" + data.Detail + "')";

							n++;
						}												
					}
				}
			}
			if (n % 500 != 0)
                SqlStringList.Add(sql);  //不够500values的SQL语言也要添加进去            
            InsertSqlStringList(SqlStringList);
		}
        /// <summary>
        /// 获取KeyTable表的Key和ID
        /// </summary>
        /// <param name="IdExcel"></param>
        /// <returns></returns>
        private Dictionary<string, int> SelectKeyNames(int IdExcel)
        {
            Dictionary<string, int> KeyNames = new Dictionary<string, int>();
            String sql = "select KeyName, ID from KeyTable where Excel_ID = " + IdExcel;

            MySqlCommand mycom = new MySqlCommand(sql, this.conn, myTran);  //建立执行命令语句对象
            MySqlDataReader reader = mycom.ExecuteReader();    //需要关闭
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
            String sql = "select distinct EntityName, ID from EntityTable where Excel_ID = " + IdExcel;

            MySqlCommand mycom = new MySqlCommand();
            mycom.Connection = this.conn;
            mycom.CommandText = sql;
            if(isBegin)
                mycom.Transaction = myTran;
            MySqlDataReader reader = mycom.ExecuteReader();    //需要关闭
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
        private List<DateTime> GetDrawDataTableDateTimes(int IdEntity)
        {
            List<DateTime> PeriousDateTimes = new List<DateTime>();
            String sql = "select Date from DrawDataTable where Entity_ID = " + IdEntity;

            MySqlCommand mycom = new MySqlCommand(sql, this.conn, myTran);  //建立执行命令语句对象
            MySqlDataReader reader = mycom.ExecuteReader();    //需要关闭
            try
            {
                while (reader.Read())
                {
                    if (reader.HasRows)
                    {
                        PeriousDateTimes.Add(reader.GetDateTime(0));
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
            MySqlCommand mycom = new MySqlCommand(sql, this.conn, myTran);  //建立执行命令语句对象
            MySqlDataReader reader = mycom.ExecuteReader();    //需要关闭
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
        /// <summary>
        /// 执行批量插入
        /// </summary>
        /// <param name="SqlStringList"></param>
        private void InsertSqlStringList(List<string> SqlStringList)
        {
            MySqlCommand command = new MySqlCommand();
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

            MySqlCommand mycom = new MySqlCommand();
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

        //表ExcelTable的InsertSql
        private String GetInsertIntoExcelTableSql(bool isInfo, string signal, string tableDesc)
        {
            String sql = String.Format("insert into ExcelTable (TableDesc,ExcelSignal,IsInfo) values ('{0}', '{1}', {2})", tableDesc, signal, isInfo);
            return sql;
        }
        //表KeyTable的InsertSql
        private String GetInsertIntoKeyTableSql(int idExcel,int idGroup, string keyName)
        {
            String sql = String.Format("insert into KeyTable ( Excel_ID, Group_ID, KeyName ) values ('{0}','{1}','{2}')", idExcel, idGroup, keyName);
            return sql;
        }
        //表EntityTable的InsertSql
        private String GetInsertIntoEntityTableSql(int idExcel, String entityName, String entityRemark = "")
        {
            String sql = String.Format("insert into EntityTable ( Excel_ID, EntityName, Remark ) values ('{0}','{1}','{2}')", idExcel, entityName, entityRemark);
            return sql;
        }
        //表InfoTable的InsertSql
        private String GetInsertIntoInfoTableSql(int idKey, int idEntity, String value)
        {
            String sql = String.Format("insert into InfoTable ( Key_ID, Entity_ID, Value ) values ('{0}','{1}','{2}')", idKey, idEntity, value);
            return sql;
        }

        //表DrawDataTable的InsertSql
        private String GetInsertIntoDrawDataTableSql(int idExcel, int idEntity, DateTime date, double maxValue, double midValue, double minValue, string detail)
        {
            String sql = String.Format("insert into DrawDataTable ( Excel_ID, Entity_ID, Date, EntityMaxValue, EntityMidValue, EntityMinValue, Detail ) values ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", idExcel, idEntity, date, maxValue, midValue, minValue, detail);
            return sql;
        }

        //表GroupTable的InsertSql
        private String GetInsertIntoGroupTableSql(int idExcel, string groupName = "未分组", string remark= "")
        {
            String sql = String.Format("insert into GroupTable ( Excel_ID, GroupName, Remark ) values ('{0}','{1}','{2}')", idExcel, groupName, remark);
            return sql;
        }

        /// <summary>
        /// 根据signal获取Excel_ID
        /// </summary>
        /// <param name="signal"></param>
        /// <returns></returns>
        private int getIdExcel(string signal)
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
        private int getDefaultGroupId(int idGroup)
        {
            string sql = String.Format("select b.ID from GroupTable a, GroupTable b where a.Excel_ID = b.Excel_ID and a.ID = '{0}' and b.GroupName = '未分组' ", idGroup);
            int id = getID(sql);
            return id;
        }
        /// <summary>
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
        private int getIdEntity(int idExcel, String entityName)
        {
            string sql = String.Format("select ID from EntityTable where Excel_ID = {0} and EntityName = '{1}'", idExcel, entityName);
            int id = getID(sql);
            return id;
        }
        private int getID(String sql)
        {
            MySqlCommand mycom = new MySqlCommand(sql, this.conn, this.myTran);  //建立执行命令语句对象
            int id = 0;
            try
            {
                id = Convert.ToInt32(mycom.ExecuteScalar());   //查询返回一个值的时候，用ExecuteScalar()更节约资源，快捷
            }
            catch { throw; }
            return id;
        }

        //*****注意，sql语句多行写的话，开始要注意留空格*******

        /// <summary>
        /// 新建一个Group
        ///在选定的某种Excel的基础上新建一个Group
        ///传入Signal标志
        /// </summary>
        public bool AddKeyGroup(string Signal, string GroupName)
        {
            bool flag = false;
            int IdExcel = getIdExcel(Signal);
            string sql = GetInsertIntoGroupTableSql(IdExcel, GroupName);
            //判断数据库是否打开
            if (!this.isOpen)
            {
                OpenConnect();
            }
            MySqlCommand mycom;
            //if (isBegin)
            //    mycom = new MySqlCommand(sql, this.conn, this.myTran);  //建立执行命令语句对象
            //else
                mycom = new MySqlCommand(sql, this.conn);  //建立执行命令语句对象

            try
            {
                if (mycom.ExecuteNonQuery() > 0)
                    flag = true;
                return flag;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 删除一个Group
        /// </summary>
        public bool DeleteKeyGroup(int Group_ID)
        {
            bool flag = false;
            int defaultGroup = getDefaultGroupId(Group_ID);
            string sql = String.Format("delete from GroupTable where ID = '{0}'; update KeyTable set Group_ID = '{1}' where Group_ID = '{2}'", Group_ID , defaultGroup, Group_ID);
            //判断数据库是否打开
            if (!this.isOpen)
            {
                OpenConnect();
            }
            MySqlCommand mycom;
            mycom = new MySqlCommand(sql, this.conn);  //建立执行命令语句对象

            try
            {
                if (mycom.ExecuteNonQuery() > 0)
                    flag = true;
                return flag;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 修改某个Group的GroupName
        /// </summary>
        public bool ModifyKeyGroup(int Group_ID, string GroupName)
        {
            bool flag = false;
            string sql = String.Format("Update GroupTable set GroupName = '{0}' where ID = '{1}'", GroupName, Group_ID);
            //判断数据库是否打开
            if (!this.isOpen)
            {
                OpenConnect();
            }
            MySqlCommand mycom;
            mycom = new MySqlCommand(sql, this.conn);  //建立执行命令语句对象

            try
            {
                if (mycom.ExecuteNonQuery() > 0)
                    flag = true;
                return flag;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 添加一些属性名到某一Group
        /// </summary>
        public bool AddKeysToGroup(int Group_ID, List<int> Key_Ids)
        {
            bool flag = false;
            int n = 0;
            string sql = "Update KeyTable set Group_ID = " + Group_ID + " where ID in ( " ;

            foreach (int idKey in Key_Ids)
            {
                if (n != 0)
                    sql += ", ";
                sql +=  idKey ;
                n++;
            }
            if (n == 0)
                sql += "'')";
            else
                sql += ")";           
            //判断数据库是否打开
            if (!this.isOpen)
            {
                OpenConnect();
            }
            MySqlCommand mycom;
            mycom = new MySqlCommand(sql, this.conn);  //建立执行命令语句对象

            try
            {
                if (mycom.ExecuteNonQuery() > 0)
                    flag = true;
                return flag;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 列举所有的Info Excel等待分组，所有基础数据excel表
        ///返回的Dictionary中
        ///key:desc描述
        ///string:signal简写（全局唯一）
        /// </summary>
        public Dictionary<string, string> ListExcelToGroup()
        {
            Dictionary<string, string> list = new Dictionary<string, string>();
            string sql = String.Format("select distinct TableDesc, ExcelSignal from ExcelTable where isInfo = true ");
            //判断数据库是否打开
            if (!isOpen)
            {
                OpenConnect();
            }
            MySqlCommand mycom = new MySqlCommand(sql, this.conn);  //建立执行命令语句对象
            MySqlDataReader reader = mycom.ExecuteReader();    //需要关闭
            try
            {
                while (reader.Read())
                {
                    if (reader.HasRows)
                    {
                        list.Add(reader.GetString(0), reader.GetString(1));
                    }
                }
                return list;
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
        /// 修改Entity的Remark********************************************
        /// </summary>
        public bool ModifyEntityRemark(string entityName, string remark)
        {
            bool flag = false;
            //根据entityName获取idExcel,idEntity
            int idExcel = 0;
            int idEntity = 0;
            string presql = String.Format("select ID, Excel_ID from EntityTable where EntityName = '{0}' ", entityName);
            //判断数据库是否打开
            if (!this.isOpen)
            {
                OpenConnect();
            }
            MySqlCommand mycom = new MySqlCommand();
            mycom.Connection = this.conn;
            mycom.CommandText = presql;
            MySqlDataReader reader = mycom.ExecuteReader();    //需要关闭
            try
            {
                while (reader.Read())
                {
                    if (reader.HasRows)
                    {
                        idEntity = reader.IsDBNull(0) ? 0 : Convert.ToInt32(reader.GetValue(0));
                        idExcel = reader.IsDBNull(1) ? 0 : Convert.ToInt32(reader.GetValue(1));
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
            if(idEntity == 0)
            {
                MessageBox.Show("实体名不存在，建议核查！");
                return flag;
            }

            //获取idKey       
            int idKey = getIdKey(idExcel,"备注");
            if(idKey == 0)
            {
                MessageBox.Show("实体名不属于基础数据，建议核查！");
                return flag;
            }
            //modify Entity remark
            string sql = String.Format("Replace into InfoTable (Key_ID,Entity_ID,Value) values ('{0}','{1}','{2}') ", idKey, idEntity, remark);
            mycom.CommandText = sql;
            try
            {
                if (mycom.ExecuteNonQuery() > 0)
                    flag = true;
                return flag;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 修改阈值Total_hold和Diff_hold
        ///修改某一种Excel表的阈值，这里的Excel表必须是测量数据表
        /// </summary>
        public bool ModifyThreshold(string signal, double Total_hold, double Diff_hold)
        {
            bool flag = false;
            string sql = String.Format("Update ExcelTable set Total_hold = '{0}', Diff_hold = '{1}' where ExcelSignal = '{2}'", Total_hold, Diff_hold, signal);
            //判断数据库是否打开
            if (!this.isOpen)
            {
                OpenConnect();
            }
            MySqlCommand mycom;
            mycom = new MySqlCommand(sql, this.conn);  //建立执行命令语句对象

            try
            {
                if (mycom.ExecuteNonQuery() > 0)
                    flag = true;
                return flag;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 复制MySQL中的数据到Sqlite中
        /// </summary>
        public MySqlDataReader LoadTableData(string tablename)
        {
            string sql = "select * from " + tablename;
            //判断数据库是否打开
            if (!this.isOpen)
            {
                OpenConnect();
            }
            MySqlCommand mycom = new MySqlCommand(sql, this.conn);  //建立执行命令语句对象
            MySqlDataReader reader = mycom.ExecuteReader();    //需要关闭

            return reader;
        }

        /// <summary>
        /// 查询当前Excel表信息
        /// 可用于查询当前阈值和分组时使用
        /// </summary>
        /// <returns>返回一个ExcelTable</returns>
        public List<ExcelTable> ListExcelsMessage()
        {
            List<ExcelTable> excelTable = new List<ExcelTable>();
            string sql = String.Format("select distinct ExcelSignal,TableDesc,Total_hold,Diff_hold from ExcelTable order by ExcelSignal ");
            //判断数据库是否打开
            if (!isOpen)
            {
                OpenConnect();
            }
            MySqlCommand mycom = new MySqlCommand(sql, this.conn);  //建立执行命令语句对象
            MySqlDataReader reader = mycom.ExecuteReader();    //需要关闭
            try
            {
                while (reader.Read())
                {
                    ExcelTable excel = new ExcelTable();
                    if (reader.HasRows)
                    {
                        excel.Signal = reader.GetString(0);
                        excel.TableDesc = reader.GetString(1);
                        excel.Total_hold = reader.IsDBNull(2) ? 0 : Convert.ToSingle(reader.GetValue(2));
                        excel.Diff_hold = reader.IsDBNull(3) ? 0 : Convert.ToSingle(reader.GetValue(3));
                        excelTable.Add(excel);
                    }
                }
                return excelTable;
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
        /// 查询一种表的分组信息
        /// </summary>
        /// <param name="signal">表的简称</param>
        /// <returns></returns>
        public List<Group> loadGroupForAExcel(string signal)
        {
            List<Group> groupList = new List<Group>();
            string sql = String.Format("select gt.ID, GroupName from GroupTable gt, ExcelTable et where gt.Excel_ID = et.ID and et.ExcelSignal = '{0}' order by gt.ID ", signal);
            //判断数据库是否打开
            if (!isOpen)
            {
                OpenConnect();
            }
            MySqlCommand mycom = new MySqlCommand(sql, this.conn);  //建立执行命令语句对象
            MySqlDataReader reader = mycom.ExecuteReader();    //需要关闭
            try
            {
                while (reader.Read())
                {
                    Group g = new Group();
                    if (reader.HasRows)
                    {
                        g.Id = reader.GetInt32(0);
                        g.GroupName = reader.GetString(1);
                        groupList.Add(g);
                    }
                }
                return groupList;
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
        /// 通过Group_Id查询所有与之相关的KeyName
        /// </summary>
        /// <param name="group_id"></param>
        /// <returns></returns>
        public List<KeyTableRow> loadKeyNameForAGroup(int group_id)
        {
            List<KeyTableRow> keyGroup = new List<KeyTableRow>();
            string sql = String.Format("select ID, KeyName from KeyTable where Group_ID = '{0}' ", group_id);
            //判断数据库是否打开
            if (!isOpen)
            {
                OpenConnect();
            }
            MySqlCommand mycom = new MySqlCommand(sql, this.conn);  //建立执行命令语句对象
            MySqlDataReader reader = mycom.ExecuteReader();    //需要关闭
            try
            {
                while (reader.Read())
                {
                    KeyTableRow keyRow = new KeyTableRow();
                    if (reader.HasRows)
                    {
                        keyRow.Id =  reader.GetInt32(0);
                        keyRow.KeyName = reader.GetString(1);
                        keyGroup.Add(keyRow);
                    }
                }
                return keyGroup;
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
        /// 通过Signal来查询与之相关的所有的KeyName
        /// </summary>
        /// <param name="signal"></param>
        /// <returns></returns>
        public List<KeyTableRow> loadKeyNameForAExcel(string signal)
        {
            List<KeyTableRow> keyGroup = new List<KeyTableRow>();
            int idExcel = getIdExcel(signal);
            string sql = String.Format("select ID, KeyName from KeyTable where Excel_ID = '{0}' ", idExcel);
            //判断数据库是否打开
            if (!isOpen)
            {
                OpenConnect();
            }
            MySqlCommand mycom = new MySqlCommand(sql, this.conn);  //建立执行命令语句对象
            MySqlDataReader reader = mycom.ExecuteReader();    //需要关闭
            try
            {
                while (reader.Read())
                {
                    KeyTableRow keyRow = new KeyTableRow();
                    if (reader.HasRows)
                    {
                        keyRow.Id = reader.GetInt32(0);
                        keyRow.KeyName = reader.GetString(1);
                        keyGroup.Add(keyRow);
                    }
                }
                return keyGroup;
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
}
