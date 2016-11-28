using System;
using System.Collections.Generic;
using System.Linq;
using Revit.Addin.RevitTooltip.Dto;
using Revit.Addin.RevitTooltip.Impl;
using MySql.Data.MySqlClient;

namespace Revit.Addin.RevitTooltip.Intface
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
        private int hasEntityRemark = 0;

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
                hasEntityRemark = 0;
                return 1;
            }
            catch (Exception e)
            {
                myTran.Rollback();    // 事务回滚
                isBegin = false;
                hasEntityRemark = 0;
                //删除entitytable表中当前sheet的entity
                DeleteCurrentEntity(sheetInfo, InsertIntoExcelTable(sheetInfo));

                throw new Exception("事务操作出错，系统信息：" + e.Message);
            }

        }
        /// <summary>
        /// 插入基础数据表
        /// </summary>
        /// <param name="sheetInfo"></param>
        private void InsertInfoData(SheetInfo sheetInfo)
        {
            //插入到ExcelTable表,并返回ID
			int IdExcel = InsertIntoExcelTable(sheetInfo);

            //插入到KeyTable表，并返回当前表的键值对<属性名，ID>
			Dictionary<string, int> CurrentKeys = InsertIntoKeyTable(sheetInfo.KeyNames, IdExcel);

            //插入到EntityTable表，并返回新插入的实体数据的键值对<实体名,ID>
            Dictionary<string, int> UpdateEntities = InsertIntoEntityTable(sheetInfo.EntityNames, IdExcel);

            //插入到InfoTable表
            InsertIntoInfoTable(sheetInfo, CurrentKeys, UpdateEntities);
        }
        /// <summary>
        /// 插入绘图数据表
        /// </summary>
        /// <param name="sheetInfo"></param>
        private void InsertDrawData(SheetInfo sheetInfo)
        {
			int IdExcel = InsertIntoExcelTable(sheetInfo);

            Dictionary<string, int> UpdateEntities = InsertIntoEntityTable(sheetInfo.EntityNames, IdExcel);

            InsertIntoDrawDataTable(sheetInfo, UpdateEntities, IdExcel);
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
            string sql = "delete from exceltable ";
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
        /// 插入到ExcelTable表
        /// 返回对应ID
        /// </summary>
        /// <param name="sheetInfo"></param>
        /// <returns></returns>
        private int InsertIntoExcelTable(SheetInfo sheetInfo)
        {
			bool isInfo = sheetInfo.Tag;
			string signal = sheetInfo.ExcelTableData.Signal;
			string tableDesc = sheetInfo.ExcelTableData.TableDesc;
            int id = getIdExcel(signal);
            if (id == 0)
            {
                String sql = GetInsertIntoExcelTableSql(isInfo,signal,tableDesc);
                ExecuteOneSql(sql);
                id = getIdExcel(signal);
            }
            //update数据库中的TableDesc******************待实现************************
            return id;
        }
        /// <summary>
        /// 插入属性到KeyTable表
        /// 返回当前表的属性及对应ID，<KeyName,ID>
        /// </summary>
        /// <param name="KeyNames"></param>
        /// <param name="IdExcel"></param>
        /// <returns></returns>
        private Dictionary<string, int> InsertIntoKeyTable(List<string> KeyNames, int IdExcel)
        {
            Dictionary<string, int> PeriousKeys = SelectKeyNames(IdExcel);
            String sql = "";
            int n = 0;
            foreach (string name in KeyNames)
            {
                if (name.Equals("备注"))
                {
                    hasEntityRemark = 1;
                    continue;
                }

                //如果没有在原表中匹配到，则添加到插入语句当中
                if (!PeriousKeys.Keys.Contains(name))
                {
                    if (n == 0)
                        sql = GetInsertIntoKeyTableSql(IdExcel, name);
                    else
                        sql = sql + ",('" + IdExcel + "','" + name + "')";
                    n++;
                }
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
            String sql = "";
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

            MySqlCommand mycom = new MySqlCommand(sql, this.conn, myTran);  //建立执行命令语句对象
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
        private void ExecuteOneSql(String sql)
        {
            if (sql.Equals(""))
                return;
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

        //表ExcelTable的InsertSql
        private String GetInsertIntoExcelTableSql(bool isInfo, string signal, string tableDesc)
        {
            String sql = String.Format("insert into ExcelTable (TableDesc,ExcelSignal,IsInfo) values ('{0}', '{1}', {2})", tableDesc, signal, isInfo);
            return sql;
        }
        //表KeyTable的InsertSql
        private String GetInsertIntoKeyTableSql(int idExcel, String keyName)
        {
            String sql = String.Format("insert into KeyTable ( Excel_ID, KeyName ) values ('{0}','{1}')", idExcel, keyName);
            return sql;
        }
        //表EntityTable的InsertSql
        private String GetInsertIntoEntityTableSql(int idExcel, String entityName, String entityColumn = "测点编号", String entityRemark = "")
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
        private String GetInsertIntoGroupTableSql(int idExcel, string groupName, String remark= "")
        {
            String sql = String.Format("insert into GroupTable ( Excel_ID, GroupName, Remark ) values ('{0}','{1}','{2}')", idExcel, groupName, remark);
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
        /// 后期可能会删除*********************************************
        /// </summary>
        public MySqlDataReader GetCXTable()
        {
            String sql = "Select e.Entity,c.Date,c.Value from CXView c, EntityTable e where c.ID_Entity = e.ID "
                       + " group by c.ID_Entity,c.Date order by c.ID_Entity,c.Date ";

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
            if (isBegin)
                mycom = new MySqlCommand(sql, this.conn, this.myTran);  //建立执行命令语句对象
            else
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
            string sql = String.Format("delete from GroupTable where ID = '{0}'; update KeyTable set Group_ID = null where Group_ID = '{1}'", Group_ID , Group_ID);

            //判断数据库是否打开
            if (!this.isOpen)
            {
                OpenConnect();
            }
            MySqlCommand mycom;
            if (isBegin)
                mycom = new MySqlCommand(sql, this.conn, this.myTran);  //建立执行命令语句对象
            else
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
            if (isBegin)
                mycom = new MySqlCommand(sql, this.conn, this.myTran);  //建立执行命令语句对象
            else
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
            if (isBegin)
                mycom = new MySqlCommand(sql, this.conn, this.myTran);  //建立执行命令语句对象
            else
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
        /// 列举所有的Info Excel等待分组，这里不包括测量数据的excel
        ///返回的Dictionary中
        ///string2:signal简写（全局唯一）
        ///string1:desc描述
        /// </summary>
        public Dictionary<string, string> ListExcelToGroup()
        {
            Dictionary<string, string> list = new Dictionary<string, string>();
            string sql = String.Format("select ExcelSignal, TableDesc from ExcelTable where ID NOT IN( select distinct Excel_ID from GroupTable )" );
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
        /// 修改Entity的Remark
        /// </summary>
        public bool ModifyEntityRemark(string entityName, string remark)
        {
            bool flag = false;
            string sql = String.Format("Update EntityTable set Remark = '{0}' where EntityName = '{1}'", remark, entityName);
            //判断数据库是否打开
            if (!this.isOpen)
            {
                OpenConnect();
            }
            MySqlCommand mycom;
            if (isBegin)
                mycom = new MySqlCommand(sql, this.conn, this.myTran);  //建立执行命令语句对象
            else
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
        /// 修改阈值Total_hold和Diff_hold
        ///修改某一种Excel表的阈值，这里的Excel表必须是测量数据表
        /// </summary>
        public bool ModifyThreshold(string signal, double Total_hold, double Diff_hold)
        {
            bool flag = false;
            string sql = String.Format("Update ExcelTabel set Total_hold = '{0}', Diff_hold = '{1}' where ExcelSignal = '{2}'", Total_hold, Diff_hold, signal);
            //判断数据库是否打开
            if (!this.isOpen)
            {
                OpenConnect();
            }
            MySqlCommand mycom;
            if (isBegin)
                mycom = new MySqlCommand(sql, this.conn, this.myTran);  //建立执行命令语句对象
            else
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
    }
}
