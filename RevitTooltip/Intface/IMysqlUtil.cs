﻿using System.Collections.Generic;
using Revit.Addin.RevitTooltip.Dto;
using MySql.Data.MySqlClient;

namespace Revit.Addin.RevitTooltip.Intface
{
    public interface IMysqlUtil
    {
        /// <summary>
        /// 用测试语句测试当前Mysql的可用状态
        /// </summary>
        bool IsReady { get; }
        
        /// <summary>
        /// 插入SheetInfo
		///返回成功完成了多少个Entity相关的数据
        /// </summary>
        void InsertSheetInfo(SheetInfo sheetInfo);

        /// <summary>
        /// 新建一个Group
        ///在选定的某种Excel的基础上新建一个Group
        ///传入Signal标志
        /// </summary>
        Group AddNewGroup(string Signal, string GroupName);

        /// <summary>
        /// 删除一个Group
        /// </summary>
        bool DeleteKeyGroup(int Group_ID);

        /// <summary>
        /// 修改某个Group的GroupName
        /// </summary>
        bool ModifyKeyGroup(int Group_ID, string GroupName);

        /// <summary>
        /// 添加一些属性名到某一Group
        /// </summary>
        bool AddKeysToGroup(int? Group_ID, List<int> Key_Ids);

        /// <summary>
        /// 列举所有的Info Excel等待分组，这里不包括测量数据的excel
        ///返回的Dictionary中
        ///string2:signal简写（全局唯一）
        ///string1:desc描述
        /// </summary>
        Dictionary<string, string> ListExcelToGroup();

        /// <summary>
        /// 修改Entity的Remark
        /// </summary>
        bool ModifyEntityRemark(string EntityName, string Remark);
        /// <summary>
        /// 查询当前Excel表信息
        /// 可用于查询当前阈值和分组时使用
        /// </summary>
        /// <returns>返回一个ExcelTable</returns>
        List<ExcelTable> ListExcelsMessage(bool isInfo);
        /// <summary>
        /// 修改阈值Total_hold和Diff_hold
        ///修改某一种Excel表的阈值，这里的Excel表必须是测量数据表
        /// </summary>
        bool ModifyThreshold(string signal, float Total_hold, float Diff_hold);

       
        /// <summary>
        /// 查询一种表的分组信息
        /// </summary>
        /// <param name="signal">表的简称</param>
        /// <returns></returns>
        List<Group> loadGroupForAExcel(string signal);
        /// <summary>
        /// 通过Group_Id查询所有与之相关的KeyName
        /// </summary>
        /// <param name="group_id"></param>
        /// <returns></returns>
        //List<CKeyName> loadKeyNameForAGroup(int group_id);
        ///// <summary>
        ///// 通过Signal来查询与之相关的所有的KeyName
        ///// </summary>
        ///// <param name="signal"></param>
        ///// <returns></returns>
        List<CKeyName> loadKeyNameForExcelAndGroup(string signal,int Group_id);
        /// <summary>
        /// 修改Key的分组
        /// </summary>
        /// <param name="Group_id">为Null时，表示取消</param>
        /// <param name="Key_id">KeyName的ID</param>
        void updateKeyGroup(int? Group_id,int Key_id);

    }
}