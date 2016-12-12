using Revit.Addin.RevitTooltip.Dto;
using System;
using System.Collections.Generic;

namespace Revit.Addin.RevitTooltip.Intface
{
    public interface ISQLiteHelper
    {
        /// <summary>
        /// 插入SheetInfo
        /// </summary>
        void InsertSheetInfo(SheetInfo sheetInfo);

        /// <summary>
        /// 查询InfoTable
        ///返回的数据是已分组数据
        /// </summary>
        InfoEntityData SelectInfoData(string EntityName);

        /// <summary>
        /// 查询DrawDataTable
        ///查询Entity时间序列数据
        ///根据传入的起始时间查询
        /// </summary>
        DrawEntityData SelectDrawEntityData(string EntityName, DateTime? StartTime, DateTime? EndDate);


        /// <summary>
        /// 查询DrawDataTable
        ///查询Entity某日期的数据
        /// </summary>
        DrawData SelectDrawData(string EntityName, DateTime Date);


        /// <summary>
        /// 查询Total_hold异常点
        ///返回该类型的所有异常点
        /// </summary>
        List<string> SelectTotalThresholdEntity(string ExcelSignal, float TotalThreshold);

        /// <summary>
        /// 查询Diff_hold异常点
        ///返回该类型的所有异常点
        /// </summary>
        List<string> SelectDiffThresholdEntity(string ExcelSignal, float DiffThreshold);

        /// <summary>
        /// 通过传入的Signal，查询与之对应的所有的测点
        ///传入的Signal应该是测量数据的signal
        /// </summary>
        List<CEntityName> SelectAllEntities(string ExcelSignal);

        /// <summary>
        /// 复制MySQL中的数据到Sqlite中
        /// </summary>
        bool LoadDataToSqlite();
        /// <summary>
        /// 查询所有的测量数据类型
        /// </summary>
        /// <returns></returns>
        List<ExcelTable> SelectDrawTypes();
        /// <summary>
        ///修改备注
        /// </summary>
        /// <param name="EntityName">实体名</param>
        /// <param name="Remark">备注</param>
        /// <returns>是否成功</returns>
        bool ModifyEntityRemark(string EntityName, string Remark);
    }
}