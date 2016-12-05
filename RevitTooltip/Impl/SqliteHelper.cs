using Revit.Addin.RevitTooltip.Intface;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Revit.Addin.RevitTooltip.Dto;

namespace Revit.Addin.RevitTooltip.Impl
{
    class SqliteHelper : ISQLiteHelper
    {
        public SqliteHelper(RevitTooltip settings) {
            throw new NotImplementedException();
        }
        public int InsertSheetInfo(SheetInfo sheetInfo)
        {
            throw new NotImplementedException();
        }

        public bool LoadDataToSqlite()
        {
            throw new NotImplementedException();
        }

        public List<string> SelectAllEntities(string ExcelSignal)
        {
            throw new NotImplementedException();
        }

        public List<string> SelectDiffThresholdEntity(string ExcelSignal, string DiffThreshold)
        {
            throw new NotImplementedException();
        }

        public DrawData SelectDrawData(string EntityName, DateTime Date)
        {
            throw new NotImplementedException();
        }

        public DrawEntityData SelectDrawEntityData(string EntityName, DateTime? StartTime, DateTime? EndDate)
        {
            throw new NotImplementedException();
        }

        public InfoEntityData SelectInfoData(string EntityName)
        {
            throw new NotImplementedException();
        }

        public List<string> SelectTotalThresholdEntity(string ExcelSignal, string TotalThreshold)
        {
            throw new NotImplementedException();
        }
    }
}
