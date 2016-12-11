using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Revit.Addin.RevitTooltip.Dto;
using Revit.Addin.RevitTooltip.Impl;
using Revit.Addin.RevitTooltip.Intface;
using MySql.Data;
using MySql.Data.MySqlClient;

namespace ConsoleApplication1
{
    class SqliteTest
    {
        static void Main(string[] args)
        {
            //连接打开数据库
            SQLiteHelper sqlite = SQLiteHelper.CreateInstance();
            sqlite.UpdateDB();

            //IExcelReader excel = new ExcelReader();
            //Revit.Addin.RevitTooltip.Dto.SheetInfo sheetInfo = excel.loadExcelData("C:\\workspace\\慧之建项目\\project\\RevitVersionChange-master\\数据文档\\基础构件-JC-I.xls");
            //Revit.Addin.RevitTooltip.Dto.SheetInfo sheetInfo1 = excel.loadExcelData("C:\\workspace\\慧之建项目\\project\\RevitVersionChange-master\\数据文档\\测斜汇总-CX-CX.xls");

            //sqlite.InsertSheetInfo(sheetInfo);
            //System.Console.WriteLine("一次插入完成");

            //sqlite.InsertSheetInfo(sheetInfo1);
            //System.Console.WriteLine("二次插入完成");

            //查询绘图数据
            string entityName = "CX1";
            DateTime startDate = Convert.ToDateTime("2015-07-01 00:00:00");
            DateTime Date2 = Convert.ToDateTime("2015-07-13 00:00:00");

            DrawEntityData drawEntityData = new DrawEntityData();
            drawEntityData = sqlite.SelectDrawEntityData(entityName, startDate, Date2);
            OutputDrawEntityData(drawEntityData);

            DrawData drawData = new DrawData();
            drawData = sqlite.SelectDrawData(entityName, Convert.ToDateTime("2015-07-03 00:00:00"));
            Console.WriteLine(entityName);
            OutputDrawData(drawData);
            Console.ReadKey();

            //测异常点
            string excelSignal = "CX";
            double totalThreshold = 6;
            double DiffThreshold = 1.5;

            List<string> t_Entities = new List<string>();
            t_Entities = sqlite.SelectTotalThresholdEntity(excelSignal, totalThreshold);
            Console.WriteLine("total异常点****************");
            OutputEntities(t_Entities);

            List<string> d_Entities = new List<string>();
            d_Entities = sqlite.SelectDiffThresholdEntity(excelSignal, DiffThreshold);
            Console.WriteLine("diff异常点****************");
            OutputEntities(d_Entities);

            List<string> Entities = new List<string>();
            Entities = sqlite.SelectAllEntities(excelSignal);
            Console.WriteLine("所有Entity****************");
            OutputEntities(Entities);

            //System.Console.WriteLine("查询InfoTable");
            //string entityName = "1-5底板";
            //InfoEntityData infoData = new InfoEntityData();
            //infoData = sqlite.SelectInfoData(entityName);
            //ouputInfoEntityData(infoData);



            Console.ReadKey();
        }

        public static void OutputDrawEntityData(DrawEntityData drawEntityData)
        {
            if (drawEntityData.EntityName != null)
            {
                Console.WriteLine(drawEntityData.EntityName);
                Console.WriteLine("Date  MaxValue   MidValue  MinValue");
                foreach (DrawData d in drawEntityData.Data)
                {
                    OutputDrawData(d);
                    //Console.WriteLine(d.Date + " " + d.MaxValue + " " + d.MidValue + " " + d.MinValue);
                }
            }
        }

        public static void OutputDrawData(DrawData drawData)
        {
            Console.WriteLine(drawData.Date + " " + drawData.MaxValue + " " + drawData.MidValue + " " + drawData.MinValue);
        }

        public static void OutputEntities(List<string> Entities)
        {
            foreach (string s in Entities)
            {
                Console.WriteLine(s);
            }
        }

        public static void ouputInfoEntityData(InfoEntityData infoData)
        {
            Console.WriteLine(infoData.EntityName);
            Console.WriteLine("*********infoData**********");
            foreach (string s in infoData.Data.Keys)
            {
                Console.WriteLine(s + " " + infoData.Data[s]);
            }
            Console.WriteLine("*********group message**********");
            foreach (string group in infoData.GroupMsg.Keys)
            {
                Console.WriteLine(group);
                foreach (string m in infoData.GroupMsg[group])
                {
                    Console.WriteLine(" " + m);
                }
            }
        }
    }
}
