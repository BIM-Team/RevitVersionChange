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
    class MysqlTest
    {
        static void Main(string[] args)
        {
            //连接打开数据库
            MysqlUtil mysql = MysqlUtil.CreateInstance();
            //打开连接
            mysql.OpenConnect();

            IExcelReader excel = new ExcelReader();
            Revit.Addin.RevitTooltip.Dto.SheetInfo sheetInfo = excel.loadExcelData("C:\\workspace\\慧之建项目\\project\\RevitVersionChange-master\\数据文档\\基础构件-JC-I.xls");

            Revit.Addin.RevitTooltip.Dto.SheetInfo sheetInfo1 = excel.loadExcelData("C:\\workspace\\慧之建项目\\project\\RevitVersionChange-master\\数据文档\\测斜汇总-CX-CX.xls");


            //System.Console.WriteLine("###########################\n索引：" + sheetInfo.SheetIndex);

            System.Console.WriteLine("####################将数据插入到数据库:");

            mysql.InsertSheetInfo(sheetInfo);
            System.Console.WriteLine("一次插入完成");

            //mysql.DeleteAllData();
            mysql.InsertSheetInfo(sheetInfo1);
            System.Console.WriteLine("二次插入完成");

            Console.ReadKey();

            if (mysql.AddKeyGroup("JX", "Group1"))
                System.Console.WriteLine("成功添加Group1");

            Console.ReadKey();

            if (mysql.ModifyKeyGroup(1, "Group1Modify"))
                System.Console.WriteLine("成功将Group1名称改为Group1Modify");
            Console.ReadKey();

            List<int> Key_Ids = new List<int>();
            Key_Ids.Add(1);
            Key_Ids.Add(2);
            Key_Ids.Add(3);

            if (mysql.AddKeysToGroup(1, Key_Ids))
                System.Console.WriteLine("添加Key到Group1Modify");
            Console.ReadKey();

            Dictionary<string, string> ListExcel = new Dictionary<string, string>();
            ListExcel = mysql.ListExcelToGroup();
            foreach (string s in ListExcel.Keys)
            {
                System.Console.WriteLine(s + " " + ListExcel[s]);
            }
            Console.ReadKey();

            mysql.DeleteKeyGroup(1);
            ListExcel = mysql.ListExcelToGroup();
            foreach (string s in ListExcel.Keys)
            {
                System.Console.WriteLine(s + " " + ListExcel[s]);
            }
            Console.ReadKey();


        }
    }
}
