//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using Revit.Addin.RevitTooltip.Dto;
//using Revit.Addin.RevitTooltip.Impl;
//using Revit.Addin.RevitTooltip.Intface;
//using MySql.Data;
//using MySql.Data.MySqlClient;

//namespace ConsoleApplication1
//{
//    class MysqlTest
//    {
//        static void Main(string[] args)
//        {
//            //连接打开数据库
//            MysqlUtil mysql = MysqlUtil.CreateInstance();
//            //打开连接
//            mysql.OpenConnect();
//            //            System.Console.WriteLine("#######删除*********");
//            //            mysql.DeleteAllData();
//            //            Console.ReadKey();
//            //            IExcelReader excel = new ExcelReader();
//            //            Revit.Addin.RevitTooltip.Dto.SheetInfo sheetInfo = excel.loadExcelData("C:\\workspace\\慧之建项目\\project\\RevitVersionChange-master\\数据文档\\基础构件-JC-I.xls");

//            //            Revit.Addin.RevitTooltip.Dto.SheetInfo sheetInfo1 = excel.loadExcelData("C:\\workspace\\慧之建项目\\project\\RevitVersionChange-master\\数据文档\\测斜汇总-CX-CX.xls");


//            //            //System.Console.WriteLine("###########################\n索引：" + sheetInfo.SheetIndex);

//            //            System.Console.WriteLine("####################将数据插入到数据库:");

//            //            mysql.InsertSheetInfo(sheetInfo);
//            //            System.Console.WriteLine("一次插入完成");

//            //            //mysql.DeleteAllData();
//            //            mysql.InsertSheetInfo(sheetInfo1);
//            //            System.Console.WriteLine("二次插入完成");

//            //            Console.ReadKey();

//            //            System.Console.WriteLine("列举所有的Info Excel等待分组,所有基础数据excel表");
//            //            Dictionary<string, string> ListExcel = new Dictionary<string, string>();
//            //            ListExcel = mysql.ListExcelToGroup();
//            //            foreach (string s in ListExcel.Keys)
//            //            {
//            //                System.Console.WriteLine(s + " " + ListExcel[s]);
//            //            }
//            //            Console.ReadKey();

//            //            System.Console.WriteLine("添加group到GroupTable,ExcelSignal='JC'");
//            //            if (mysql.AddKeyGroup("JC", "Group1"))
//            //                System.Console.WriteLine("成功添加Group1");
//            //            Console.ReadKey();

//            //            System.Console.WriteLine("查询'JC'表的分组信息");
//            //            List<Group> groupList = new List<Group>();
//            //            groupList = mysql.loadGroupForAExcel("JC");
//            //            outputGroupList(groupList);
//            //            Console.ReadKey();

//            //            System.Console.WriteLine("修改Group1的groupName，为Group1Modify");
//            //            int idGroup = getIdGroupFromGroupList(groupList, "Group1");
//            //            if(idGroup == 0)
//            //            {
//            //                System.Console.WriteLine("Group1Modify的id获取有问题");
//            //                Console.ReadKey();
//            //            }               
//            //            if (mysql.ModifyKeyGroup(idGroup, "Group1Modify"))
//            //                System.Console.WriteLine("成功将Group1名称改为Group1Modify");
//            //            Console.ReadKey();

//            //            System.Console.WriteLine("添加一些属性名到Group1Modify");
//            //            List<int> Key_Ids = new List<int>();
//            //            Key_Ids.Add(1);
//            //            Key_Ids.Add(2);
//            //            Key_Ids.Add(3);
//            //            if (mysql.AddKeysToGroup(idGroup, Key_Ids))
//            //                System.Console.WriteLine("添加Key到Group1Modify");
//            //            Console.ReadKey();

//            //            System.Console.WriteLine("通过Group_Id查询所有与之相关的KeyName");
//            //            List<KeyTableRow> keyGroup2 = new List<KeyTableRow>();
//            //            keyGroup2 = mysql.loadKeyNameForAGroup(idGroup);
//            //            outputKeyTableRow(keyGroup2);
//            //            Console.ReadKey();

//            //            System.Console.WriteLine("通过Signal:'JC'来查询与之相关的所有的KeyName");
//            //            List<KeyTableRow> keyGroup1 = new List<KeyTableRow>();
//            //            keyGroup1 = mysql.loadKeyNameForAExcel("JC");
//            //            outputKeyTableRow(keyGroup1);
//            //            Console.ReadKey();
//            ////************************
//            //            System.Console.WriteLine("删除Group1Modify");
//            //            mysql.DeleteKeyGroup(idGroup);
//            //            List<Group> groupList2 = new List<Group>();
//            //            groupList2 = mysql.loadGroupForAExcel("JC");
//            //            outputGroupList(groupList2);
//            //            Console.ReadKey();

//            //System.Console.WriteLine("查询当前所有Excel表信息");
//            //List<ExcelTable> excelTable = new List<ExcelTable>();
//            //excelTable = mysql.ListExcelsMessage();
//            //outputExcelMessageList(excelTable);
//            //Console.ReadKey();

//            System.Console.WriteLine("修改Entity的Remark");
//            string entityName = "1-5底板";
//            string remark = "基础数据--方量";
//            if (mysql.ModifyEntityRemark(entityName, remark))
//            {
//                System.Console.WriteLine("修改Entity的Remark------success");
//            }

//            System.Console.WriteLine("修改阈值Total_hold和Diff_hold");
//            if (mysql.ModifyThreshold("JC", 2, 5))
//            {
//                System.Console.WriteLine("修改阈值Total_hold和Diff_hold------success");
//            }
//        }

//        public static int getIdGroupFromGroupList(List<Group> groupList,string groupName)
//        {
//            int idGroup = 0;
//            foreach(Group g in groupList)
//            {
//                if (g.GroupName.Equals(groupName))
//                    idGroup = g.Id;
//            }
//            return idGroup;
//        }

//        public static void outputKeyTableRow(List<KeyTableRow> keyGroup)
//        {
//            foreach(KeyTableRow keytable in keyGroup)
//            {
//                Console.WriteLine(keytable.Id + " " + keytable.KeyName);
//            }
//        } 

//        public static void outputGroupList(List<Group> groupList)
//        {
//            foreach (Group g in groupList)
//            {
//                Console.WriteLine(g.Id + " " + g.GroupName);
//            }
//        }

//        public static void outputExcelMessageList(List<ExcelTable> excelTable)
//        {
//            foreach(ExcelTable et in excelTable)
//            {
//                Console.WriteLine(et.Signal + " " + et.TableDesc + " " + et.Total_hold + " " + et.Diff_hold);
//            }
//        }
//    }
//}
