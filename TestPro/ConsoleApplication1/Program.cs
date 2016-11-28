//using Revit.Addin.RevitTooltip.Dto;
//using Revit.Addin.RevitTooltip.Impl;
//using Revit.Addin.RevitTooltip.Intface;
//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;

//namespace ConsoleApplication1
//{
//    class Program
//    {
//        static void Main(string[] args)
//        {
//            IExcelReader excel = new ExcelReader();
//            Revit.Addin.RevitTooltip.Dto.SheetInfo result = excel.loadExcelData("C:\\workspace\\慧之建项目\\project\\RevitVersionChange-master\\数据文档\\圈梁-Q-C.xls");
//            Console.WriteLine("下面打印ExcelTableData数据");
//            if (result.ExcelTableData != null)
//            {
//                Console.WriteLine("signal:" + result.ExcelTableData.Signal);
//                Console.WriteLine("Desc:" + result.ExcelTableData.TableDesc);
//            }
//            Console.WriteLine("打印是否为属性数据的标志位：" + result.Tag);

//            if (result.EntityNames != null)
//            {
//                Console.WriteLine("打印实例名列表:列表总数位" + result.EntityNames.Count());
//                foreach (string i in result.EntityNames)
//                {
//                    Console.Write(i + "\t");
//                }
//            }
//            Console.WriteLine();
//            if (result.KeyNames != null)
//            {
//                Console.WriteLine("打印属性名列表：列表长度为：" + result.KeyNames.Count());
//                foreach (string i in result.KeyNames)
//                {
//                    Console.Write(i + "\t");
//                }
//            }
//            Console.WriteLine();
//            /*
//            if (result.InfoRows != null)
//            {
//                Console.WriteLine("打印属性数据::实体数" + result.InfoRows.Count());
//                foreach (InfoEntityData info in result.InfoRows)
//                {
//                    Console.WriteLine(info.EntityName + "--->对应属性数据总共:" + info.Data.Count());
//                    foreach (string key in info.Data.Keys)
//                    {
//                        Console.WriteLine(key + ">>" + info.Data[key]);
//                    }

//                }
//            }
//            Console.WriteLine();
//            if (result.DrawRows != null)
//            {
//                Console.WriteLine("打印测量数据::测量数据的总数" + result.DrawRows.Count());
//                foreach (DrawEntityData d in result.DrawRows)
//                {
//                    Console.WriteLine("测量实体：" + d.EntityName + "对应数据条数：" + d.Data.Count());
//                    foreach (DrawData k in d.Data)
//                    {
//                        Console.WriteLine("时间：：" + k.Date + "\t最大值::" + k.MaxValue + "\t最小值::" + k.MinValue + "\t中位数" + k.MidValue);
//                        Console.WriteLine("详细为：" + k.Detail);
//                    }
//                }
//            }
//            */
//            Console.ReadKey();
//        }
//    }
//}
