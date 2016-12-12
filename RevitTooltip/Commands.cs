#region Namespaces
using System;
using System.IO;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
#endregion

using System.Windows.Forms;
using Revit.Addin.RevitTooltip.UI;

namespace Revit.Addin.RevitTooltip
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    [Autodesk.Revit.Attributes.Journaling(Autodesk.Revit.Attributes.JournalingMode.NoCommandData)]
    public abstract class TooltipCommandBase : IExternalCommand
    {
        #region Class Members
        public Autodesk.Revit.ApplicationServices.Application RevitApp = null;
        public UIDocument RevitUiDoc = null;
        public Document RevitDoc = null;
        #endregion

        #region Class Implementation
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // member initialization
            RevitApp = commandData.Application.Application;
            RevitUiDoc = commandData.Application.ActiveUIDocument;
            RevitDoc = RevitUiDoc.Document;
            //
            // command run
            return RunIt(commandData, ref message, elements);
        }


        #endregion

        /// <summary>
        /// Implements this method to run the external command
        /// </summary>
        /// <param name="commandData"></param>
        /// <param name="message"></param>
        /// <param name="elements"></param>
        /// <returns></returns>
        public abstract Result RunIt(ExternalCommandData commandData, ref string message, Autodesk.Revit.DB.ElementSet elements);

        /// <summary>
        /// Indicates whether administrator is required to run this command.
        /// </summary>
        /// <returns></returns>
        protected virtual bool RequireAdmin()
        {
            return false;
        }
    }

    [Transaction(TransactionMode.Manual)]
    public class CmdSettings : TooltipCommandBase
    {
        public override Result RunIt(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                NewSettings settingForm = new NewSettings(App.Instance.settings);
                settingForm.Show();
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return Result.Failed;
            }
        }
    }

    

    #region CommandReloadExcelData
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    [Autodesk.Revit.Attributes.Journaling(Autodesk.Revit.Attributes.JournalingMode.NoCommandData)]
    public class CmdLoadSQLiteData : TooltipCommandBase
    {
        public override Result RunIt(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                if (App.Instance.Sqlite.LoadDataToSqlite()) {
                MessageBox.Show("数据更新成功");
                }
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return Result.Failed;
            }
        }
    }
    #endregion

    #region CommandAbout
    /// <summary>
    /// This command allows user to show about dialog.
    /// </summary>
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    [Autodesk.Revit.Attributes.Journaling(Autodesk.Revit.Attributes.JournalingMode.NoCommandData)]
    public class CommandAbout : TooltipCommandBase
    {
        public override Result RunIt(ExternalCommandData commandData, ref string message, Autodesk.Revit.DB.ElementSet elements)
        {
            // About form may needs 
            AboutForm aboutFrm = new AboutForm();
            aboutFrm.ShowDialog();
            return Result.Succeeded;
        }
    }
    #endregion

    #region CommandHelp
    /// <summary>
    /// This command allows user to open the help document directly
    /// </summary>
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    [Autodesk.Revit.Attributes.Journaling(Autodesk.Revit.Attributes.JournalingMode.NoCommandData)]
    public class CommandHelp : TooltipCommandBase
    {
        public override Result RunIt(ExternalCommandData commandData, ref string message, Autodesk.Revit.DB.ElementSet elements)
        {
            // open the help document directly
            try
            {
                string helpDoc = App.GetHelpDoc();
                System.Diagnostics.Process.Start(helpDoc);
            }
            catch (System.Exception)
            {
                MessageBox.Show("打开帮助文件失败！请联系开发技术支持人员！");
            }
            return Result.Succeeded;
        }
    }
    #endregion

    [Transaction(TransactionMode.Manual)]
    public class CmdElementInfo : TooltipCommandBase
    {
        public override Result RunIt(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            if (!File.Exists(Path.Combine(App.Instance.settings.SqliteFilePath, App.Instance.settings.SqliteFileName)))
            {
                MessageBox.Show("本地数据文件不存在，请先更新");
                return Result.Succeeded;
            }
            DockablePane panel = commandData.Application.GetDockablePane(new DockablePaneId(ElementInfoPanel.GetInstance().Id));
            panel.Show();
            commandData.Application.Idling += App.Instance.IdlingHandler;
            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    public class CmdImageControl : TooltipCommandBase
    {
        public override Result RunIt(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            if (!File.Exists(Path.Combine(App.Instance.settings.SqliteFilePath, App.Instance.settings.SqliteFileName))) {
                MessageBox.Show("本地数据文件不存在，请先更新");
                return Result.Succeeded;
            }
            ImageControl.Instance().setExcelType(App.Instance.Sqlite.SelectDrawTypes());
            DockablePane imagePanel = commandData.Application.GetDockablePane(new DockablePaneId(ImageControl.Instance().Id));
            imagePanel.Show();
            commandData.Application.Idling += App.Instance.IdlingHandler;

            //对模型的处理，后续可能删除
        //Material color_red = null;
        //Material color_gray = null;
        //Material color_blue = null;
        //FilteredElementCollector elementCollector = new FilteredElementCollector(commandData.Application.ActiveUIDocument.Document);
        //IEnumerable<Material> allMaterial = elementCollector.OfClass(typeof(Material)).Cast<Material>();
        //        foreach (Material elem in allMaterial)
        //        {
        //            if (elem.Name.Equals(Res.String_Color_Red))
        //            {
        //                color_red = elem;
        //            }
        //            if (elem.Name.Equals(Res.String_Color_Gray))
        //            {
        //                color_gray = elem;
        //            }
        //            if (elem.Name.Equals(Res.String_Color_Blue))
        //            {
        //                color_blue = elem;
        //            }
        //            if (color_gray != null && color_red != null&& color_blue!=null)
        //            {
        //                break;
        //            }
        //        }

           
            return Result.Succeeded;
        }
    }

    //[Transaction(TransactionMode.Manual)]
    //public class CmdLoadExcelToDB : TooltipCommandBase
    //{
    //    public override Result RunIt(ExternalCommandData commandData, ref string message, ElementSet elements)
    //    {
    //        //消息对话框确认
    //        if (MessageBox.Show("确认导入Excel表格数据到Mysql数据库？","",MessageBoxButtons.OKCancel) == DialogResult.Cancel)
    //        {
    //            return Result.Succeeded;
    //        }
    //        ProcessBarForm processBarForm = new ProcessBarForm( MysqlUtil.CreateInstance());
    //        if (processBarForm.ShowDialog() == DialogResult.OK) {
    //        processBarForm.Dispose();
    //        }    
    //        return Result.Succeeded;
    //    }
    //}
    //[Transaction(TransactionMode.Manual)]
    //public class CmdLoadExcelToSQLite : TooltipCommandBase
    //{
    //    public override Result RunIt(ExternalCommandData commandData, ref string message, ElementSet elements)
    //    {

    //        //消息对话框确认
    //        if (MessageBox.Show("确认导入Excel表格数据到本地数据库？", "", MessageBoxButtons.OKCancel) == DialogResult.Cancel)
    //        {
    //            return Result.Succeeded;
    //        }
    //        ProcessBarForm processBarForm = new ProcessBarForm(SQLiteHelper.CreateInstance());
    //        if (processBarForm.ShowDialog() == DialogResult.OK)
    //        {
    //            processBarForm.Dispose();
    //        }
    //        return Result.Succeeded;
    //    }
    //}


}
