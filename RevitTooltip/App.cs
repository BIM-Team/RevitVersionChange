#region Namespaces
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using System.IO;
#endregion // Namespaces

using Res = Revit.Addin.RevitTooltip.Properties.Resources;
using System.Windows.Media.Imaging;
using System.Diagnostics;
using Revit.Addin.RevitTooltip.UI;
using System.Threading;
using System;
using Revit.Addin.RevitTooltip.Intface;
using Revit.Addin.RevitTooltip.Impl;
using Revit.Addin.RevitTooltip.Dto;

namespace Revit.Addin.RevitTooltip
{

    /// <summary>
    /// 插件主程序入口
    /// </summary>
    public class App : IExternalApplication
    {
        /// <summary>
        /// 单列模式
        /// </summary>
        private static App _app = null;

        private string m_previousDocPathName = null;
        //模型联动时使用
        private string _previousselectedNoInfoEntity = null;
        private string selectedNoInfoEntity = null;
        /// <summary>
        /// 模型联动
        /// </summary>
        public string SelectedNoInfoEntity
        {
            set
            {
                this.selectedNoInfoEntity = value;
            }
        }
        //打开文档就构造一份
        private Dictionary<string, ElementId> keyNameToIdMap = null;

        private bool isSettingChange = false;
        /// <summary>
        /// 模型相关的配置
        /// 原本存放在模型中间
        /// </summary>
        internal RevitTooltip settings = null;
        internal RevitTooltip Settings
        {
            set
            {
                if (null == this.settings || !this.settings.Equals(value))
                {
                    this.settings = value;
                    this.mysql = new MysqlUtil(value);
                    this.sqlite = new SQLiteHelper(value);
                    this.isSettingChange = true;

                }
            }
        }
        /// <summary>
        /// 属性面板
        /// </summary>
        ElementInfoPanel m_elementInfoPanel = null;
        /// <summary>
        /// 绘图的控制面板
        /// </summary>
        ImageControl m_ImageControl = null;
        /// <summary>
        /// 用于记录当前选择的内容；用于判断是否和前一次选择的内容相同
        /// </summary>
        private int m_selectedElementId = -1;
        /// <summary>
        /// UI_APP
        /// </summary>
        UIControlledApplication m_uiApp = null;
        /// <summary>
        /// 获取单例
        /// </summary>
        public static App Instance
        {
            get { return _app; }
        }
        /// <summary>
        /// 点击显示属性面板
        /// </summary>
        internal PushButton ElementInfoButton { get; set; }
        /// <summary>
        /// 点击弹出折线图面板
        /// </summary>
        internal PushButton SurveyImageInfoButton { get; set; }


        /// <summary>
        /// 点击重新加载:从Mysql加载到SQLite
        /// </summary>
        internal PushButton LoadDataToSqliteButton { get; set; }
        /// <summary>
        /// MySQL的实例对象
        /// </summary>
        private IMysqlUtil mysql = null;
        /// <summary>
        /// MySQL的实例对象
        /// </summary>
        public IMysqlUtil MySql
        {
            get { return this.mysql; }
        }
        private ISQLiteHelper sqlite = null;
        /// <summary>
        /// 返回SQLite实例
        /// </summary>
        public ISQLiteHelper Sqlite
        {
            get { return this.sqlite; }
        }
        //记录当前打开的文档
        private Document current_doc;
        /// <summary>
        /// 获取当前打开的文档
        /// </summary>
        public Document CurrentDoc
        {
            get { return this.current_doc; }
        }


        public Result OnStartup(UIControlledApplication app)
        {
            //UI对象
            m_uiApp = app;
            //APP的引用
            _app = this;
            //事件绑定：打开了一个文档的视图
            app.ViewActivated += OnViewActivated;
            //事件绑定：闲置事件
            //app.Idling += IdlingHandler;


            //加载格式文件
            string file = Path.Combine(Path.GetDirectoryName(typeof(App).Assembly.Location), "MahApps.Metro.dll");
            if (File.Exists(file))
                System.Reflection.Assembly.LoadFrom(file);
            //属性面板
            m_elementInfoPanel = ElementInfoPanel.GetInstance();
            m_ImageControl = ImageControl.Instance();
            //注册Dockable面板
            app.RegisterDockablePane(new DockablePaneId(m_elementInfoPanel.Id), "构件信息", m_elementInfoPanel);
            app.RegisterDockablePane(new DockablePaneId(m_ImageControl.Id), "测点信息", m_ImageControl);
            //
            //
            string tabName = Res.String_AppTabName;
            string addinAssembly = this.GetType().Assembly.Location;
            string addinDir = Path.GetDirectoryName(addinAssembly);
            string userManual = Path.Combine(addinDir, "help.pdf");
            //
            // 创建帮助选项
            ContextualHelp cHelp = new ContextualHelp(ContextualHelpType.Url, userManual);
            app.CreateRibbonTab(tabName);
            RibbonPanel ribbonPanel = app.CreateRibbonPanel(tabName, Res.String_AppPanelName);

            // 设置按钮
            PushButton cmdButton = (PushButton)ribbonPanel.AddItem(
                new PushButtonData("Tooltip_Settings", Res.CommandName_Settings,
                    addinAssembly, "Revit.Addin.RevitTooltip.CmdSettings"));
            BitmapSource image = Utils.ConvertFromBitmap(Res.settings.ToBitmap());
            cmdButton.Image = cmdButton.LargeImage = image;
            cmdButton.ToolTip = Res.CommandDescription_Settings;
            cmdButton.SetContextualHelp(cHelp);
            //添加分割
            ribbonPanel.AddSeparator();
            //点击显示属性面板
            ElementInfoButton = (PushButton)ribbonPanel.AddItem(
                    new PushButtonData("ElementInfo", Res.Command_ElementInfo,
                        addinAssembly, "Revit.Addin.RevitTooltip.CmdElementInfo"));
            image = Utils.ConvertFromBitmap(Res.tooltip_on.ToBitmap());
            ElementInfoButton.Image = ElementInfoButton.LargeImage = image;
            ElementInfoButton.ToolTip = Res.CommandDescription_TooltipOn;
            ElementInfoButton.SetContextualHelp(cHelp);
            //添加分割
            ribbonPanel.AddSeparator();
            //打开绘图面板
            SurveyImageInfoButton = (PushButton)ribbonPanel.AddItem(
                new PushButtonData("SurveyImageInfo", Res.Command_SurveyImageInfo,
                    addinAssembly, "Revit.Addin.RevitTooltip.CmdImageControl"));
            image = Utils.ConvertFromBitmap(Res.tooltip_on.ToBitmap());
            SurveyImageInfoButton.Image = SurveyImageInfoButton.LargeImage = image;
            SurveyImageInfoButton.ToolTip = Res.CommandDescription_SurveyImage;
            SurveyImageInfoButton.SetContextualHelp(cHelp);
            //添加分割
            ribbonPanel.AddSeparator();
            //从Mysql导入数据到Sqlite
            LoadDataToSqliteButton = (PushButton)ribbonPanel.AddItem(
                new PushButtonData("CommandReloadSQLiteData", Res.Command_ReloadExcelData,
                    addinAssembly, "Revit.Addin.RevitTooltip.CmdLoadSQLiteData"));
            image = Utils.ConvertFromBitmap(Res.refresh);
            LoadDataToSqliteButton.Image = LoadDataToSqliteButton.LargeImage = image;
            LoadDataToSqliteButton.ToolTip = Res.CommandDescription_ReloadExcelData;
            LoadDataToSqliteButton.SetContextualHelp(cHelp);
            // 分割
            ribbonPanel.AddSeparator();

            // About\Help
            PushButtonData aboutButtonData = new PushButtonData("AboutButton",
                Res.CommandName_About, addinAssembly, "Revit.Addin.RevitTooltip.CommandAbout");
            aboutButtonData.Image = Utils.ConvertFromBitmap(Res.about_16.ToBitmap());
            aboutButtonData.ToolTip = Properties.Resources.CommandDescription_About;
            aboutButtonData.SetContextualHelp(cHelp);
            PushButtonData helpButtonData = new PushButtonData("HelpButton",
                Res.CommandName_Help, addinAssembly, "Revit.Addin.RevitTooltip.CommandHelp");
            helpButtonData.Image = Utils.ConvertFromBitmap(Res.help_16.ToBitmap());
            helpButtonData.ToolTip = Properties.Resources.CommandDescription_Help;
            helpButtonData.SetContextualHelp(cHelp);
            ribbonPanel.AddStackedItems(aboutButtonData, helpButtonData);
            //设置成隐藏
            // this.SetPanelEnabled(false);

            return Result.Succeeded;


        }


        private void OnViewActivated(object sender, ViewActivatedEventArgs e)
        {

            try
            {
                if (string.IsNullOrEmpty(m_previousDocPathName) || m_previousDocPathName != e.Document.PathName)
                {
                    settings = ExtensibleStorage.GetTooltipInfo(e.Document.ProjectInformation);
                    this.mysql = new MysqlUtil(settings);
                    this.sqlite = new SQLiteHelper(settings);
                    m_previousDocPathName = e.Document.PathName;
                    current_doc = e.Document;

                    //过滤测点待使用
                    //打开文档就过滤一次
                    keyNameToIdMap = new Dictionary<string, ElementId>();
                    BuiltInParameter testParam = BuiltInParameter.ALL_MODEL_TYPE_NAME;
                    ParameterValueProvider pvp = new ParameterValueProvider(new ElementId(testParam));
                    FilterStringEquals eq = new FilterStringEquals();
                    FilterRule rule = new FilterStringRule(pvp, eq, Res.String_ParameterSurveyType, false);
                    ElementParameterFilter paramFilter = new ElementParameterFilter(rule);
                    Document document = current_doc;
                    FilteredElementCollector elementCollector = new FilteredElementCollector(document).OfClass(typeof(FamilyInstance));
                    IList<Element> elems = elementCollector.WherePasses(paramFilter).ToElements();
                    foreach (var elem in elems)
                    {
                        string param_value = string.Empty;
                        Parameter param = elem.get_Parameter(Res.String_ParameterName);
                        if (null != param && param.StorageType == StorageType.String)
                        {
                            param_value = param.AsString();
                            if (!string.IsNullOrWhiteSpace(param_value))
                            {
                                keyNameToIdMap.Add(param_value, elem.Id);
                            }
                        }
                    }

                }
                //重新打开视图则隐藏Panel
                DockablePane panel = m_uiApp.GetDockablePane(new DockablePaneId(ElementInfoPanel.GetInstance().Id));
                if (panel != null)
                {
                    ElementInfoPanel.GetInstance().Update(new InfoEntityData());
                    panel.Hide();
                }
                DockablePane imageControl = m_uiApp.GetDockablePane(new DockablePaneId(ImageControl.Instance().Id));
                if (imageControl != null)
                {
                    imageControl.Hide();
                }
                //App.Instance.SetPanelEnabled(true);

            }
            catch (System.Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }
        }

        public Result OnShutdown(UIControlledApplication a)
        {

            return Result.Succeeded;
        }
        /// <summary>
        /// Idling event handler.
        /// </summary>
        /// <remarks>
        /// We keep the handler very simple. First check
        /// if we still have the form. If not, unsubscribe 
        /// from Idling, for we no longer need it and it 
        /// makes Revit speedier. If the form is around, 
        /// check if it has a request ready and process 
        /// it accordingly.
        /// </remarks>
        public void IdlingHandler(
          object sender,
          IdlingEventArgs args)
        {
            UIApplication uiapp = sender as UIApplication;
            UIDocument uidoc = uiapp.ActiveUIDocument;

            // UI document is null if the project is closed.

            if (null != uidoc)
            {
                //存放实体
                string entity = string.Empty;
                bool isSurvey = false;
                // ElementInfoPanel
#if (Since2016)
                Element selectElement = uidoc.Document.GetElement(uidoc.Selection.GetElementIds().FirstOrDefault());
#else
                Element selectElement = uidoc.Selection.Elements.Cast<Element>().FirstOrDefault<Element>();
#endif
                if (selectElement != null)
                {
                    entity = Utils.GetParameterValueAsString(selectElement, Res.String_ParameterName);
                    if (!string.IsNullOrEmpty(entity))
                    {
                        isSurvey = Res.String_ParameterSurveyType.Equals(selectElement.Name);

                        if (m_selectedElementId != selectElement.Id.IntegerValue)
                        {
                            m_selectedElementId = selectElement.Id.IntegerValue;
                            // isSurvey = true;
                            if (!isSurvey)
                            {//不是测量数据
                             //sqlite
                                InfoEntityData infoEntityData = Sqlite.SelectInfoData(entity);
                                ElementInfoPanel.GetInstance().Update(infoEntityData);
                            }
                            else
                            {//测量数据绘制折线图
                                DrawEntityData drawEntityData = App.Instance.Sqlite.SelectDrawEntityData(entity, null, null);
                                NewImageForm.Instance().EntityData = drawEntityData;
                                NewImageForm.Instance().Show();
                            }
                        }
                    }
                }
                else
                {
                    if (m_selectedElementId != -1)
                    {//清空数据
                        m_selectedElementId = -1;
                        ElementInfoPanel.GetInstance().Update(null);
                        if (NewImageForm.Instance().Visible)
                        {
                            NewImageForm.Instance().EntityData = new DrawEntityData();
                        }
                    }
                }
            }
        }

        /// <summary>
        /// SettingForm的闲置事件处理函数
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        public void SettingIdlingHandler(
          object sender,
          IdlingEventArgs args)
        {
            UIApplication uiapp = sender as UIApplication;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            if (null != uidoc)
            {
                //把settings保存至模型
                if (isSettingChange)
                {
                    ExtensibleStorage.StoreTooltipInfo(CurrentDoc.ProjectInformation, settings);
                    //App.Instance.CurrentDoc.Save();
                    isSettingChange = false;
                }
            }
        }
        /// <summary>
        /// imageControl的闲置事件处理函数
        /// 模型联动
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        public void ImageControlIdlingHandler(
                  object sender,
                  IdlingEventArgs args)
        {
            UIApplication uiapp = sender as UIApplication;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            if (null != uidoc)
            {
                //模型联动
                if (!string.IsNullOrWhiteSpace(selectedNoInfoEntity) && !selectedNoInfoEntity.Equals(_previousselectedNoInfoEntity))
                {
                    IList<UIView> uiViews = uidoc.GetOpenUIViews();
                    if (uiViews != null && uiViews.Count != 0)
                    {
                        _previousselectedNoInfoEntity = selectedNoInfoEntity;
                        UIView currentUIView = uiViews[uiViews.Count - 1];
                        try
                        {
                            uidoc.ShowElements(keyNameToIdMap[selectedNoInfoEntity]);
                            currentUIView.ZoomSheetSize();
                        }
                        catch (Exception)
                        {
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Gets help document
        /// </summary>
        /// <returns></returns>
        public static string GetHelpDoc()
        {
            string addinAssembly = typeof(Revit.Addin.RevitTooltip.App).Assembly.Location;
            string userManual = Path.Combine(Path.GetDirectoryName(addinAssembly), "help.pdf");
            return userManual;
        }

        public string GetProductName()
        {
            return "BIMRevit2014-2016";
        }


    }
}
