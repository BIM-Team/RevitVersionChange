﻿using Res = Revit.Addin.RevitTooltip.Properties.Resources;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Revit.Addin.RevitTooltip.Intface;
using Revit.Addin.RevitTooltip.Impl;
using Revit.Addin.RevitTooltip.Dto;

namespace Revit.Addin.RevitTooltip.UI
{
    public partial class NewSettings : Form
    {
        /// <summary>
        /// 记录当前的设置是否有更改
        /// </summary>
        private bool IsSettingChanged = false;
        /// <summary>
        /// 记录待处理的文件（excel）列表
        /// </summary>
        private string[] fileNames=null;
        /// <summary>
        /// 记录当前的文件集是否已经导入到Mysql
        /// </summary>
        private bool IsInMySql = false;
        /// <summary>
        /// 记录当前的文件集是否已经导入到Sqlite
        /// </summary>
        private bool IsInSqlite = false;
        /// <summary>
        /// Excel的读取工具，该工具只在这里使用，外部使用
        /// </summary>
        private IExcelReader excelReader = null;
        public NewSettings()
        {
            excelReader = new ExcelReader();
            
            InitializeComponent();
        }
        internal NewSettings(RevitTooltip settings)
        {
            if (settings != null) {
                this.textAddr.Text = settings.DfServer;
                this.textDB.Text = settings.DfDB;
                this.textPort.Text = settings.DfPort;
                this.textUser.Text = settings.DfUser;
                this.textPass.Text = settings.DfPassword;
                this.textSqlitePath.Text = settings.SqliteFilePath;
                this.textSqliteName.Text = settings.SqliteFileName;
            }
            excelReader = new ExcelReader();
            InitializeComponent();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = Res.String_SelectExcelFile;
            ofd.DefaultExt = ".xls";
            ofd.FilterIndex = 0;
            ofd.RestoreDirectory = true;
            ofd.Filter = "Excel 97-2003 Workbook(*.xls)|*.xls";
            ofd.Multiselect = true;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                this.fileNames=ofd.FileNames;
                this.IsInMySql = false;
                this.IsInSqlite = false;
                StringBuilder buider = new StringBuilder();
                foreach (string s in this.fileNames) {
                    buider.Append(s).Append(";");
                }
                buider.Remove(buider.Length-1,1);
                this.textFilePath.Text = buider.ToString();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string connectUrl = "server = " + this.textAddr.Text.Trim()+
                "; user =" + this.textUser.Text.Trim() +
                "; database =" + this.textDB.Text.Trim() +
                "; port = " + this.textPort.Text.Trim() +
                "; password =" + this.textPass.Text.Trim() +
                "; charset = utf8";
            MySqlConnection conn = new MySqlConnection(connectUrl);

            try
            {
                conn.Open();
                MessageBox.Show("连接成功");
            }
            catch (Exception ex)
            {
                MessageBox.Show("连接失败：" + ex.Message);
            }
            finally {
                conn.Close();
                conn.Dispose();
            }

        }
        private void textBox_TextChanged(object sender, EventArgs e)
        {
            this.IsSettingChanged = true;

        }

        private void button8_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.Filter = "Sqlite文件（*.db）|*.db";
            saveDialog.FileName = "SqliteDfFile";
            saveDialog.DefaultExt = "db";
            saveDialog.AddExtension = false;
            saveDialog.RestoreDirectory = true;
            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                String fullPath = saveDialog.FileName;
                String path = fullPath.Substring(0, fullPath.LastIndexOf("\\"));
                String fileName = fullPath.Substring(fullPath.LastIndexOf("\\") + 1);
                this.textSqlitePath.Text = path;
                this.textSqliteName.Text = fileName;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (this.IsSettingChanged) {
                button2.Enabled = false;
                RevitTooltip newSetting = RevitTooltip.Default;
                newSetting.DfServer = this.textAddr.Text.Trim();
                newSetting.DfDB = this.textDB.Text.Trim();
                newSetting.DfPort = this.textPort.Text.Trim();
                newSetting.DfUser = this.textUser.Text.Trim();
                newSetting.DfPassword = this.textPass.Text.Trim();
                newSetting.SqliteFilePath = this.textSqlitePath.Text.Trim();
                newSetting.SqliteFileName = this.textSqliteName.Text.Trim();
                App._app.Settings = newSetting;
                ExtensibleStorage.StoreTooltipInfo(App._app.CurrentDoc.ProjectInformation, newSetting);
                App._app.CurrentDoc.Save();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }

        private void buttonInMysql_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("确定导入到Mysql?", "", MessageBoxButtons.OKCancel) ==DialogResult.OK) {
                this.progressBar1.Value = 0;
                for (int i=0;i< fileNames.Length;i++) {
                    SheetInfo info = this.excelReader.loadExcelData(fileNames[i]);
                    App._app.MySql.InsertSheetInfo(info);
                    this.progressBar1.Value =(int) ((i + 1.0) / fileNames.Length) * progressBar1.Maximum;
                }   
                    this.IsInMySql = true;
                    MessageBox.Show("导入数据成功");
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("确定导入到本地？", "", MessageBoxButtons.OKCancel) == DialogResult.OK) {
                this.progressBar1.Value = 0;
                for (int i=0;i<fileNames.Length;i++) {
                   SheetInfo info= this.excelReader.loadExcelData(fileNames[i]);
                    App._app.Sqlite.InsertSheetInfo(info);
                    this.progressBar1.Value = (int)((i + 1.0) / fileNames.Length) * progressBar1.Maximum;
                }
                    this.IsInSqlite = true;
                    MessageBox.Show("导入成功");
            }
        }

        private void tabControl1_Selected(object sender, TabControlEventArgs e)
        {
            TabPage p = e.TabPage;
            if (p == this.tabPageFile) {
                if (App._app.MySql.IsReady())
                {
                    this.buttonInMysql.Enabled = false;
                }
                else {
                    this.buttonInMysql.Enabled = true;
                }
            }
            if (p == this.tabPageThreshold) {
                if (App._app.MySql.IsReady())
                {
                    this.dataGridView1.Enabled = true;
                    List<ExcelTable> tables = App._app.MySql.ListExcelsMessage();                  this.dataGridView1.DataSource = tables;    
                }
                else {
                    this.dataGridView1.Enabled = false;
                    this.dataGridView1.DataSource = null;
                }
            }
            if (p == this.tabPagePro) {
                if (App._app.MySql.IsReady())
                {
                    this.splitContainer1.Enabled = true;
                    this.combExcel.DataSource = App._app.MySql.ListExcelsMessage();
                }
                else {
                    this.splitContainer1.Enabled = false;
                }

            }
            



        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            int col = e.ColumnIndex;
            string Signal = this.dataGridView1.CurrentRow.Cells[0].Value.ToString();
            float value = Convert.ToSingle(this.dataGridView1.CurrentCell.Value);
            if (col == 2)
            {
                App._app.MySql.ModifyThreshold(Signal,value,0);
            }
            else if (col == 3)
            {
                App._app.MySql.ModifyThreshold(Signal,0 ,value);
            }
            else {
                throw new Exception("无效的编辑");
            }
        }

       

        private void combExcel_SelectionChangeCommitted(object sender, EventArgs e)
        {
            string signal = combExcel.SelectedValue.ToString();
            if (!string.IsNullOrWhiteSpace(signal))
            {
                //CombGroup
                List<Group> groups = App._app.MySql.loadGroupForAExcel(signal);
                Group newOne = new Group();
                newOne.GroupName = "新建";
                groups.Add(newOne);
                this.combGroup.DataSource = groups;
                //dataGrid
                List<KeyTableRow> data = App._app.MySql.loadKeyNameForAExcel(signal);
                this.dataGridView2.DataSource = data;
            }
            else {
                throw new Exception("无效的选择");
            }
        }

        private void combGroup_SelectionChangeCommitted(object sender, EventArgs e)
        {
            
            string id = this.combGroup.SelectedValue.ToString();
            if (string.IsNullOrWhiteSpace(id)) {
                throw new Exception("无效的选择");
            }
            int _id = Convert.ToInt32(id);
            List<KeyTableRow> data=App._app.MySql.loadKeyNameForAGroup(_id);
            List<KeyTableRow> datasource = (List<KeyTableRow>)this.combGroup.DataSource;
            //全部先重置
            foreach (DataGridViewRow row in this.dataGridView2.Rows)
            {
                row.Cells[1].Value = 0;
            }
            foreach (var item in data) {
                int index = datasource.IndexOf(item);
                ((DataGridViewCheckBoxCell)this.dataGridView2.Rows[index].Cells[0]).Value = 1;
            }
            
        }

        private void combGroup_SelectedIndexChanged(object sender, EventArgs e)
        {
            string id = combGroup.SelectedValue.ToString();
            string signal = combExcel.SelectedValue.ToString();
            if (string.IsNullOrWhiteSpace(id))
            {
                string newGroup = Microsoft.VisualBasic.Interaction.InputBox("请输入新的组名", "新建分组", "newGroupName", 0, 0);
                if (string.IsNullOrWhiteSpace(newGroup))
                {
                    MessageBox.Show("无效的组名");
                }
                else
                {
                    Group group = App._app.MySql.AddKeyGroup(signal, newGroup);
                    List<Group> groups = (List<Group>)this.combGroup.DataSource;
                    groups.Add(group);
                    this.combGroup.Select(groups.Count - 1, 1);
                }
            }

        }

        private void button5_Click(object sender, EventArgs e)
        {
            string id = this.combGroup.SelectedValue.ToString();
            if (string.IsNullOrWhiteSpace(id)) {
                throw new Exception("无效的选择(GroupName)");
            }
            int _id = Convert.ToInt32(id);
            List<int> keyName_id = new List<int>();
            foreach (DataGridViewRow row in this.dataGridView1.Rows) {
                if (row.Cells[0].Value.ToString().Equals("1")) {
                    keyName_id.Add(Convert.ToInt32(row.Cells[1].Value.ToString()));
                }
            }
            App._app.MySql.AddKeysToGroup(_id, keyName_id);
            MessageBox.Show("添加成功");
        }

        private void button9_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }
    }
}
