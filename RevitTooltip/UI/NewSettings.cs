using Res = Revit.Addin.RevitTooltip.Properties.Resources;
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
            InitializeComponent();
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
                App.Instance.Settings = newSetting;
                this.IsSettingChanged = false;
            }
            button2.Enabled = true;
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
                    App.Instance.MySql.InsertSheetInfo(info);
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
                    App.Instance.Sqlite.InsertSheetInfo(info);
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
                if (App.Instance.MySql.IsReady)
                {
                    this.buttonInMysql.Enabled = true;
                }
                else {
                    this.buttonInMysql.Enabled = false;
                }
            }
            if (p == this.tabPageThreshold) {
                if (App.Instance.MySql.IsReady)
                {
                    this.dataGridView1.Enabled = true;
                    List<ExcelTable> tables = App.Instance.MySql.ListExcelsMessage(false);                  this.dataGridView1.DataSource = tables;    
                }
                else {
                    this.dataGridView1.Enabled = false;
                    this.dataGridView1.DataSource = null;
                }
            }
            if (p == this.tabPagePro) {
                if (App.Instance.MySql.IsReady)
                {
                    this.splitContainer1.Enabled = true;
                    this.combExcel.DataSource = App.Instance.MySql.ListExcelsMessage(true);
                }
                else {
                    this.splitContainer1.Enabled = false;
                }

            }
            



        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            int row = e.RowIndex;
            string Signal = this.dataGridView1.CurrentRow.Cells[1].Value.ToString();
            try
            {
                float Total_hold = float.Parse(this.dataGridView1.CurrentRow.Cells[3].Value.ToString());
                float Diff_hold = float.Parse(this.dataGridView1.CurrentRow.Cells[4].Value.ToString());
                App.Instance.MySql.ModifyThreshold(Signal, Total_hold, Diff_hold);
            }
            catch (Exception) {
                this.dataGridView1.CancelEdit();
                MessageBox.Show("无效的编辑");

            }
        }

       

        private void combExcel_SelectionChangeCommitted(object sender, EventArgs e)
        {
            string signal = combExcel.SelectedValue.ToString();

            if (!string.IsNullOrWhiteSpace(signal))
            {
                //CombGroup
                this.combGroup.DataSource = App.Instance.MySql.loadGroupForAExcel(signal); 
                ////dataGrid
                //List<CKeyName> data = App.Instance.MySql.loadKeyNameForAExcel(signal);
                //this.dataGridView2.DataSource = data;
                //this.dataGridView2.ReadOnly = true;
            }
            else {
                throw new Exception("无效的选择");
            }
        }

        private void combGroup_SelectionChangeCommitted(object sender, EventArgs e)
        {
            
            int id = Convert.ToInt32(this.combGroup.SelectedValue.ToString());
            if (id == -1)
            {
                string newGroupName = Microsoft.VisualBasic.Interaction.InputBox("请输入新的组名", "新建分组", "newGroupName", 30, 30);
                if (string.IsNullOrWhiteSpace(newGroupName))
                {
                    MessageBox.Show("无效的组名");
                }
                else
                {
                string signal = this.combExcel.SelectedValue.ToString();
                Group newOne = App.Instance.MySql.AddNewGroup(signal, newGroupName);
                id = newOne.Id;
                List<Group> groups = App.Instance.MySql.loadGroupForAExcel(signal);
                this.combGroup.DataSource = groups;
                this.combGroup.SelectedIndex= groups.IndexOf(newOne);
                }
            }
            else {
                this.dataGridView2.ReadOnly = false;
                string signal = this.combExcel.SelectedValue.ToString();
                List<CKeyName> data=App.Instance.MySql.loadKeyNameForExcelAndGroup(signal, id);
                this.dataGridView2.DataSource = data;
            }
            
        }


        private void button5_Click(object sender, EventArgs e)
        {
            int group_id = Convert.ToInt32(this.combGroup.SelectedValue.ToString());
            
            List<int> OK_ids = new List<int>();
            foreach (DataGridViewRow row in this.dataGridView2.Rows) {
                if (Convert.ToBoolean(row.Cells[0].Value.ToString()))
                {
                    OK_ids.Add(Convert.ToInt32(row.Cells[1].Value.ToString()));
                }
            }
            if (App.Instance.MySql.AddKeysToGroup(group_id, OK_ids)) {
            MessageBox.Show("修改成功");
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }
    }
}
