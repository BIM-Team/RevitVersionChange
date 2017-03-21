using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Revit.Addin.RevitTooltip.Dto;
using System.Drawing.Drawing2D;

namespace Revit.Addin.RevitTooltip.UI
{
    public partial class ChildForm : Form
    {
        public ChildForm()
        {
            InitializeComponent();
            this.dataGridView1.AutoGenerateColumns = false;
            List<CEntityName> items= App.Instance.Sqlite.SelectAllEntitiesAndErr("CX");
            this.comboBox1.DisplayMember = "EntityName";
            this.comboBox1.ValueMember = "EntityName";
            this.comboBox1.DataSource = items;
            if (items.Count != 0) {
                this.comboBox1.SelectedIndex = 0;
            }
        }
        /// <summary>
        /// 引用父窗口
        /// </summary>
        public NewImageForm FatherForm { get; set; }
        //private List<CEntityName> all_entity;
        public List<CEntityName> All_Entities { set {
                comboBox1.DataSource = value;
            } }

        private List<DrawData> details = new List<DrawData>();
        

        private void splitContainer1_Panel2_Paint(object sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            float height = this.splitContainer1.Panel2.ClientRectangle.Height;
            float width = this.splitContainer1.Panel2.ClientRectangle.Width;
            float startX = width / 10, endX = width - 10;
            float startY = height - 30, endY = 10;
            Font font = new Font("Arial", 9, System.Drawing.FontStyle.Regular);
            if (details.Count == 0)
            {
                g.DrawString("没有数据", font, Brushes.Black, (startX + endX - g.MeasureString("没有数据", font).Width) / 2, (startY + endY) / 2);
                return;
            }
            float MaxValue = 0L;
            float MinValue = float.MaxValue;
            int CountX = 0;
            foreach (DrawData one in details) {
                String[] arr = one.Detail.Split(';');
                int len = arr.Count();
                float v_max = one.MaxValue;
                float v_min = one.MinValue;
                if (v_max - MaxValue > 0.01){ MaxValue = v_max; }
                if (v_min - MinValue < 0.01) { MinValue = v_min; }
                if (len > CountX) { CountX = len; }
            }
            float divX = (endX - startX) / CountX;
            float divY =  (startY-endY)/ (MaxValue - MinValue);

            float divYY = (startY - endY) / 10;
            float divYV = (MaxValue - MinValue) / 10;
                Pen mypen = new Pen(System.Drawing.Color.Blue, 1);
                //画坐标轴使用
                Pen mypen1 = new Pen(System.Drawing.Color.Blue, 2);
                Pen dotPen = new Pen(Color.FromArgb(128, Color.Black), 0.3f);
            try
            {
                g.Clear(System.Drawing.Color.White);
                //用于画正常的线段
                dotPen.DashStyle = DashStyle.Dot;
                //画X轴
                g.DrawLine(mypen1, startX, startY, endX, startY);
                //画Y轴
                g.DrawLine(mypen1, startX, endY, startX, startY);
                //画横线
                for (int i = 0; i <= 10; i++)
                {
                    float newY = startY - i * divYY;
                    float v_Y = (float)Math.Round(MinValue + i * divYV, 2, MidpointRounding.AwayFromZero);
                    g.DrawLine(dotPen, startX, newY, endX, newY);
                    String s_Y = v_Y.ToString();
                    g.DrawString(s_Y, font, Brushes.Black, startX - g.MeasureString(s_Y, font).Width, newY);
                }
                //画竖线
                for (int j = 0; j <= CountX; j++)
                {
                    float newX = startX + j * divX;
                    g.DrawLine(dotPen, newX, startY, newX, endY);
                    float v_X = (float)(j * 0.5);
                    String s_X = v_X.ToString();
                    if (j!=0&&j % 5 == 0&&CountX-j>2) {
                    g.DrawString(s_X, font, Brushes.Black, newX- g.MeasureString(s_X, font).Width/2, startY + g.MeasureString(s_X, font).Height/2);
                    }
                    if (j == CountX) {
                    g.DrawString(s_X, font, Brushes.Black, newX - g.MeasureString(s_X, font).Width, startY + g.MeasureString(s_X, font).Height/2);
                    }
                }
                foreach (DrawData one in details)
                {
                    String[] arr = one.Detail.Split(';');
                    int len = arr.Count();
                    float pre_x = 0;
                    float pre_y = 0;
                    Random radom = new Random();
                    int c_r = radom.Next(100,255);
                    int c_g = radom.Next(100,255);
                    int c_b = 0;
                    if (c_r + c_g < 400) { c_b = 400 - c_r - c_g; }

                    Pen temp_pen = new Pen(Color.FromArgb(c_r, c_g, c_b), 1);
                    for (int h = 0; h < len; h++)
                    {
                        String[] s_arr = arr[h].Split(':');
                        int y_index = (int)(Convert.ToSingle(s_arr[0]) / 0.5);
                        float v_y = Convert.ToSingle(s_arr[1]);
                        float curr_x = startX + y_index * divX;
                        float curr_y = startY - (v_y - MinValue) * divY;

                        if (h != 0)
                        {
                            g.DrawLine(temp_pen, pre_x, pre_y, curr_x, curr_y);
                            if (h == len / 2) {
                                g.DrawString(one.UniId, font, Brushes.Black, (pre_x+curr_x-g.MeasureString(one.UniId,font).Width)/2,(pre_y+curr_y)/2);
                            }
                        }
                        pre_x = curr_x;
                        pre_y = curr_y;
                    }
                    temp_pen.Dispose();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally {
                g.Dispose();
                mypen.Dispose();
                mypen1.Dispose();
                dotPen.Dispose();
            }

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            CEntityName item = comboBox1.SelectedItem as CEntityName;
            if (item != null) {
                DrawEntityData drawEntityData= App.Instance.Sqlite.SelectDrawEntityData(item.EntityName,null,null);
                this.dataGridView1.DataSource = drawEntityData.Data;
            }
        }

        

        private void label2_Click(object sender, EventArgs e)
        {
            details = new List<DrawData>();
            this.splitContainer1.Panel2.Invalidate(this.splitContainer1.Panel2.ClientRectangle);

        }

        private void ChildForm_VisibleChanged(object sender, EventArgs e)
        {
            if (this.Visible && this.FatherForm.Visible) {
                this.FatherForm.Visible = false;
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            List<DrawData> dataViewSource = this.dataGridView1.DataSource as List<DrawData>;
            int index = this.dataGridView1.CurrentRow.Index;
            if (dataViewSource != null && index < dataViewSource.Count && index >= 0)
            {
                this.details.Add(dataViewSource[this.dataGridView1.CurrentRow.Index]);
                this.splitContainer1.Panel2.Invalidate(this.splitContainer1.Panel2.ClientRectangle);
            }

        }
    }
}
