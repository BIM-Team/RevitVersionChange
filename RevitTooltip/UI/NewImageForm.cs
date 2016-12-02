using Revit.Addin.RevitTooltip.Dto;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Revit.Addin.RevitTooltip.UI
{
    public partial class NewImageForm : Form
    {
        public NewImageForm()
        {
            InitializeComponent();
        }
        /// <summary>
        /// 数据，根据这个数据绘制折线图
        /// </summary>
        private List<DrawData> data;

        /// <summary>
        /// 数据，根据这个数据绘制折线图
        /// </summary>
        public List<DrawData> Data {
            get {
                return this.data;
            }
            set {
                if (!value.Equals(this.data)) {
                    this.data = value;
                    this.panel1.Invalidate(this.panel1.ClientRectangle);
                }
            }
        }
        private void panel1_Paint(object sender, PaintEventArgs e)
        {

            Graphics g = e.Graphics;

            float height = this.panel1.ClientRectangle.Height;
            float width = this.panel1.ClientRectangle.Width;
            float startX = 40, endX = width - 10;
            float startY = height - 30, endY = 10;
            Font font = new Font("Arial", 9, System.Drawing.FontStyle.Regular);
            if (null == data || data.Count == 0)
            {
                g.DrawString("没有数据", font, Brushes.Black, (startX + endX - g.MeasureString("没有数据", font).Width) / 2, (startY + endY) / 2);
                return;
            }
            int length = data.Count;
            int div = length / 5;

            float  Max = data[0].MaxValue;
            foreach (DrawData row in data) {
                if (Max < row.MaxValue) {
                    Max = row.MaxValue;
                }
            }
            float divX = (endX - startX) / length;
            float divY = (startY - endY) / Max;
            try
            {
                //清除屏幕
                g.Clear(System.Drawing.Color.White);
                //

                //用于画正常的线段
                Pen mypen = new Pen(System.Drawing.Color.Blue, 1);
                mypen.DashStyle = DashStyle.Dash;
                //画坐标轴使用
                Pen mypen1 = new Pen(System.Drawing.Color.Blue, 2);
                //用于画错误的线段
                Pen pen_error = new Pen(System.Drawing.Color.Red, 2);
                pen_error.DashStyle = DashStyle.Dash;
                //
                Pen pen_error1 = new Pen(System.Drawing.Color.Green, 2);
                pen_error1.DashStyle = DashStyle.Dash;
                //用于连接xy轴
                Pen dotPen = new Pen(System.Drawing.Color.Black, 0.5f);
                dotPen.DashStyle = DashStyle.Dot;
                Pen dotPen1 = new Pen(System.Drawing.Color.Red, 0.5f);
                dotPen1.DashStyle = DashStyle.Dot;
                //画X轴
                g.DrawLine(mypen1, startX, startY, endX, startY);
                //画Y轴
                g.DrawLine(mypen1, startX, endY, startX, startY);
                float y_b = startY;
                float x_b = startX;
                float x = startX - divX;
                float y = startY;
                float value_b = 0;
                int num = 0;
                foreach (DrawData row in data)
                {
                    ////x轴的字
                    float value = row.MaxValue;
                    string str = row.Date.ToShortDateString();
                    x += divX;
                    y = startY - value * divY;
                    //
                    ////y轴的字
                    if (num % div == 0 && (length - num) >= div)
                    {

                        g.DrawString(str, font, Brushes.Black, x - g.MeasureString(str, font).Width / 2, startY + g.MeasureString(str, font).Height / 2);
                        g.DrawString(value.ToString(), font, Brushes.Black, startX - g.MeasureString(value.ToString(), font).Width, y - g.MeasureString(value.ToString(), font).Height / 2);
                        g.DrawLine(dotPen, startX, y, x, y);
                        g.DrawLine(dotPen, x, y, x, startY);
                    }
                    if (num == length - 1)
                    {
                        g.DrawString(str, font, Brushes.Black, endX - g.MeasureString(str, font).Width, startY + g.MeasureString(str, font).Height / 2);
                        g.DrawString(value.ToString(), font, Brushes.Black, startX - g.MeasureString(value.ToString(), font).Width, y - g.MeasureString(value.ToString(), font).Height / 2);
                        g.DrawLine(dotPen, startX, y, x, y);
                        g.DrawLine(dotPen, x, y, x, startY);
                    }
                    num++;

                    if (value_b != 0)
                    {
                        if (value > App.settings.AlertNumber)
                        {
                            if (Math.Abs(x - x_b) < 1 || Math.Abs(y - y_b) < 1)
                            {
                                g.DrawLine(pen_error, x_b, y_b, x + 1, y + 1);
                            }
                            else
                            {
                                g.DrawLine(pen_error, x_b, y_b, x, y);
                            }

                        }
                        else if (Math.Abs(value - value_b) > App.settings.AlertNumberAdd)
                        {

                            if (Math.Abs(x - x_b) < 1 || Math.Abs(y - y_b) < 1)
                            {
                                g.DrawLine(pen_error1, x_b, y_b, x + 1, y + 1);
                            }
                            else
                            {
                                g.DrawLine(pen_error1, x_b, y_b, x, y);
                            }
                        }
                        else
                        {
                            g.DrawLine(mypen, x_b, y_b, x, y);
                        }
                    }


                    y_b = y;
                    x_b = x;
                    value_b = value;

                }
                float alert = (float)(startY - App.settings.AlertNumber * divY);
                g.DrawString(App.settings.AlertNumber.ToString(), font, Brushes.Red, startX -
                    g.MeasureString(App.settings.AlertNumber.ToString(), font).Width, alert - g.MeasureString(App.settings.AlertNumber.ToString(), font).Height / 2);
                g.DrawLine(dotPen1, startX, alert, endX, alert);
                mypen.Dispose();
                mypen1.Dispose();
                dotPen.Dispose();
                dotPen1.Dispose();
                pen_error.Dispose();
                pen_error1.Dispose();
                g.Dispose();

            }
            catch (Exception exception)
            {
                throw exception;
            }
        }
    }
}
