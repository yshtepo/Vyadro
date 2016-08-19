using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data;
using System.Drawing;

namespace EXtoWD
{
    public partial class Form1 : Form
    {
        public string Anim, Year;
        public Form1()
        {
            
            InitializeComponent();
            comboBox2.Text = "Column";
            comboBox2.Items.Add("Column");
            comboBox2.Items.Add("Lines");
            comboBox2.Items.Add("Pie");
            comboBox2.Items.Add("Bar");
            comboBox2.Items.Add("Funnel");
            comboBox2.Items.Add("PointAndFigure");

            comboBox1.Items.Clear();
            if (System.IO.File.Exists("your_base.mdb"))
            {
                string conStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=your_base.mdb;Jet OLEDB:Engine Type=5";
                System.Data.OleDb.OleDbConnection connectDb = new System.Data.OleDb.OleDbConnection(conStr);
                connectDb.Open();
                DataTable cbTb = connectDb.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                foreach (DataRow row in cbTb.Rows)
                {
                    string tbName = row["TABLE_NAME"].ToString();
                    comboBox1.Items.Add(tbName);
                }
                connectDb.Close();
            }
        }

        /*private void button2_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text.ToString() == "")
            {
                MessageBox.Show("Выбери таблицу");
            }
            else
            {
                string tableName = comboBox1.Text.ToString();
                DataSet chartDs = new DataSet();
                string conStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=your_base.mdb;Jet OLEDB:Engine Type=5";
                System.Data.OleDb.OleDbConnection connectDb = new System.Data.OleDb.OleDbConnection(conStr);
                connectDb.Open();
                string comm = "SELECT * FROM "+ tableName;
                System.Data.OleDb.OleDbDataAdapter dbAdp = new System.Data.OleDb.OleDbDataAdapter(comm, conStr);
                dbAdp.Fill(chartDs);
                dataGridView1.DataSource = chartDs.Tables[0];
                this.dataGridView1.Controls.Clear();
                this.chart1.Controls.Clear();


                //build chart

                this.chart1.Series.Clear();
                for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    System.Windows.Forms.DataVisualization.Charting.Series ser = chart1.Series.Add(dataGridView1.Rows[i].Cells[0].Value.ToString());
                    if (comboBox2.Text.ToString() == "Lines")
                    {
                        ser.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
                    }
                    else if (comboBox2.Text.ToString() == "Pie")
                    {
                        ser.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Pie;
                    }
                    else if (comboBox2.Text.ToString() == "Bar")
                    {
                        ser.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Bar;
                    }
                    else if (comboBox2.Text.ToString() == "Funnel")
                    {
                        ser.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Funnel;
                    }
                    else if (comboBox2.Text.ToString() == "PointAndFigure")
                    {
                        ser.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.PointAndFigure;
                    }
                    else
                    {
                        ser.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
                    }
                    
                }

                for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    for (int j = 1; j < dataGridView1.Columns.Count - 2; j++)
                    {
                        chart1.Series[i].Points.AddXY(dataGridView1.Columns[j].HeaderText, dataGridView1.Rows[i].Cells[j].Value);
                    }
                }

                //chart1.Series[""].Points.AddXY(wSheet.Cells[1, 1].Text.ToString(), wSheet.Cells[2, 1].Text.ToString());
                //chart1.Series["2"].Points.AddXY(wSheet.Cells[1, 2].Text.ToString(), wSheet.Cells[2, 2].Text.ToString());
            }

        }
        */

        private void button3_Click(object sender, EventArgs e)
        {
            this.chart1.Series.Clear();
            for (int i = 0; i < dataGridView1.SelectedRows.Count; i++)
            {
                DataGridViewRow row = dataGridView1.SelectedRows[i];
                System.Windows.Forms.DataVisualization.Charting.Series ser = chart1.Series.Add(row.Cells[0].Value.ToString());
                if (comboBox2.Text.ToString() == "Lines")
                {
                    ser.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
                }
                else if (comboBox2.Text.ToString() == "Pie")
                {
                    ser.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Pie;
                }
                else if (comboBox2.Text.ToString() == "Bar")
                {
                    ser.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Bar;
                }
                else
                {
                    ser.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
                }
            }
            for (int i = 0; i < dataGridView1.SelectedRows.Count; i++)
            {
                for (int j = 1; j < dataGridView1.Columns.Count - 2; j++)
                {
                    chart1.Series[i].Points.AddXY(dataGridView1.Columns[j].HeaderText, dataGridView1.SelectedRows[i].Cells[j].Value);
                }
            }

        }

        private void chart1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                contextMenuStrip1.Show(this, e.Location);
            }
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            //chart1.SaveImage("Chart.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
            //Image img = Image.FromFile("Chart.jpg");
            Bitmap chart = new Bitmap(panel1.Width, chart1.Height);
            panel1.DrawToBitmap(chart, new Rectangle(0, 0, panel1.Width, panel1.Height));
            Clipboard.SetImage(chart);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DataCollect dc = new DataCollect();
            dc.Show();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (System.IO.File.Exists("your_base.mdb"))
            {
                comboBox1.Items.Clear();
                string conStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=your_base.mdb;Jet OLEDB:Engine Type=5";
                System.Data.OleDb.OleDbConnection connectDb = new System.Data.OleDb.OleDbConnection(conStr);
                connectDb.Open();
                DataTable cbTb = connectDb.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                foreach (DataRow row in cbTb.Rows)
                {
                    string tbName = row["TABLE_NAME"].ToString();
                    comboBox1.Items.Add(tbName);
                }
                connectDb.Close();
            }
            else 
            {
                MessageBox.Show("База данных еще не создана. Воспользуйтесь кнопкой 'загрузить данные' для создания БД.", "Внимание!", MessageBoxButtons.OK, 
                    MessageBoxIcon.Warning);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (System.IO.File.Exists("your_base.mdb"))
            {
                this.dataGridView1.Controls.Clear();
                this.chart1.Controls.Clear();
                string tableName = comboBox1.Text.ToString();
                DataSet chartDs = new DataSet();
                string conStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=your_base.mdb;Jet OLEDB:Engine Type=5";
                System.Data.OleDb.OleDbConnection connectDb = new System.Data.OleDb.OleDbConnection(conStr);
                connectDb.Open();
                string comm = "SELECT * FROM " + tableName + "WHERE Anim = '" + this. ;
                System.Data.OleDb.OleDbDataAdapter dbAdp = new System.Data.OleDb.OleDbDataAdapter(comm, conStr);
                dbAdp.Fill(chartDs);
            }
            else
            {
                MessageBox.Show("База данных еще не создана. Воспользуйтесь кнопкой 'загрузить данные' на главном окне для создания БД.", "Внимание!", MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
            
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            groupBox1.Controls.Clear();
            string tableName = comboBox1.Text.ToString();
            DataSet chartDs = new DataSet();
            string conStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=your_base.mdb;Jet OLEDB:Engine Type=5";
            System.Data.OleDb.OleDbConnection connectDb = new System.Data.OleDb.OleDbConnection(conStr);
            connectDb.Open();
            string comm = "SELECT * FROM " + tableName;
            System.Data.OleDb.OleDbDataAdapter dbAdp = new System.Data.OleDb.OleDbDataAdapter(comm, conStr);
            dbAdp.Fill(chartDs);
            dataGridView1.DataSource = chartDs.Tables[0];
            DataTable dt = chartDs.Tables[0];
            Label lb = new Label();
            lb.Text = dataGridView1.Columns[0].HeaderText;
            lb.Location = new Point(10, 20);
            lb.Parent = groupBox1;
            ComboBox cbf = new ComboBox();
            foreach (DataRow row in chartDs.Tables[0].Rows)
            {
                string aName = row[0].ToString();
                cbf.Items.Add(aName);
            }
            cbf.Text = "Не указано";
            cbf.Location = new Point(10, 43);
            cbf.Parent = groupBox1;

            Label lb2 = new Label();
            lb2.Text = dataGridView1.Columns[3].HeaderText;
            lb2.Location = new Point(151, 20);
            lb2.Parent = groupBox1;
            ComboBox cbf2 = new ComboBox();
            foreach (DataRow row in chartDs.Tables[0].Rows)
            {
                string aName = row[3].ToString();
                cbf2.Items.Add(aName);
            }
            cbf2.Text = "Не указано";
            cbf2.Location = new Point(151, 43);
            cbf2.Parent = groupBox1;
            
               
                       
        }


    }
}