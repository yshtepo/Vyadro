using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace EXtoWD
{
    public partial class DataCollect : Form
    {
        public DataCollect()
        {
            InitializeComponent();
            for (int i = 2000; i <= DateTime.Now.Year; i++)
            {
                comboBox1.Items.Add(i);
            }

            comboBox2.Items.Add("1");
            comboBox2.Items.Add("2");
            comboBox2.Items.Add("3");
            comboBox2.Items.Add("4");
            comboBox2.Items.Add("5");
            comboBox2.Items.Add("6");
            comboBox2.Items.Add("7");
            comboBox2.Items.Add("8");
            comboBox2.Items.Add("9");
            comboBox2.Items.Add("10");
            comboBox2.Items.Add("11");
            comboBox2.Items.Add("12");

        }


        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text.ToString() == "")
            {
                MessageBox.Show("Выбери год", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (comboBox2.Text.ToString() == "")
            {
                MessageBox.Show("Выбери месяц", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (textBox1.Text == "")
            {           
                MessageBox.Show("Выбери файл", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                string strConn;
                //Check for Excel version
                if (textBox1.Text.Substring(textBox1.Text.LastIndexOf('.')).ToLower() == ".xlsx")
                {
                    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + textBox1.Text + ";Extended Properties=\"Excel 12.0;HDR=YES; IMEX=0\"";
                }
                else
                {
                    strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + textBox1.Text + ";Extended Properties=\"Excel 8.0;HDR=YES; IMEX=0\"";
                }

                System.Data.OleDb.OleDbConnection con = new System.Data.OleDb.OleDbConnection(strConn);
                con.Open();
                DataSet ds = new DataSet();
                DataTable shemaTable = con.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables,
                    new object[] { null, null, null, "TABLE" });
                string sheet1 = (string)shemaTable.Rows[0].ItemArray[2];
                string select = String.Format("SELECT * FROM [{0}]", sheet1);
                System.Data.OleDb.OleDbDataAdapter ad = new System.Data.OleDb.OleDbDataAdapter(select, con);
                ad.Fill(ds);
                if (System.IO.File.Exists("wciom_base.mdb"))
                {
                    int year = Convert.ToInt32(comboBox1.Text.ToString());
                    int month = Convert.ToInt32(comboBox2.Text.ToString());
                    string conStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=wciom_base.mdb;Jet OLEDB:Engine Type=5";
                    System.Data.OleDb.OleDbConnection connectDb = new System.Data.OleDb.OleDbConnection(conStr);
                    connectDb.Open();
                    System.Data.OleDb.OleDbCommand myCMD = new System.Data.OleDb.OleDbCommand();
                    myCMD.Connection = connectDb;
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        myCMD.CommandText = "Insert into SURVEY (Anim, weight, height, s_year, s_month) VALUES (\"" +
                            ds.Tables[0].Rows[i][0] + "\", " + ds.Tables[0].Rows[i][1] + ", " + ds.Tables[0].Rows[i][2] + ", " + year + ", " + month + ")";
                        myCMD.ExecuteNonQuery();
                    }
                    MessageBox.Show("Данные загружены в БД", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    int year = Convert.ToInt32(comboBox1.Text.ToString());
                    int month = Convert.ToInt32(comboBox2.Text.ToString());
                    ADOX.Catalog cat = new ADOX.Catalog();
                    string connstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Jet OLEDB:Engine Type=5";
                    cat.Create(String.Format(connstr, "wciom_base.mdb"));
                    cat = null;
                    string conStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=wciom_base.mdb;Jet OLEDB:Engine Type=5";
                    Querry("CREATE TABLE SURVEY(Anim varchar(255), weight int, height int, s_year int, s_month int);", "wciom_base.mdb");
                    System.Data.OleDb.OleDbConnection connectDb = new System.Data.OleDb.OleDbConnection(conStr);
                    connectDb.Open();
                    System.Data.OleDb.OleDbCommand myCMD = new System.Data.OleDb.OleDbCommand();
                    myCMD.Connection = connectDb;
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        myCMD.CommandText = "Insert into SURVEY (Anim, weight, height, s_year, s_month) VALUES (\"" + 
                            ds.Tables[0].Rows[i][0] + "\", " + ds.Tables[0].Rows[i][1] + ", " + ds.Tables[0].Rows[i][2] + ", " + year + ", " + month + ")";
                        myCMD.ExecuteNonQuery();
                    }
                    //string comm = "Insert into SURVEY (Anim, weight, height) VALUES (hare, 10, 20)";
                    //System.Data.OleDb.OleDbDataAdapter dbAdp = new System.Data.OleDb.OleDbDataAdapter(comm, conStr);
                    //dbAdp.Update(ds.Tables[0]);
                    MessageBox.Show("Данные загружены в БД", "", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                con.Close();
                
            }  
        }

        public void Querry(string Que, string DataBasePath)
        {
            string conn = "provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + DataBasePath;
            System.Data.OleDb.OleDbConnection connect = new System.Data.OleDb.OleDbConnection(conn);
            connect.Open();
            using (System.Data.OleDb.OleDbCommand command = new System.Data.OleDb.OleDbCommand(Que, connect))
            {
                try
                {
                    command.ExecuteNonQuery();
                }
                catch (System.Data.OleDb.OleDbException ex)
                {
                    MessageBox.Show("Произошла ошибка при создании таблицы\n" + ex.Message);
                }
            }
            connect.Close();
        }
        

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "*.xls; *.xlsx";
            ofd.Filter = " Excel 2003(*.xls)|*.xls|Excel 2007(*.xlsx)|*.xlsx";
            ofd.Title = "Chose document Excel for upload data";
            ofd.ShowDialog();
            textBox1.Text = ofd.FileName;
        }

    }
}
