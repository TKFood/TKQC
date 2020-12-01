using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using NPOI.SS.UserModel;
using System.Configuration;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
using System.Reflection;
using System.Threading;
using System.Globalization;


namespace TKQC
{
    public partial class FrmNUTRITION : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();       
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds1 = new DataSet();

        string ID;

        public FrmNUTRITION()
        {
            InitializeComponent();

            comboBox1load();
            comboBox1load2();
        }

        #region FUNCTION
        public void comboBox1load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT  [ID],[NAME] FROM [TKQC].[dbo].[NUTRITIONTYPE] ORDER BY ID ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("NAME", typeof(string));

            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "NAME";
            comboBox1.DisplayMember = "NAME";
            sqlConn.Close();


        }

        public void comboBox1load2()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT  [ID],[NAME] FROM [TKQC].[dbo].[NUTRITIONTYPE] ORDER BY ID ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("NAME", typeof(string));

            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "NAME";
            comboBox2.DisplayMember = "NAME";
            sqlConn.Close();


        }

        public void SEARCHNUTRITIONBASE(string TYPE)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
              

                sbSql.AppendFormat(@"  
                                    SELECT
                                    [TYPE] AS '類別'
                                    ,[MB001] AS '品號'
                                    ,[MB002] AS '品名'
                                    ,[CALORIES] AS '熱量Kcal/100g'
                                    ,[FAT] AS '脂肪g/100g'
                                    ,[SATURATEDFAT] AS '飽和脂肪g/100g'
                                    ,[TRANSFAT] AS '反式脂肪g/100g'
                                    ,[CHOLESTEROL] AS '膽固醇mg/100g'
                                    ,[SODIUM] AS '鈉mg/100g'
                                    ,[CARBOHYDRATES] AS '碳水化合物g/100g'
                                    ,[DIETARYFIBER] AS '膳食纖維g/100g'
                                    ,[SUGAR] AS '糖g/100g'
                                    ,[ADDSUGAR] AS '添加糖g/100g'
                                    ,[PROTEIN] AS '蛋白質g/100g'
                                    ,[VITANMIND] AS '維生素D mcg/100g'
                                    ,[CALCIUM] AS '鈣 mg/100g'
                                    ,[IRON] AS '鐵mg/100g'
                                    ,[POTASSIUM] AS '鉀mg/100g'
                                    ,[ID]
                                    FROM [TKQC].[dbo].[NUTRITIONBASE]
                                    WHERE [TYPE]='{0}'
                                    ORDER BY ID
                                    ", TYPE);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                }
                else
                {
                    if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView2.DataSource = ds1.Tables["TEMPds1"];
                        dataGridView2.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                    }
                }

            }
            catch
            {

            }
            finally
            {
                sqlConn.Close();
            }
        }
        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            SETTEXTBOXNULL1();

            if (dataGridView2.CurrentRow != null)
            {
                int rowindex = dataGridView2.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView2.Rows[rowindex];

                    ID = row.Cells["ID"].Value.ToString();
                    comboBox2.Text= row.Cells["類別"].Value.ToString();
                    textBox211.Text = row.Cells["品號"].Value.ToString();
                    textBox212.Text = row.Cells["品名"].Value.ToString();
                    textBox221.Text = row.Cells["熱量Kcal/100g"].Value.ToString();
                    textBox222.Text = row.Cells["脂肪g/100g"].Value.ToString();
                    textBox223.Text = row.Cells["飽和脂肪g/100g"].Value.ToString();
                    textBox224.Text = row.Cells["反式脂肪g/100g"].Value.ToString();
                    textBox231.Text = row.Cells["膽固醇mg/100g"].Value.ToString();
                    textBox232.Text = row.Cells["鈉mg/100g"].Value.ToString();
                    textBox233.Text = row.Cells["碳水化合物g/100g"].Value.ToString();
                    textBox234.Text = row.Cells["膳食纖維g/100g"].Value.ToString();
                    textBox241.Text = row.Cells["糖g/100g"].Value.ToString();
                    textBox242.Text = row.Cells["添加糖g/100g"].Value.ToString();
                    textBox243.Text = row.Cells["蛋白質g/100g"].Value.ToString();
                    textBox244.Text = row.Cells["維生素D mcg/100g"].Value.ToString();
                    textBox251.Text = row.Cells["鈣 mg/100g"].Value.ToString();
                    textBox252.Text = row.Cells["鐵mg/100g"].Value.ToString();
                    textBox253.Text = row.Cells["鉀mg/100g"].Value.ToString();
                }
                else
                {
                    
                }
            }
        }

        public void SETTEXTBOXNULL1()
        {
            textBox211.Text = null;
            textBox212.Text = null;
            textBox221.Text = null;
            textBox222.Text = null;
            textBox223.Text = null;
            textBox224.Text = null;
            textBox231.Text = null;
            textBox232.Text = null;
            textBox233.Text = null;
            textBox234.Text = null;
            textBox241.Text = null;
            textBox242.Text = null;
            textBox243.Text = null;
            textBox244.Text = null;
            textBox251.Text = null;
            textBox252.Text = null;
            textBox253.Text = null;
        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCHNUTRITIONBASE(comboBox1.Text.Trim());
        }

        #endregion

        
    }
}
