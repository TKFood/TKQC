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

        public FrmNUTRITION()
        {
            InitializeComponent();

            comboBox1load();
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

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCHNUTRITIONBASE(comboBox1.Text.Trim());
        }

        #endregion
    }
}
