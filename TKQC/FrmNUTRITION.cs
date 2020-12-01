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
        int result;

        string ID;

        public FrmNUTRITION()
        {
            InitializeComponent();

            comboBox1load();
            comboBox1load2();
            comboBox1load3();
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

        public void comboBox1load3()
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
            comboBox3.DataSource = dt.DefaultView;
            comboBox3.ValueMember = "NAME";
            comboBox3.DisplayMember = "NAME";
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
                    comboBox2.Text = row.Cells["類別"].Value.ToString();
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

      

        public void UPDATENUTRITIONBASE(
                                        string ID
                                        , string TYPE
                                        , string MB001
                                        , string MB002
                                        , decimal CALORIES
                                        , decimal FAT
                                        , decimal SATURATEDFAT
                                        , decimal TRANSFAT
                                        , decimal CHOLESTEROL
                                        , decimal SODIUM
                                        , decimal CARBOHYDRATES
                                        , decimal DIETARYFIBER
                                        , decimal SUGAR
                                        , decimal ADDSUGAR
                                        , decimal PROTEIN
                                        , decimal VITANMIND
                                        , decimal CALCIUM
                                        , decimal IRON
                                        , decimal POTASSIUM)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();


                sbSql.AppendFormat(@" 
                                    UPDATE  [TKQC].[dbo].[NUTRITIONBASE]
                                    SET [TYPE]='{1}'
                                    ,[MB001]='{2}'
                                    ,[MB002]='{3}'
                                    ,[CALORIES]={4}
                                    ,[FAT]={5}
                                    ,[SATURATEDFAT]={6}
                                    ,[TRANSFAT]={7}
                                    ,[CHOLESTEROL]={8}
                                    ,[SODIUM]={9}
                                    ,[CARBOHYDRATES]={10}
                                    ,[DIETARYFIBER]={11}
                                    ,[SUGAR]={12}
                                    ,[ADDSUGAR]={13}
                                    ,[PROTEIN]={14}
                                    ,[VITANMIND]={15}
                                    ,[CALCIUM]={16}
                                    ,[IRON]={17}
                                    ,[POTASSIUM]={18}
                                    WHERE [ID]={0}
                                   
                                    ", ID, TYPE, MB001, MB002, CALORIES, FAT, SATURATEDFAT, TRANSFAT, CHOLESTEROL, SODIUM, CARBOHYDRATES, DIETARYFIBER, SUGAR, ADDSUGAR, PROTEIN, VITANMIND, CALCIUM, IRON, POTASSIUM);


                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  

                    MessageBox.Show("完成");
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

        public void ADDNUTRITIONBASE(
                                        string ID
                                        , string TYPE
                                        , string MB001
                                        , string MB002
                                        , decimal CALORIES
                                        , decimal FAT
                                        , decimal SATURATEDFAT
                                        , decimal TRANSFAT
                                        , decimal CHOLESTEROL
                                        , decimal SODIUM
                                        , decimal CARBOHYDRATES
                                        , decimal DIETARYFIBER
                                        , decimal SUGAR
                                        , decimal ADDSUGAR
                                        , decimal PROTEIN
                                        , decimal VITANMIND
                                        , decimal CALCIUM
                                        , decimal IRON
                                        , decimal POTASSIUM)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);


                ID = FINDMAXID();

                sqlConn.Close();
                sqlConn.Open();

                sbSql.Clear();
                sbSql.AppendFormat(@" 
                                    INSERT INTO [TKQC].[dbo].[NUTRITIONBASE]
                                    ([ID],[TYPE],[MB001],[MB002],[CALORIES],[FAT],[SATURATEDFAT],[TRANSFAT],[CHOLESTEROL],[SODIUM],[CARBOHYDRATES],[DIETARYFIBER],[SUGAR],[ADDSUGAR],[PROTEIN],[VITANMIND],[CALCIUM],[IRON],[POTASSIUM])
                                    VALUES
                                    ({0},'{1}','{2}','{3}',{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18})
                                   
                                    ", ID, TYPE, MB001, MB002, CALORIES, FAT, SATURATEDFAT, TRANSFAT, CHOLESTEROL, SODIUM, CARBOHYDRATES, DIETARYFIBER, SUGAR, ADDSUGAR, PROTEIN, VITANMIND, CALCIUM, IRON, POTASSIUM);


                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  

                    MessageBox.Show("完成");
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

        public string FINDMAXID()
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
                                    SELECT MAX([ID])+1  AS 'ID' FROM [TKQC].[dbo].[NUTRITIONBASE]
                                    ");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();

                if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["TEMPds1"].Rows[0]["ID"].ToString();
                }
                else
                {
                    return "1";
                }

            }
            catch
            {
                return "1";
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public void DELNUTRITIONBASE(string ID)
        {
            if(!string.IsNullOrEmpty(ID))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();

                    sbSql.Clear();
                    sbSql.AppendFormat(@" 
                                        DELETE [TKQC].[dbo].[NUTRITIONBASE]
                                        WHERE [ID]='{0}'                                   
                                        ", ID);


                    cmd.Connection = sqlConn;
                    cmd.CommandTimeout = 60;
                    cmd.CommandText = sbSql.ToString();
                    cmd.Transaction = tran;
                    result = cmd.ExecuteNonQuery();

                    if (result == 0)
                    {
                        tran.Rollback();    //交易取消
                    }
                    else
                    {
                        tran.Commit();      //執行交易  

                        MessageBox.Show("完成");
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

        public void SETTEXTBOXNULL2()
        {
            textBox311.Text = null;
            textBox312.Text = null;
            textBox321.Text = null;
            textBox322.Text = null;
            textBox323.Text = null;
            textBox324.Text = null;
            textBox331.Text = null;
            textBox332.Text = null;
            textBox333.Text = null;
            textBox334.Text = null;
            textBox341.Text = null;
            textBox342.Text = null;
            textBox343.Text = null;
            textBox344.Text = null;
            textBox351.Text = null;
            textBox352.Text = null;
            textBox353.Text = null;
        }

        public void SETTEXTBOX0()
        {
            textBox311.Text = null;
            textBox312.Text = null;
            textBox321.Text = "0";
            textBox322.Text = "0";
            textBox323.Text = "0";
            textBox324.Text = "0";
            textBox331.Text = "0";
            textBox332.Text = "0";
            textBox333.Text = "0";
            textBox334.Text = "0";
            textBox341.Text = "0";
            textBox342.Text = "0";
            textBox343.Text = "0";
            textBox344.Text = "0";
            textBox351.Text = "0";
            textBox352.Text = "0";
            textBox353.Text = "0";
        }

        public void SETTEXTBOXREADONLY1()
        {
            textBox211.ReadOnly = false;
            textBox212.ReadOnly = false;
            textBox221.ReadOnly = false;
            textBox222.ReadOnly = false;
            textBox223.ReadOnly = false;
            textBox224.ReadOnly = false;
            textBox231.ReadOnly = false;
            textBox232.ReadOnly = false;
            textBox233.ReadOnly = false;
            textBox234.ReadOnly = false;
            textBox241.ReadOnly = false;
            textBox242.ReadOnly = false;
            textBox243.ReadOnly = false;
            textBox244.ReadOnly = false;
            textBox251.ReadOnly = false;
            textBox252.ReadOnly = false;
            textBox253.ReadOnly = false;
        }

        public void SETTEXTBOXREADONLY2()
        {
            textBox211.ReadOnly = true;
            textBox212.ReadOnly = true;
            textBox221.ReadOnly = true;
            textBox222.ReadOnly = true;
            textBox223.ReadOnly = true;
            textBox224.ReadOnly = true;
            textBox231.ReadOnly = true;
            textBox232.ReadOnly = true;
            textBox233.ReadOnly = true;
            textBox234.ReadOnly = true;
            textBox241.ReadOnly = true;
            textBox242.ReadOnly = true;
            textBox243.ReadOnly = true;
            textBox244.ReadOnly = true;
            textBox251.ReadOnly = true;
            textBox252.ReadOnly = true;
            textBox253.ReadOnly = true;
        }

        public void SETTEXTBOXREADONLY3()
        {
            textBox311.ReadOnly = false;
            textBox312.ReadOnly = false;
            textBox321.ReadOnly = false;
            textBox322.ReadOnly = false;
            textBox323.ReadOnly = false;
            textBox324.ReadOnly = false;
            textBox331.ReadOnly = false;
            textBox332.ReadOnly = false;
            textBox333.ReadOnly = false;
            textBox334.ReadOnly = false;
            textBox341.ReadOnly = false;
            textBox342.ReadOnly = false;
            textBox343.ReadOnly = false;
            textBox344.ReadOnly = false;
            textBox351.ReadOnly = false;
            textBox352.ReadOnly = false;
            textBox353.ReadOnly = false;
        }

        public void SETTEXTBOXREADONLY4()
        {
            textBox311.ReadOnly = true;
            textBox312.ReadOnly = true;
            textBox321.ReadOnly = true;
            textBox322.ReadOnly = true;
            textBox323.ReadOnly = true;
            textBox324.ReadOnly = true;
            textBox331.ReadOnly = true;
            textBox332.ReadOnly = true;
            textBox333.ReadOnly = true;
            textBox334.ReadOnly = true;
            textBox341.ReadOnly = true;
            textBox342.ReadOnly = true;
            textBox343.ReadOnly = true;
            textBox344.ReadOnly = true;
            textBox351.ReadOnly = true;
            textBox352.ReadOnly = true;
            textBox353.ReadOnly = true;
        }


        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCHNUTRITIONBASE(comboBox1.Text.Trim());
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SETTEXTBOXREADONLY1();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            UPDATENUTRITIONBASE(ID,comboBox2.Text ,textBox211.Text.Trim(),textBox212.Text,Convert.ToDecimal(textBox221.Text),Convert.ToDecimal(textBox222.Text),Convert.ToDecimal(textBox223.Text),Convert.ToDecimal(textBox224.Text),Convert.ToDecimal(textBox231.Text),Convert.ToDecimal(textBox232.Text),Convert.ToDecimal(textBox233.Text),Convert.ToDecimal(textBox234.Text),Convert.ToDecimal(textBox241.Text),Convert.ToDecimal(textBox242.Text),Convert.ToDecimal(textBox243.Text),Convert.ToDecimal(textBox244.Text),Convert.ToDecimal(textBox251.Text),Convert.ToDecimal(textBox252.Text),Convert.ToDecimal(textBox253.Text));
            SETTEXTBOXREADONLY2();

            SEARCHNUTRITIONBASE(comboBox1.Text.Trim());
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELNUTRITIONBASE(ID);
                SEARCHNUTRITIONBASE(comboBox1.Text.Trim());
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }
        private void button5_Click(object sender, EventArgs e)
        {
            SETTEXTBOXREADONLY3();
            SETTEXTBOX0();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            ADDNUTRITIONBASE(ID, comboBox3.Text, textBox311.Text.Trim(), textBox312.Text, Convert.ToDecimal(textBox321.Text), Convert.ToDecimal(textBox322.Text), Convert.ToDecimal(textBox323.Text), Convert.ToDecimal(textBox324.Text), Convert.ToDecimal(textBox331.Text), Convert.ToDecimal(textBox332.Text), Convert.ToDecimal(textBox333.Text), Convert.ToDecimal(textBox334.Text), Convert.ToDecimal(textBox341.Text), Convert.ToDecimal(textBox342.Text), Convert.ToDecimal(textBox343.Text), Convert.ToDecimal(textBox344.Text), Convert.ToDecimal(textBox351.Text), Convert.ToDecimal(textBox352.Text), Convert.ToDecimal(textBox353.Text));

            SETTEXTBOXNULL2();           

            SETTEXTBOXREADONLY4();
            SEARCHNUTRITIONBASE(comboBox1.Text.Trim());

        }
        #endregion


    }
}
