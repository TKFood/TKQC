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
using FastReport;
using FastReport.Data;
using TKITDLL;

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

            comboBoxload();
            comboBoxload2();
            comboBoxload3();
            comboBoxload4();
            comboBoxload6();
            comboBox7load("穀類");
        }

        #region FUNCTION
        public void comboBoxload()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);


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

        public void comboBoxload2()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);


            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT  [ID],[NAME] FROM [TKQC].[dbo].[NUTRITIONTYPE] ORDER BY ID ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));

            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "NAME";
            comboBox2.DisplayMember = "NAME";
            sqlConn.Close();


        }

        public void comboBoxload3()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);


            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT  [ID],[NAME] FROM [TKQC].[dbo].[NUTRITIONTYPE] ORDER BY ID ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));

            da.Fill(dt);
            comboBox3.DataSource = dt.DefaultView;
            comboBox3.ValueMember = "NAME";
            comboBox3.DisplayMember = "NAME";
            sqlConn.Close();


        }

        public void comboBoxload4()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);


            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT  [ID],[NAME] FROM [TKQC].[dbo].[NUTRITIONTYPE] ORDER BY ID ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));

            da.Fill(dt);
            comboBox4.DataSource = dt.DefaultView;
            comboBox4.ValueMember = "ID";
            comboBox4.DisplayMember = "NAME";
            sqlConn.Close();


        }

        public void comboBoxload6()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);


            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT  [ID],[NAME] FROM [TKQC].[dbo].[NUTRITIONTYPE] ORDER BY ID ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));

            da.Fill(dt);
            comboBox6.DataSource = dt.DefaultView;
            comboBox6.ValueMember = "ID";
            comboBox6.DisplayMember = "NAME";
            sqlConn.Close();

      

        }


        public void SEARCHNUTRITIONBASE(string TYPE)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);



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
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);



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
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);




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
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);



                sbSql.Clear();


                sbSql.AppendFormat(@"  
                                    SELECT ISNULL(MAX([ID])+1,1)  AS 'ID' FROM [TKQC].[dbo].[NUTRITIONBASE]
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

        public string FINDMAXNUTRITIONPRODDETAILID()
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);



                sbSql.Clear();


                sbSql.AppendFormat(@"  
                                    SELECT ISNULL(MAX([ID])+1,1)  AS 'ID' FROM [TKQC].[dbo].[NUTRITIONPRODDETAIL]
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
                    //20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);



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

        public void SETTEXTBOXNULL3()
        {
            textBox5.Text = null;
            textBox6.Text = null;
            
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


        public void SERACHNUTRITIONPROD(string MB002)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);



                sbSql.Clear();

                if(!string.IsNullOrEmpty(MB002))
                {
                    sbSql.AppendFormat(@"
                                        SELECT  [PRODID] AS '成品編號',[PRODNAME]  AS '成品名' 
                                        FROM [TKQC].[dbo].[NUTRITIONPROD]
                                        WHERE([PRODID] LIKE '%{0}%' OR[PRODNAME] LIKE '%{0}%')
                                        ORDER BY [PRODID],[PRODNAME]
                                    ", MB002);
                }
                else
                {
                    sbSql.AppendFormat(@"
                                        SELECT  [PRODID] AS '成品編號',[PRODNAME]  AS '成品名' 
                                        FROM [TKQC].[dbo].[NUTRITIONPROD]                                       
                                        ORDER BY [PRODID],[PRODNAME]
                                    ");
                }
               

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView1.DataSource = ds1.Tables["TEMPds1"];
                        dataGridView1.AutoResizeColumns();
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

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            string PRODID = null;

            textBox7.Text = null;
            textBox8.Text = null;
            textBox9.Text = null;
            textBox10.Text = null;
            textBox14.Text = null;

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];

                    PRODID = row.Cells["成品編號"].Value.ToString();
                    textBox7.Text = row.Cells["成品編號"].Value.ToString();
                    textBox8.Text = row.Cells["成品名"].Value.ToString();
                    textBox9.Text = row.Cells["成品編號"].Value.ToString();
                    textBox10.Text = row.Cells["成品名"].Value.ToString();
                    textBox14.Text = row.Cells["成品編號"].Value.ToString();

                    SERACHNUTRITIONPRODDETAIL(PRODID);
                }
                else
                {

                }
            }
        }

        public void SERACHNUTRITIONPRODDETAIL(string PRODID)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);



                sbSql.Clear();

                sbSql.AppendFormat(@"
                                   
                                    SELECT  
                                    [PRODID] AS '成品編號'
                                    ,[PRODNAME] AS '成品名'
                                    ,[MB001] AS '原料編號'
                                    ,[MB002] AS '原料名'
                                    ,[USEDANOUNT] AS '添加量'
                                    ,[ID]
                                    FROM [TKQC].[dbo].[NUTRITIONPRODDETAIL]
                                    WHERE [PRODID]='{0}'
                                    ORDER BY [PRODID],[MB001]
                                    ", PRODID);


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                {
                    dataGridView3.DataSource = null;
                }
                else
                {
                    if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView3.DataSource = ds1.Tables["TEMPds1"];
                        dataGridView3.AutoResizeColumns();
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

        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            textBox2.Text = null;
            textBox3.Text = null;
            textBox4.Text = null;
            textBox13.Text = null;
            comboBox5.Text = null;

            if (dataGridView3.CurrentRow != null)
            {
                int rowindex = dataGridView3.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView3.Rows[rowindex];

                    textBox2.Text = row.Cells["ID"].Value.ToString();
                    textBox3.Text = row.Cells["原料編號"].Value.ToString();
                    textBox4.Text = row.Cells["添加量"].Value.ToString();
                    textBox13.Text = row.Cells["成品編號"].Value.ToString();
                    comboBox5.Text = row.Cells["原料名"].Value.ToString();

                }
                else
                {
                    textBox1.Text = null;
                    textBox2.Text = null;
                    textBox3.Text = null;
                    comboBox5.Text = null;
                }
            }
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox5.Text = null;
            comboBox5load(comboBox4.Text);
        }

        public void comboBox5load(string TYPE)
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);


            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT [TYPE],[MB001],[MB002] FROM [TKQC].[dbo].[NUTRITIONBASE] WHERE [TYPE]='{0}' ORDER BY [MB001],[MB002]",TYPE);

            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MB001", typeof(string));
            dt.Columns.Add("MB002", typeof(string));

            da.Fill(dt);
            comboBox5.DataSource = dt.DefaultView;
            comboBox5.ValueMember = "MB001";
            comboBox5.DisplayMember = "MB002";
            sqlConn.Close();
        }

        public void comboBox7load(string TYPE)
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);


            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT [TYPE],[MB001],[MB002] FROM [TKQC].[dbo].[NUTRITIONBASE] WHERE [TYPE]='{0}' ORDER BY [MB001],[MB002]", TYPE);

            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MB001", typeof(string));
            dt.Columns.Add("MB002", typeof(string));

            da.Fill(dt);
            comboBox7.DataSource = dt.DefaultView;
            comboBox7.ValueMember = "MB001";
            comboBox7.DisplayMember = "MB002";
            sqlConn.Close();
        }

        public void UPDATENUTRITIONPROD(string PRODID, string PRODNAME)
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);



                sqlConn.Close();
                sqlConn.Open();

                sbSql.Clear();
                sbSql.AppendFormat(@" 
                                    UPDATE  [TKQC].[dbo].[NUTRITIONPROD]
                                    SET [PRODID]='{0}',[PRODNAME]='{1}'
                                    WHERE [PRODID]='{0}'
                                                                      
                                        ", PRODID, PRODNAME);


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

        public void ADDNUTRITIONPROD(string PRODID,string PRODNAME)
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);



                sqlConn.Close();
                sqlConn.Open();

                sbSql.Clear();
                sbSql.AppendFormat(@" 
                                    INSERT INTO [TKQC].[dbo].[NUTRITIONPROD]
                                    ([PRODID],[PRODNAME])
                                    VALUES
                                    ('{0}','{1}')
                                                                      
                                        ", PRODID, PRODNAME);


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

        public void DELETENUTRITIONPROD(string PRODID)
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);



                sqlConn.Close();
                sqlConn.Open();

                sbSql.Clear();
                sbSql.AppendFormat(@" 
                                    DELETE [TKQC].[dbo].[NUTRITIONPROD]
                                    WHERE [PRODID]='{0}'
                                                                      
                                        ", PRODID);


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
        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox7.Text = null;
            comboBox7load(comboBox6.Text);
        }
        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (!string.IsNullOrEmpty(comboBox7.SelectedValue.ToString()))
                {
                    textBox11.Text = comboBox7.SelectedValue.ToString();
                    textBox12.Text = null;

                }
            }
            catch
            {

            }
           
           
        }

        public void ADDNUTRITIONPRODDETAIL(string ID, string PRODID, string PRODNAME, string MB001, string MB002,decimal USEDANOUNT)
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);



                ID = FINDMAXNUTRITIONPRODDETAILID();

                sqlConn.Close();
                sqlConn.Open();

                sbSql.Clear();
                sbSql.AppendFormat(@" 
                                    INSERT INTO [TKQC].[dbo].[NUTRITIONPRODDETAIL]
                                    ([ID],[PRODID],[PRODNAME],[MB001],[MB002],[USEDANOUNT])
                                    VALUES
                                    ('{0}','{1}','{2}','{3}','{4}',{5})
                                                                      
                                        ", ID, PRODID, PRODNAME, MB001, MB002, USEDANOUNT);


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


        public void UPDATENUTRITIONPRODDETAIL(string ID, string MB001, string MB002, decimal USEDANOUNT)
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);




                sqlConn.Close();
                sqlConn.Open();

                sbSql.Clear();
                sbSql.AppendFormat(@" 
                                   UPDATE [TKQC].[dbo].[NUTRITIONPRODDETAIL]
                                    SET [MB001]='{1}',[MB002]='{2}',[USEDANOUNT]='{3}'
                                    WHERE [ID]='{0}'
                                                                      
                                        ", ID,  MB001, MB002, USEDANOUNT);


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

        public void DELETENUTRITIONPRODDETAIL(string ID)
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);




                sqlConn.Close();
                sqlConn.Open();

                sbSql.Clear();
                sbSql.AppendFormat(@" 
                                   DELETE [TKQC].[dbo].[NUTRITIONPRODDETAIL]                                
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

        public void SETFASTREPORT(string PRODID)
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL1(PRODID);
            Report report1 = new Report();
            report1.Load(@"REPORT\營養計算.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;


            
            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL1(string PRODID)
        {
            StringBuilder SB = new StringBuilder();

           
            SB.AppendFormat(@"  
                            SELECT [NUTRITIONPRODDETAIL].[ID] AS '序號'
                            ,[NUTRITIONPROD].[PRODID] AS '成品編號'
                            ,[NUTRITIONPROD].[PRODNAME] AS '成品名'
                            ,[NUTRITIONPRODDETAIL].[MB001] AS '原料編號'
                            ,[NUTRITIONPRODDETAIL].[MB002] AS '原料名'
                            ,[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '添加量'
                            ,[NUTRITIONBASE].[CALORIES]*[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '熱量Kcal/100g'
                            ,[NUTRITIONBASE].[FAT]*[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '脂肪g/100g'
                            ,[NUTRITIONBASE].[SATURATEDFAT]*[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '飽和脂肪g/100g'
                            ,[NUTRITIONBASE].[TRANSFAT]*[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '反式脂肪g/100g'
                            ,[NUTRITIONBASE].[CHOLESTEROL]*[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '膽固醇mg/100g'
                            ,[NUTRITIONBASE].[SODIUM]*[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '鈉mg/100g'
                            ,[NUTRITIONBASE].[CARBOHYDRATES]*[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '碳水化合物g/100g'
                            ,[NUTRITIONBASE].[DIETARYFIBER]*[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '膳食纖維g/100g'
                            ,[NUTRITIONBASE].[SUGAR]*[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '糖g/100g'
                            ,[NUTRITIONBASE].[ADDSUGAR]*[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '添加糖g/100g'
                            ,[NUTRITIONBASE].[PROTEIN]*[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '蛋白質g/100g'
                            ,[NUTRITIONBASE].[VITANMIND]*[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '維生素D mcg/100g'
                            ,[NUTRITIONBASE].[CALCIUM]*[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '鈣 mg/100g'
                            ,[NUTRITIONBASE].[IRON]*[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '鐵mg/100g'
                            ,[NUTRITIONBASE].[POTASSIUM]*[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '鉀mg/100g'
                            FROM [TKQC].[dbo].[NUTRITIONPROD],[TKQC].[dbo].[NUTRITIONPRODDETAIL],[TKQC].[dbo].[NUTRITIONBASE] 
                            WHERE [NUTRITIONPROD].[PRODID]=[NUTRITIONPRODDETAIL].[PRODID]
                            AND [NUTRITIONPRODDETAIL].MB001=[NUTRITIONBASE].MB001
                            AND [NUTRITIONPROD].[PRODID]='{0}' 
                            ORDER BY [NUTRITIONPRODDETAIL].[MB001]
                            ", PRODID);



            return SB;

        }

        public void SETFASTREPORT2(string PRODID)
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL2(PRODID);
            Report report1 = new Report();
            report1.Load(@"REPORT\營養計算-台灣8大.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;



            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            report1.Preview = previewControl2;
            report1.Show();
        }

        public StringBuilder SETSQL2(string PRODID)
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@"  
                            SELECT [NUTRITIONPRODDETAIL].[ID] AS '序號'
                            ,[NUTRITIONPROD].[PRODID] AS '成品編號'
                            ,[NUTRITIONPROD].[PRODNAME] AS '成品名'
                            ,[NUTRITIONPRODDETAIL].[MB001] AS '原料編號'
                            ,[NUTRITIONPRODDETAIL].[MB002] AS '原料名'
                            ,[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '添加量'
                            ,[NUTRITIONBASE].[CALORIES]*[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '熱量Kcal/100g'
                            ,[NUTRITIONBASE].[FAT]*[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '脂肪g/100g'
                            ,[NUTRITIONBASE].[SATURATEDFAT]*[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '飽和脂肪g/100g'
                            ,[NUTRITIONBASE].[TRANSFAT]*[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '反式脂肪g/100g'
                            ,[NUTRITIONBASE].[CHOLESTEROL]*[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '膽固醇mg/100g'
                            ,[NUTRITIONBASE].[SODIUM]*[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '鈉mg/100g'
                            ,[NUTRITIONBASE].[CARBOHYDRATES]*[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '碳水化合物g/100g'
                            ,[NUTRITIONBASE].[DIETARYFIBER]*[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '膳食纖維g/100g'
                            ,[NUTRITIONBASE].[SUGAR]*[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '糖g/100g'
                            ,[NUTRITIONBASE].[ADDSUGAR]*[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '添加糖g/100g'
                            ,[NUTRITIONBASE].[PROTEIN]*[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '蛋白質g/100g'
                            ,[NUTRITIONBASE].[VITANMIND]*[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '維生素D mcg/100g'
                            ,[NUTRITIONBASE].[CALCIUM]*[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '鈣 mg/100g'
                            ,[NUTRITIONBASE].[IRON]*[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '鐵mg/100g'
                            ,[NUTRITIONBASE].[POTASSIUM]*[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '鉀mg/100g'
                            FROM [TKQC].[dbo].[NUTRITIONPROD],[TKQC].[dbo].[NUTRITIONPRODDETAIL],[TKQC].[dbo].[NUTRITIONBASE] 
                            WHERE [NUTRITIONPROD].[PRODID]=[NUTRITIONPRODDETAIL].[PRODID]
                            AND [NUTRITIONPRODDETAIL].MB001=[NUTRITIONBASE].MB001
                            AND [NUTRITIONPROD].[PRODID]='{0}' 
                            ORDER BY [NUTRITIONPRODDETAIL].[MB001]
                            ", PRODID);

                

            return SB;

        }

        public void SETFASTREPORT3(string PRODID)
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL3(PRODID);
            Report report1 = new Report();
            report1.Load(@"REPORT\營養計算-美規14.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;



            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            report1.Preview = previewControl3;
            report1.Show();
        }

        public StringBuilder SETSQL3(string PRODID)
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@"  
                            SELECT [NUTRITIONPRODDETAIL].[ID] AS '序號'
                            ,[NUTRITIONPROD].[PRODID] AS '成品編號'
                            ,[NUTRITIONPROD].[PRODNAME] AS '成品名'
                            ,[NUTRITIONPRODDETAIL].[MB001] AS '原料編號'
                            ,[NUTRITIONPRODDETAIL].[MB002] AS '原料名'
                            ,[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '添加量'
                            ,[NUTRITIONBASE].[CALORIES]*[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '熱量Kcal/100g'
                            ,[NUTRITIONBASE].[FAT]*[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '脂肪g/100g'
                            ,[NUTRITIONBASE].[SATURATEDFAT]*[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '飽和脂肪g/100g'
                            ,[NUTRITIONBASE].[TRANSFAT]*[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '反式脂肪g/100g'
                            ,[NUTRITIONBASE].[CHOLESTEROL]*[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '膽固醇mg/100g'
                            ,[NUTRITIONBASE].[SODIUM]*[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '鈉mg/100g'
                            ,[NUTRITIONBASE].[CARBOHYDRATES]*[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '碳水化合物g/100g'
                            ,[NUTRITIONBASE].[DIETARYFIBER]*[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '膳食纖維g/100g'
                            ,[NUTRITIONBASE].[SUGAR]*[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '糖g/100g'
                            ,[NUTRITIONBASE].[ADDSUGAR]*[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '添加糖g/100g'
                            ,[NUTRITIONBASE].[PROTEIN]*[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '蛋白質g/100g'
                            ,[NUTRITIONBASE].[VITANMIND]*[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '維生素D mcg/100g'
                            ,[NUTRITIONBASE].[CALCIUM]*[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '鈣 mg/100g'
                            ,[NUTRITIONBASE].[IRON]*[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '鐵mg/100g'
                            ,[NUTRITIONBASE].[POTASSIUM]*[NUTRITIONPRODDETAIL].[USEDANOUNT] AS '鉀mg/100g'
                            FROM [TKQC].[dbo].[NUTRITIONPROD],[TKQC].[dbo].[NUTRITIONPRODDETAIL],[TKQC].[dbo].[NUTRITIONBASE] 
                            WHERE [NUTRITIONPROD].[PRODID]=[NUTRITIONPRODDETAIL].[PRODID]
                            AND [NUTRITIONPRODDETAIL].MB001=[NUTRITIONBASE].MB001
                            AND [NUTRITIONPROD].[PRODID]='{0}' 
                            ORDER BY [NUTRITIONPRODDETAIL].[MB001]
                            ", PRODID);



            return SB;

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
            if(!string.IsNullOrEmpty(textBox311.Text))
            {
                ADDNUTRITIONBASE(ID, comboBox3.Text, textBox311.Text.Trim(), textBox312.Text, Convert.ToDecimal(textBox321.Text), Convert.ToDecimal(textBox322.Text), Convert.ToDecimal(textBox323.Text), Convert.ToDecimal(textBox324.Text), Convert.ToDecimal(textBox331.Text), Convert.ToDecimal(textBox332.Text), Convert.ToDecimal(textBox333.Text), Convert.ToDecimal(textBox334.Text), Convert.ToDecimal(textBox341.Text), Convert.ToDecimal(textBox342.Text), Convert.ToDecimal(textBox343.Text), Convert.ToDecimal(textBox344.Text), Convert.ToDecimal(textBox351.Text), Convert.ToDecimal(textBox352.Text), Convert.ToDecimal(textBox353.Text));

                SETTEXTBOXNULL2();

                SETTEXTBOXREADONLY4();
                SEARCHNUTRITIONBASE(comboBox1.Text.Trim());
            }
            else
            {
                MessageBox.Show("請填寫品號");
            }
            

        }
        private void button6_Click(object sender, EventArgs e)
        {
            SERACHNUTRITIONPROD(textBox1.Text.Trim());
        }


        private void button8_Click(object sender, EventArgs e)
        {
            UPDATENUTRITIONPRODDETAIL(textBox2.Text,textBox3.Text,comboBox5.Text,Convert.ToDecimal(textBox4.Text));
            SERACHNUTRITIONPRODDETAIL(textBox13.Text);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELETENUTRITIONPRODDETAIL(textBox2.Text);
                SERACHNUTRITIONPRODDETAIL(textBox13.Text);
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }
        private void button11_Click(object sender, EventArgs e)
        {
            if(!string.IsNullOrEmpty(textBox5.Text))
            {
                ADDNUTRITIONPROD(textBox5.Text.Trim(), textBox6.Text.Trim());
                SERACHNUTRITIONPROD(textBox1.Text.Trim());
                SETTEXTBOXNULL3();
            }
            else
            {
                MessageBox.Show("請填寫品號");
            }

        }
        private void button10_Click(object sender, EventArgs e)
        {
            UPDATENUTRITIONPROD(textBox7.Text.Trim(), textBox8.Text.Trim());
            SERACHNUTRITIONPROD(textBox1.Text.Trim());


        }
        private void button12_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELETENUTRITIONPROD(textBox7.Text.Trim());
                SERACHNUTRITIONPROD(textBox1.Text.Trim());
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
          
        }

        private void button14_Click(object sender, EventArgs e)
        {
            ADDNUTRITIONPRODDETAIL(ID,textBox9.Text,textBox10.Text, textBox11.Text,comboBox7.Text, Convert.ToDecimal(textBox12.Text));
            SERACHNUTRITIONPRODDETAIL(textBox9.Text);
        }
        private void button13_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(textBox14.Text);
            SETFASTREPORT2(textBox14.Text);
            SETFASTREPORT3(textBox14.Text);
        }


        #endregion


    }
}
