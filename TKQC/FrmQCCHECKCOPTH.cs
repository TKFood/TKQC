using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI;
using NPOI.HPSF;
using NPOI.HSSF;
using NPOI.HSSF.UserModel;
using NPOI.POIFS;
using NPOI.Util;
using NPOI.HSSF.Util;
using NPOI.HSSF.Extractor;
using System.IO;
using System.Data.SqlClient;
using NPOI.SS.UserModel;
using System.Configuration;
using NPOI.XSSF.UserModel;
using FastReport;
using FastReport.Data;
using System.Net.Mail;
using TKITDLL;

namespace TKQC
{
    public partial class FrmQCCHECKCOPTH : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlCommand cmd = new SqlCommand();
        SqlTransaction tran;
        DataSet ds = new DataSet();
        DataTable dt = new DataTable();
        string talbename = null;
        int rownum = 0;
        int result;

        int ROWSINDEX = 0;
        int COLUMNSINDEX = 0;
        public FrmQCCHECKCOPTH()
        {
            InitializeComponent();
        }

        #region FUNCTION       
        public void Search(string SDATES,string EDATES)
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


                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  
                                    SELECT 
                                   
                                    [ISIN] AS '是否允收'
                                    ,[TG003] AS '進貨日期'
                                    ,[TG005] AS '廠商代'
                                    ,[TG021] AS '廠商名稱'
                                    ,[TH004] AS '品號'
                                    ,[TH005] AS '品名'
                                    ,[TH006] AS '規格'
                                    ,[TH007] AS '進貨數量'
                                    ,[TH008] AS '單位'
                                    ,[TH009] AS '庫別'
                                    ,[SAMPLENUMS] AS '抽樣數量'
                                    ,[CARNO] AS '運輸車'
                                    ,[CHECKITEMS] AS '檢驗項目'
                                    ,[COA] AS '提供COA'
                                    ,[INNERCHECKS] AS '內部檢驗'
                                    ,[INUMS] AS '合格數量'
                                    ,[BACKNUMS] AS '退貨數量'
                                    ,[DATES] AS '日期'
                                    ,[QCMAN] AS '驗收人員'
                                    ,[COMMENTS] AS '備註'
                                    ,[TH001]
                                    ,[TH002]
                                    ,[TH003]

                                    FROM [TKQC].[dbo].[QCPURTH]
                                    WHERE [TG003]>='{0}' AND [TG003]<='{1}'

                                    ", SDATES, EDATES);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds");
                sqlConn.Close();

                if (ds.Tables["TEMPds"].Rows.Count >= 1)
                {
                    //dataGridView1.Rows.Clear();
                    dataGridView1.DataSource = ds.Tables["TEMPds"];
                    dataGridView1.AutoResizeColumns();
                    //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                    if (ROWSINDEX > 0 || COLUMNSINDEX > 0)
                    {
                        dataGridView1.CurrentCell = dataGridView1.Rows[ROWSINDEX].Cells[COLUMNSINDEX];

                        DataGridViewRow row = dataGridView1.Rows[ROWSINDEX];
                        textBox1.Text = row.Cells["TH001"].Value.ToString() + row.Cells["TH002"].Value.ToString() + row.Cells["TH003"].Value.ToString();
                        textBox2.Text = row.Cells["品名"].Value.ToString();
                        textBox3.Text = row.Cells["進貨數量"].Value.ToString();
                        textBox4.Text = row.Cells["是否允收"].Value.ToString();
                        textBox5.Text = row.Cells["運輸車"].Value.ToString();
                        textBox6.Text = row.Cells["提供COA"].Value.ToString();
                        textBox7.Text = row.Cells["檢驗項目"].Value.ToString();
                        textBox8.Text = row.Cells["內部檢驗"].Value.ToString();
                        textBox9.Text = row.Cells["抽樣數量"].Value.ToString();
                        textBox10.Text = row.Cells["合格數量"].Value.ToString();
                        textBox11.Text = row.Cells["退貨數量"].Value.ToString();
                        textBox12.Text = row.Cells["日期"].Value.ToString();
                        textBox13.Text = row.Cells["驗收人員"].Value.ToString();
                        textBox14.Text = row.Cells["備註"].Value.ToString();

                    }
                }
                else
                {
                    dataGridView1.DataSource = null;
                }

            }
            catch
            {

            }
            finally
            {

            }
        }
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    textBox1.Text = row.Cells["TH001"].Value.ToString()+ row.Cells["TH002"].Value.ToString()+ row.Cells["TH003"].Value.ToString();
                    textBox2.Text = row.Cells["品名"].Value.ToString();
                    textBox3.Text = row.Cells["進貨數量"].Value.ToString();
                    textBox4.Text = row.Cells["是否允收"].Value.ToString();
                    textBox5.Text = row.Cells["運輸車"].Value.ToString();
                    textBox6.Text = row.Cells["提供COA"].Value.ToString();
                    textBox7.Text = row.Cells["檢驗項目"].Value.ToString();
                    textBox8.Text = row.Cells["內部檢驗"].Value.ToString();
                    textBox9.Text = row.Cells["抽樣數量"].Value.ToString();
                    textBox10.Text = row.Cells["合格數量"].Value.ToString();
                    textBox11.Text = row.Cells["退貨數量"].Value.ToString();
                    textBox12.Text = row.Cells["日期"].Value.ToString();
                    textBox13.Text = row.Cells["驗收人員"].Value.ToString();
                    textBox14.Text = row.Cells["備註"].Value.ToString();


                    ROWSINDEX = dataGridView1.CurrentCell.RowIndex;
                    COLUMNSINDEX = dataGridView1.CurrentCell.ColumnIndex;

                    rowindex = ROWSINDEX;

                    ;
                }
                else
                {
                    textBox1.Text = "";
                    textBox2.Text = "";
                    textBox3.Text = "";
                    textBox4.Text = "";
                    textBox5.Text = "";
                    textBox6.Text = "";
                    textBox7.Text = "";
                    textBox8.Text = "";
                    textBox9.Text = ""; 
                    textBox10.Text = "";
                    textBox11.Text = "";
                    textBox12.Text = "";
                    textBox13.Text = "";
                    textBox14.Text = "";

                }
            }
        }

        public void UPDATEQCPURTH(string TH001TH002TH003, string ISIN, string SAMPLENUMS, string CARNO, string CHECKITEMS, string COA, string INNERCHECKS, string INUMS, string BACKNUMS, string DATES, string QCMAN, string COMMENTS)
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
                                    UPDATE [TKQC].[dbo].[QCPURTH]
                                    SET [ISIN]='{1}'
                                    ,[SAMPLENUMS]='{2}'
                                    ,[CARNO]='{3}'
                                    ,[CHECKITEMS]='{4}'
                                    ,[COA]='{5}'
                                    ,[INNERCHECKS]='{6}'
                                    ,[INUMS]='{7}'
                                    ,[BACKNUMS]='{8}'
                                    ,[DATES]='{9}'
                                    ,[QCMAN]='{10}'
                                    ,[COMMENTS]='{11}'
                                    WHERE [TH001]+[TH002]+[TH003]='{0}'
                                        ", TH001TH002TH003
                                        , ISIN
                                        , SAMPLENUMS
                                        , CARNO
                                        , CHECKITEMS
                                        , COA
                                        , INNERCHECKS
                                        , INUMS
                                        , BACKNUMS
                                        , DATES
                                        , QCMAN
                                        , COMMENTS);



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
                                        //UPDATEMOCMANULINETEMP(NEWGUID, TEMPds);

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


        public void SETFASTREPORT(string SDATES,string EDATES)
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1.AppendFormat(@"  
                             
                                SELECT 
                                [ISIN] AS '是否允收'
                                ,[TG003] AS '進貨日期'
                                ,[TG005] AS '廠商代'
                                ,[TG021] AS '廠商名稱'
                                ,[TH004] AS '品號'
                                ,[TH005] AS '品名'
                                ,[TH006] AS '規格'
                                ,[TH007] AS '進貨數量'
                                ,[TH008] AS '單位'
                                ,[TH009] AS '庫別'
                                ,[SAMPLENUMS] AS '抽樣數量'
                                ,[CARNO] AS '運輸車'
                                ,[CHECKITEMS] AS '檢驗項目'
                                ,[COA] AS '提供COA'
                                ,[INNERCHECKS] AS '內部檢驗'
                                ,[INUMS] AS '合格數量'
                                ,[BACKNUMS] AS '退貨數量'
                                ,[DATES] AS '日期'
                                ,[QCMAN] AS '驗收人員'
                                ,[COMMENTS] AS '備註'
                                ,[TH001]
                                ,[TH002]
                                ,[TH003]


                                FROM [TKQC].[dbo].[QCPURTH]
                                WHERE [TG003]>='{0}' AND [TG003]<='{1}'
                                
                                ", SDATES, EDATES);

            Report report1 = new Report();
            report1.Load(@"REPORT\原物料品質驗收單.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            SqlConnection sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;


            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            report1.Preview = previewControl1;
            report1.Show();
        }

        public void SETFASTREPORT2(string SDATES, string EDATES)
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1.AppendFormat(@"  
                             
                               SELECT 
                                MA002 AS '供應廠商'
                                ,TD001 AS '單別'
                                ,TD002 AS '單號'
                                ,TD003 AS '序號'
                                ,TD004 AS '品號'
                                ,TD005 AS '品名'
                                ,TD006 AS '規格'
                                ,TD008 AS '採購數量'
                                ,TD009 AS '單位'
                                ,TD012 AS '預交日'
                                FROM [TK].dbo.PURTC,[TK].dbo.PURTD,[TK].dbo.PURMA
                                WHERE TC001=TD001 AND TC002=TD002
                                AND MA001=TC004
                                AND TD012>='{0}' AND TD012<='{1}'
                                ORDER BY TD012
                                
                                ", SDATES, EDATES);

            Report report1 = new Report();
            report1.Load(@"REPORT\採購預計到貨表.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            SqlConnection sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;


            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            report1.Preview = previewControl2;
            report1.Show();
        }


        public void UPDATETKPURTH()
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
                                    UPDATE [TK].dbo.PURTH
                                    SET UDF01='Y'
                                    WHERE UDF01<>'Y'
                                    AND [TH001]+[TH002]+[TH003] IN (SELECT [TH001]+[TH002]+[TH003] FROM [TKQC].[dbo].[QCPURTH] WHERE [ISIN]='Y')

                                    
                                    UPDATE [TK].dbo.PURTH
                                    SET UDF01=''
                                    WHERE ISNULL(UDF01,'')<>''
                                    AND [TH001]+[TH002]+[TH003] IN (SELECT [TH001]+[TH002]+[TH003] FROM [TKQC].[dbo].[QCPURTH] WHERE [ISIN]='N')


                                        ");



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
                                        //UPDATEMOCMANULINETEMP(NEWGUID, TEMPds);

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
            Search(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
        }

        private void button2_Click(object sender, EventArgs e)
        {
            UPDATEQCPURTH(textBox1.Text.Trim(), textBox4.Text.Trim(), textBox9.Text.Trim(), textBox5.Text.Trim(), textBox7.Text.Trim(), textBox6.Text.Trim(), textBox8.Text.Trim(), textBox10.Text.Trim(), textBox11.Text.Trim(), textBox12.Text.Trim(), textBox13.Text.Trim(), textBox14.Text.Trim());

            //更新ERP進貨單單身，是否經品保檢驗 
            UPDATETKPURTH();

            Search(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
        }
        private void button3_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(dateTimePicker3.Value.ToString("yyyyMMdd"),dateTimePicker4.Value.ToString("yyyyMMdd"));
        }
        private void button4_Click(object sender, EventArgs e)
        {
            SETFASTREPORT2(dateTimePicker5.Value.ToString("yyyyMMdd"), dateTimePicker6.Value.ToString("yyyyMMdd"));
        }
        #endregion


    }
}
