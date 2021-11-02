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
using TKITDLL;
namespace TKQC
{
    public partial class frmQCWH : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        DataSet ds = new DataSet();
        SqlTransaction tran;
        int result;
        SqlCommand cmd = new SqlCommand();

        public frmQCWH()
        {
            InitializeComponent();
        }

        #region FUNCTION

        public void SEARCH(string ISCLOSED)
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
                                    [ID]
                                    ,[LA004] AS '轉入日'
                                    ,[LA001] AS '品號'
                                    ,[MB002] AS '品名'
                                    ,[LA016] AS '批號'
                                    ,[LA011] AS '轉入數量'
                                    ,[LA006] AS '單別'
                                    ,[LA007] AS '單號'
                                    ,[LA008] AS '序號'
                                    ,[COMMENTS] AS '記錄'
                                    ,[ISCLOSED] AS '是否結案'
                                    FROM [TKQC].[dbo].[TBQCWAREHOUSE]
                                    WHERE [ISCLOSED]='{0}'

                                    ", ISCLOSED);

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
                    textBox2.Text = row.Cells["ID"].Value.ToString().Trim();
                    textBox1.Text = row.Cells["記錄"].Value.ToString().Trim();
                    comboBox2.Text = row.Cells["是否結案"].Value.ToString().Trim();



                }
                else
                {
                    textBox2.Text = null;
                    textBox1.Text = null;
                }
            }
        }

        public void UPDATETBQCWAREHOUSE(string ID,string COMMENTS,string ISCLOSED)
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
                                    UPDATE  [TKQC].[dbo].[TBQCWAREHOUSE]
                                    SET [COMMENTS]='{1}'
                                    ,[ISCLOSED]='{2}'
                                    WHERE [ID]='{0}'
                                    ",ID,COMMENTS,ISCLOSED);



                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消

                    MessageBox.Show("更新");

                }
                else
                {
                    tran.Commit();      //執行交易                    

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

        public void ADDTBQCWAREHOUSE(string SDAYS,string EDAYS)
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
                                  
                                    INSERT INTO [TKQC].[dbo].[TBQCWAREHOUSE]
                                    ([LA004],[LA001],[MB002],[LA016],[LA011],[LA006],[LA007],[LA008],[COMMENTS],[ISCLOSED])
                                    SELECT LA004,LA001,MB002,LA016,LA011,LA006,LA007,LA008,'','N'
                                    FROM [TK].dbo.INVLA,[TK].dbo.INVMB
                                    WHERE LA001=MB001
                                    AND LA005='1'
                                    AND LA009='20007'
                                    AND LA004>='{0}' AND LA004<='{1}'
                                    ",SDAYS,EDAYS);



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

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCH(comboBox1.Text.ToString());
        }
        #endregion

        private void button2_Click(object sender, EventArgs e)
        {
            ADDTBQCWAREHOUSE(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
            SEARCH(comboBox1.Text.ToString());
        }

        private void button3_Click(object sender, EventArgs e)
        {
            UPDATETBQCWAREHOUSE(textBox2.Text,textBox1.Text,comboBox2.Text);
            SEARCH(comboBox1.Text.ToString());
        }

       
    }
}
