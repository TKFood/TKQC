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
    public partial class FrmINVMBBARCODE : Form
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

        public FrmINVMBBARCODE()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SERACHNUTRITIONPROD(string MB001,string MB013)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            StringBuilder sbSqLQUERY1 = new StringBuilder();
            StringBuilder sbSqLQUERY2 = new StringBuilder();

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

                if (!string.IsNullOrEmpty(MB001))
                {
                    sbSqLQUERY1.AppendFormat(@"
                                        AND MB001 LIKE '%{0}%'
                                    ", MB001);
                }
                else
                {
                    sbSqLQUERY1.AppendFormat(@"
                                      
                                    ");
                }
                if (!string.IsNullOrEmpty(MB013))
                {
                    sbSqLQUERY2.AppendFormat(@"
                                        AND MB013 LIKE '%{0}%'
                                    ", MB013);
                }
                else
                {
                    sbSqLQUERY2.AppendFormat(@"
                                      
                                    ");
                }

                sbSql.AppendFormat(@" 
                                    SELECT MB001 AS '品號',MB002 AS '品名',MB003 AS '規格',MB013 AS '條碼'
                                    FROM [TK].dbo.INVMB
                                    WHERE 1=1
                                    {0}
                                    {1}
                                    ORDER BY MB001
                                    ", sbSqLQUERY1.ToString(), sbSqLQUERY2.ToString());


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

                        dataGridView1.Columns["品號"].Width = 200;
                        dataGridView1.Columns["品名"].Width = 200;
                        dataGridView1.Columns["規格"].Width = 200;
                        dataGridView1.Columns["條碼"].Width = 200;

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
            SERACHNUTRITIONPROD(textBox1.Text,textBox2.Text);
        }


        #endregion


    }
}
