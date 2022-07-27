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

                                    [TH001]
                                    ,[TH002]
                                    ,[TH003]
                                    ,[TG003]
                                    ,[TG005]
                                    ,[TG021]
                                    ,[TH004]
                                    ,[TH005]
                                    ,[TH006]
                                    ,[TH007]
                                    ,[TH008]
                                    ,[TH009]
                                    ,[SAMPLENUMS]
                                    ,[CARNO]
                                    ,[CHECKITEMS]
                                    ,[COA]
                                    ,[INNERCHECKS]
                                    ,[INUMS]
                                    ,[BACKNUMS]
                                    ,[DATES]
                                    ,[QCMAN]
                                    ,[COMMENTS]
                                    ,[ISIN]


                                    FROM [TKQC].[dbo].[QCPURTH]

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
        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            Search(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
        }

        #endregion
    }
}
