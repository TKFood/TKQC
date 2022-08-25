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
    public partial class FrmQCPURTGPURTH : Form
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



        public FrmQCPURTGPURTH()
        {
            InitializeComponent();
        }

        #region FUNCTION    
        public void SETFASTREPORT(string SDATES, string EDATES, string NO)
        {
            StringBuilder SQL1 = new StringBuilder();
            StringBuilder SQLQUERY = new StringBuilder();

            if (!string.IsNullOrEmpty(NO))
            {
                SQLQUERY.AppendFormat(@"  AND [TG002] LIKE '%{0}%' ", NO);
            }
            else
            {
                SQLQUERY.AppendFormat(@" ");
            }

            SQL1.AppendFormat(@"  
                             
                            
                            SELECT 
                             [ID] AS 'UOF表單編號'
                            ,[TG003] AS '進貨日期'
                            ,[TG005] AS '廠商'
                            ,[TG021] AS '廠商名稱'
                            ,[TG001] AS '單別('
                            ,[TG002] AS '單號'
                            ,[TH003] AS '序號'
                            ,[TH004] AS '品號'
                            ,[TH005] AS '品名'
                            ,[TH006] AS '規格'
                            ,[TH007] AS '進貨數量'
                            ,[TH008] AS '單位'
                            ,[TH010] AS '批號'
                            ,[TH015] AS '驗收數量'
                            ,[CHECK] AS '是否驗收'
                            ,[DETAIL01] AS '抽樣數量'
                            ,[DETAIL02] AS '檢驗項目'
                            ,[DETAIL03] AS '提供COA'
                            ,[DETAIL04] AS '內部檢驗'
                            ,[DETAIL05] AS '驗退數量'
                            ,[DETAIL06] AS '備註'
                            FROM [TKQC].[dbo].[UOFQCPURTGPURTH]
                            WHERE [TG003]>='{0}' AND [TG003]<='{1}'
                            {2}

                            ORDER BY TG001,TG002,TH003
                                
                                ", SDATES, EDATES, SQLQUERY.ToString());

            Report report1 = new Report();
            report1.Load(@"REPORT\原物料品質驗收單UOF.frx");

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
        #endregion

        #region BUTTON
        private void button4_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"), textBox1.Text.Trim());
        }
        #endregion
    }
}
