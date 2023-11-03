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
    public partial class FrmREPORTQC1006 : Form
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

        public FrmREPORTQC1006()
        {
            InitializeComponent();
        }

        #region FUNCTION       

        public void SETFASTREPORT(string SDATES, string EDATES)
        {


            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL(SDATES, EDATES);
            Report report1 = new Report();
            report1.Load(@"REPORT\1006委外送驗申請單.frx");

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

            //report1.SetParameterValue("P1", P1);


            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL(string SDATES, string EDATES)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@"                            
                           
                            SELECT 
                            [DOC_NBR] AS 'DOC_NBR'
                            ,[QCFrm004SN] AS '表單編號'
                            ,[QCFrm004Date] AS '申請日期'
                            ,[QCFrm004UserLevel] AS '申請者職級'
                            ,[QC6001] AS '交付人'
                            ,[QC6002] AS '交付人部門'
                            ,[QC6012] AS '接辦人'
                            ,[QC6003] AS '交辦原因'
                            ,[QC6004] AS '案件類別'
                            ,[QC6005] AS '特殊需求'
                            ,[QC6006] AS '指定單位'
                            ,[QC6007] AS '檢驗需求'
                            ,[QC6008] AS '品保寄件日期'
                            ,[QC6009] AS '檢驗單位報價'
                            ,[QC60010] AS '報告預計交期'
                            FROM [TKQC].[dbo].[TBUOFQC1006]
                            WHERE [QCFrm004Date]>='{0}' AND [QCFrm004Date]<='{1}'
                            ", SDATES, EDATES);

            return SB;

        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(dateTimePicker1.Value.ToString("yyyy/MM/dd"), dateTimePicker2.Value.ToString("yyyy/MM/dd"));
        }
        #endregion
    }
}
