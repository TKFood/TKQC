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
    public partial class FrmREPORTQC1002 : Form
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

        public FrmREPORTQC1002()
        {
            InitializeComponent();
        }

        #region FUNCTION       

        public void SETFASTREPORT(string SDATES, string EDATES)
        {


            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL(SDATES, EDATES);
            Report report1 = new Report();
            report1.Load(@"REPORT\1002.客訴異常處理單報表.frx");

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
                             [QCFrm002SN] AS '表單編號'
                            ,[QCFrm002Date] AS '申請日期'
                            ,[QCFrm002User] AS '申請者'
                            ,[QCFrm002Dept] AS '部門'
                            ,[QCFrm002Rank] AS '職級'
                            ,[QCFrm002CUST] AS '消費者姓名'
                            ,[QCFrm002TEL] AS '消費者電話'
                            ,[QCFrm002Add] AS '消費者地址'
                            ,[QCFrm002CU] AS '供應商/部門單位'
                            ,[QCFrm002PNO] AS '批號'
                            ,[QCFrm002CN] AS '品號'
                            ,[QCFrm002RDate] AS '受理日期'
                            ,[QCFrm002PRD] AS '品名'
                            ,[QCFrm002PKG] AS '包裝形式(單片包/罐裝)及規格'
                            ,[QCFrm002MD] AS '製造日期'
                            ,[QCFrm002ED] AS '有效日期'
                            ,[QCFrm002OD] AS '購買日期'
                            ,[QCFrm002BP] AS '購買地點'
                            ,[QCFrm002Prove] AS '購買證明'
                            ,[QCFrm002Abns] AS '客訴原因說明'
                            ,[QCFrm002Range] AS '使用範圍'
                            ,[QCFrm002RP] AS '客訴來源'
                            ,[QCFrm002RD] AS '產品預計回收日'
                            ,[QCFrm002Abn] AS '客訴原因詳述'
                            ,[QCFrm002Process] AS '業務處理方式'
                            ,[QCFrm002QCR] AS '品保建議回覆內容'
                            ,[QCFrm002ProcessR] AS '業務對外回覆'
                            ,[QCFrm002QCC] AS '品保判定'
                            ,[QCFrm002RCAU] AS '判定人員'
                            ,[QCFrm002PRRD] AS '實際產品回收日期'
                            ,[QCFrm002Cmf] AS '品保初判'
                            ,[QCFrm002False] AS '知會人員'
                            ,[REPORTS] AS '說明'
                            FROM [TKQC].[dbo].[TBUOFQC1002] WITH(NOLOCK)
                            WHERE [QCFrm002Date]>='{0}' AND [QCFrm002Date]<='{1}'
                            ORDER BY [QCFrm002Date]
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
