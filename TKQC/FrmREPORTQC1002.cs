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

        private void FrmREPORTQC1002_Load(object sender, EventArgs e)
        {
            DateTime today = DateTime.Today;
            // 當月的第一天
            DateTime firstDayOfMonth = new DateTime(today.Year, today.Month, 1);
            // 當月的最後一天
            DateTime lastDayOfMonth = firstDayOfMonth.AddMonths(1).AddDays(-1);

            dateTimePicker1.Value = firstDayOfMonth;
            dateTimePicker3.Value = firstDayOfMonth;
            dateTimePicker2.Value = lastDayOfMonth;
            dateTimePicker4.Value = lastDayOfMonth; 
        }
        public void SETFASTREPORT(string SDATES, string EDATES)
        {


            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL2(SDATES, EDATES);
            Report report1 = new Report();
            report1.Load(@"REPORT\1002.客訴異常處理單明細-橫式.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

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

      

        public void SETFASTREPORT2(string SDATES, string EDATES)
        {


            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL2(SDATES, EDATES); 
            Report report1 = new Report();
            report1.Load(@"REPORT\1002.客訴異常處理單明細.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;


            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            //report1.SetParameterValue("P1", P1);


            report1.Preview = previewControl2;
            report1.Show();
        }

        public StringBuilder SETSQL2(string SDATES, string EDATES)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@"                            
                           --20250528 查客訴單明細
                            SELECT 
                            DOC_NBR AS '客訴單號'
                            ,CURRENT_DOC.value('(Form/FormFieldValue/FieldItem[@fieldId=""QCFrm002Abns""]/@fieldValue)[1]', 'nvarchar(max)')+' '+CURRENT_DOC.value('(Form/FormFieldValue/FieldItem[@fieldId=""QCFrm002Abns""]/@customValue)[1]', 'nvarchar(max)') AS '原因'
                            , CURRENT_DOC.value('(Form/FormFieldValue/FieldItem[@fieldId=""QCFrm002RDate""]/@fieldValue)[1]', 'nvarchar(max)') AS '受理日期'
                            , CURRENT_DOC.value('(Form/FormFieldValue/FieldItem[@fieldId=""QCFrm002CUST""]/@fieldValue)[1]', 'nvarchar(max)') AS '客戶'
                            , CURRENT_DOC.value('(Form/FormFieldValue/FieldItem[@fieldId=""QCFrm002PRD""]/@fieldValue)[1]', 'nvarchar(max)') AS '產品'
                            , CURRENT_DOC.value('(Form/FormFieldValue/FieldItem[@fieldId=""QCFrm002ED""]/@fieldValue)[1]', 'nvarchar(max)') AS '有效日'
                            , CURRENT_DOC.value('(Form/FormFieldValue/FieldItem[@fieldId=""QCFrm002MD""]/@fieldValue)[1]', 'nvarchar(max)') AS '製造日'
                            , CURRENT_DOC.value('(Form/FormFieldValue/FieldItem[@fieldId=""QCFrm002Abn""]/@fieldValue)[1]', 'nvarchar(max)') AS '客訴原因詳述'
                            , CURRENT_DOC.value('(Form/FormFieldValue/FieldItem[@fieldId=""QCFrm002Process""]/@fieldValue)[1]', 'nvarchar(max)') AS '回覆內容'

                            , TB_WKF_FORM.FORM_NAME
                            , (SELECT TOP 1 NAME FROM[UOF].dbo.TB_EB_USER WHERE TB_EB_USER.USER_GUID = TB_WKF_TASK.USER_GUID) AS 'NAMES'
                            ,CURRENT_DOC
                            FROM[UOF].dbo.TB_WKF_TASK,[UOF].dbo.TB_WKF_FORM,[UOF].dbo.TB_WKF_FORM_VERSION
                            WHERE 1 = 1
                            AND TB_WKF_TASK.FORM_VERSION_ID = TB_WKF_FORM_VERSION.FORM_VERSION_ID
                            AND TB_WKF_FORM.FORM_ID = TB_WKF_FORM_VERSION.FORM_ID
                            AND TB_WKF_FORM.FORM_NAME IN('1002.客訴異常處理單')
                            AND ISNULL(TB_WKF_TASK.TASK_RESULT,'') NOT IN('2')
                            AND CONVERT(NVARCHAR, TB_WKF_TASK.BEGIN_TIME,112)>='{0}' AND TB_WKF_TASK.BEGIN_TIME<='{1}'
                            ORDER BY DOC_NBR
 
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

        private void button2_Click(object sender, EventArgs e)
        {
            SETFASTREPORT2(dateTimePicker3.Value.ToString("yyyyMMdd"), dateTimePicker4.Value.ToString("yyyyMMdd"));
        }

    
    }
}
