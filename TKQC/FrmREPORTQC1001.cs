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
    public partial class FrmREPORTQC1001 : Form
    {
        public FrmREPORTQC1001()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SETFASTREPORT(string SDAY, string EDAY)
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL1(SDAY,EDAY);
            Report report1 = new Report();
            report1.Load(@"REPORT\1001.客訴品質異常處理單.frx");

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

        public StringBuilder SETSQL1(string SDAY,string EDAY)
        {
            StringBuilder SB = new StringBuilder();

           
            SB.AppendFormat(@"  

                            SELECT 
                             [QCFrm001SN] AS '表單編號'
                            ,[QCFrm001ASN] AS 'A單編號'
                            ,[QCFrm001Date] AS '日期'
                            ,[QCFrm001User] AS '申請者'
                            ,[QCFrm001Dept] AS '部門'
                            ,[QCFrm001Rank] AS '職級'
                            ,[QCFrm001CUST] AS '供應商/部門單位'
                            ,[QCFrm001PNO] AS '批號'
                            ,[QCFrm001CN] AS '品號'
                            ,[QCFrm001PRD] AS '品名'
                            ,[QCFrm001RDate] AS '受理日期'
                            ,[QCFrm001MD] AS '製造日期'
                            ,[QCFrm001ND] AS '有效日期'
                            ,[QCFrm002Cmf] AS '品保初判'
                            ,[QCFrm002Abn] AS '客訴原因詳述'
                            ,[QCFrm002Abns] AS '客訴原因說明'
                            ,[QCFrm001Range] AS '相關單位'
                            ,[QCFrm001HB] AS '知會相關人員'
                            ,[QCFrm001RCA] AS '原因分析'
                            ,[QCFrm001RCAU] AS '原因分析填寫人'
                            ,[QCFrm001PA] AS '預防對策'
                            ,[QCFrm001QA] AS '品保審核'
                            ,[QCFrm001PA2] AS '二次預防對策'
                            ,[QCFrm001QA2] AS '二次品保審核'
                            ,[QCFrm001PA3] AS '三次預防對策'
                            ,[QCFrm001QA3] AS '三次品保審核'
                            ,[QCFrm001Cmf] AS '品保追蹤'
                            ,[QCFrm001Comp] AS '完成日期'
                            ,[QCFrm001Cmf1] AS '確認結果'
                            FROM [TKQC].[dbo].[TBUOFQC1001]
                            WHERE [QCFrm001Date]>='{0}' AND [QCFrm001Date]<='{1}'
                            ORDER BY [QCFrm001SN]
                            ", SDAY,EDAY);


            return SB;

        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(dateTimePicker1.Value.ToString("yyyy/MM/dd"),dateTimePicker2.Value.ToString("yyyy/MM/dd"));
        }
        #endregion
    }
}
