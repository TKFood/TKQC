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
    public partial class FrmREPORTSALEBACK : Form
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

        public FrmREPORTSALEBACK()
        {
            InitializeComponent();
        }


        #region FUNCTION       

        public void SETFASTREPORT(string SDATES, string EDATES)
        {


            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL(SDATES, EDATES);
            Report report1 = new Report();
            report1.Load(@"REPORT\寄倉銷貨退回報表.frx");

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
                            SELECT *
                            FROM (
                            SELECT TA001 AS '單別',TA002 AS '單號',TA003 AS '異動日期',TA004 AS '部門代號',TA005 AS '單頭備註'
                            ,TB003 AS '序號',TB004 AS '品號',TB005 AS '品名',TB006 AS '規格',TB007 AS '數量'
                            ,TB008 AS '單位',TB012 AS '轉出庫',MC1.MC002 AS '轉出',TB013 AS '轉入庫',MC2.MC002 AS '轉入',TB014 AS '批號',TB015 AS '有效日期',TB016 AS '複檢日期',TA005+' '+TB017 AS '原因備註'
                            FROM [TK].dbo.INVTA,[TK].dbo.INVTB
                            LEFT JOIN [TK].dbo.CMSMC MC1 ON MC1.MC001=TB012
                            LEFT JOIN [TK].dbo.CMSMC MC2 ON MC2.MC001=TB013
                            WHERE TA001=TB001 AND TA002=TB002
                            AND TA006='Y'
                            AND TA001 IN ('A130')
                            AND TA003>='{0}' AND TA003<='{1}'
                            UNION ALL
                            SELECT TI001 AS '單別',TI002 AS '單號',TI003 AS '異動日期',TI005 AS '部門代號',TI021+' '+TI020 AS '單頭備註'
                            ,TJ003 AS '序號',TJ004 AS '品號',TJ005 AS '品名',TJ006 AS '規格',TJ007 AS '數量'
                            ,TJ008 AS '單位','' AS '轉出庫','' AS '轉出',TJ013 AS '轉入庫',MC002 AS '轉入',TJ014 AS '批號',TJ096 AS '有效日期',TJ057 AS '複檢日期',TI021+' '+TI020+' '+TJ023 AS '原因備註'
                            FROM [TK].dbo.COPTI,[TK].dbo.COPTJ
                            LEFT JOIN [TK].dbo.CMSMC MC1 ON MC1.MC001=TJ013
                            WHERE TI001=TJ001 AND TI002=TJ002
                            AND TI019='Y'
                            AND (TJ004 LIKE '4%' OR TJ004 LIKE '5%' ) 
                            AND TI003>='{0}' AND TI003<='{1}'
                            ) AS TEMP 
                            ORDER BY 單別,單號,序號

                            ", SDATES, EDATES);

            return SB;

        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
        }
        #endregion
    }
}
