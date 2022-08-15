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
    public partial class FrnREPORTBACKS : Form
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


        public FrnREPORTBACKS()
        {
            InitializeComponent();
        }

        #region FUNCTION     

        public void SETFASTREPORT(string SDATES, string EDATES)
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1.AppendFormat(@"  
                             
                            SELECT TB001 AS '單別',TB002 AS '單號',TB003 AS '序號',TB004 AS '品號',TB005 AS '品名',TB006 AS '規格',TB007 AS '數量',TB008 AS '單位',TB014 AS '批號',TA005 AS '單頭備註',TB017 AS '單身備註'
                            FROM [TK].dbo.INVTA,[TK].dbo.INVTB
                            WHERE TA001=TB001 AND TA002=TB002
                            AND TA001 IN ('A122','A123','A130')
                            AND TB013='20007'
                            AND TA003>='{0}' AND TA003<='{1}'
                            UNION ALL
                            SELECT TI001,TI002,TI003,TI004,TI005,TI006,TI009,TI010,TI017,TH014,TI021
                            FROM [TK].dbo.INVTH,[TK].dbo.INVTI
                            WHERE TH001=TI001 AND TH002=TI002
                            AND TH001 IN ('A151')
                            AND TI008='20007'
                            AND TH003>='{0}' AND TH003<='{1}'
                            UNION ALL
                            SELECT TJ001,TJ002,TJ003,TJ004,TJ005,TJ006,TJ007,TJ008,TJ014,TI020,TJ023
                            FROM [TK].dbo.COPTI,[TK].dbo.COPTJ
                            WHERE TI001=TJ001 AND TI002=TJ002
                            AND TJ013='20007'
                            AND TI003>='{0}' AND TI003<='{1}'
                            UNION ALL
                            SELECT TJ001,TJ002,TJ003,TJ004,TJ005,TJ006,TJ009,TJ007,TJ012,TI012,TJ019
                            FROM [TK].dbo.PURTI,[TK].dbo.PURTJ
                            WHERE TI001=TJ001 AND TI002=TJ002
                            AND TI003>='{0}' AND TI003<='{1}'
                                
                                ", SDATES, EDATES);

            Report report1 = new Report();
            report1.Load(@"REPORT\品質判定報告.frx");

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
            SETFASTREPORT(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
        }
        #endregion
    }
}
