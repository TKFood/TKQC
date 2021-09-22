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
    public partial class FrmREPORTQC : Form
    {
        public FrmREPORTQC()
        {
            InitializeComponent();
        }

        #region FUNCTION

        public void SETFASTREPORT()
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL1();
            Report report1 = new Report();
            report1.Load(@"REPORT\每月新增資料-非追.frx");

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

        public StringBuilder SETSQL1()
        {
            StringBuilder SB = new StringBuilder();

            if(comboBox1.Text.Equals("客戶"))
            {
                SB.AppendFormat(@"  SELECT MA001 AS ID,MA002 AS NAME");
                SB.AppendFormat(@"  FROM [TK].dbo.COPMA");
                SB.AppendFormat(@"  WHERE CREATE_DATE>='{0}' AND  CREATE_DATE<='{1}'",dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                SB.AppendFormat(@"  AND MA001 NOT LIKE '1%'");
                SB.AppendFormat(@"  ");
                SB.AppendFormat(@"  ");
            }
            else if (comboBox1.Text.Equals("供應商"))
            {
                SB.AppendFormat(@"  SELECT MA001 AS ID,MA002 AS NAME");
                SB.AppendFormat(@"  FROM [TK].dbo.PURMA");
                SB.AppendFormat(@"  WHERE CREATE_DATE>='{0}' AND  CREATE_DATE<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                SB.AppendFormat(@"  ");
                SB.AppendFormat(@"  ");
                SB.AppendFormat(@"  ");
            }
            else if (comboBox1.Text.Equals("品號"))
            {
                SB.AppendFormat(@"  SELECT MB001 AS ID,MB002 AS NAME");
                SB.AppendFormat(@"  FROM [TK].dbo.INVMB");
                SB.AppendFormat(@"  WHERE CREATE_DATE>='{0}' AND  CREATE_DATE<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                SB.AppendFormat(@"  AND MB001 LIKE '4%'");
                SB.AppendFormat(@"  ");
                SB.AppendFormat(@"  ");
            }
            SB.AppendFormat(@"  ");
            SB.AppendFormat(@"  ");
            SB.AppendFormat(@"  ");


            return SB;

        }


        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }
        #endregion
    }
}
