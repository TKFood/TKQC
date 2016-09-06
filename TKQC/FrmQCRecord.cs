using System;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using NPOI.SS.UserModel;
using System.Configuration;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;


namespace TKQC
{
    public partial class FrmQCRecord : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();

        DataTable dt = new DataTable();


        public void SetValue()
        {
            this.label1.Text = shareArea.UserName;
        }

        public FrmQCRecord()
        {
            InitializeComponent();
            SetValue();

        }

        #region FUNCTION
        public void ExportExcel(DataSet dsExcel, string Tabelname)
        {            
            //建立Excel 2003檔案
            IWorkbook wb = new XSSFWorkbook();
            ISheet ws;

            dt = dsExcel.Tables[Tabelname];

            ////建立Excel 2007檔案
            //IWorkbook wb = new XSSFWorkbook();
            //ISheet ws;

            if (dt.TableName != string.Empty)
            {
                ws = wb.CreateSheet(dt.TableName);
            }
            else
            {
                ws = wb.CreateSheet("Sheet1");
            }

            ws.CreateRow(0);//第一行為欄位名稱
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ws.GetRow(0).CreateCell(i).SetCellValue(dt.Columns[i].ColumnName);
            }

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                ws.CreateRow(i + 1);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    ws.GetRow(i + 1).CreateCell(j).SetCellValue(dt.Rows[i][j].ToString());
                }
            }

            if (Directory.Exists(@"c:\temp\"))
            {
                //資料夾存在
            }
            else
            {
                //新增資料夾
                Directory.CreateDirectory(@"c:\temp\");
            }
            StringBuilder filename = new StringBuilder();
            filename.AppendFormat(@"c:\temp\品質異常處理單{0}.xlsx", DateTime.Now.ToString("yyyyMMdd"));

            FileStream file = new FileStream(filename.ToString(), FileMode.Create);//產生檔案
            wb.Write(file);
            file.Close();

            MessageBox.Show("匯出完成-EXCEL放在-" + filename.ToString());
            FileInfo fi = new FileInfo(filename.ToString());
            if (fi.Exists)
            {
                System.Diagnostics.Process.Start(filename.ToString());
            }
            else
            {
                //file doesn't exist
            }
        }

        public void Search()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);



            sbSql.Clear();
            sbSqlQuery.Clear();
            sbSql.Append("SELECT [QCNO],[QCDATE],[CLIENT],[FACTORY],[DEP],[TEL],[Address],[MB001],[MB002],[MB003],[LOTNO],[MANU],[TYPE],[STATUS],[PROCESS],[REASON],[PROTECT],[RESULT] ");
            sbSql.Append(" FROM [TKQC].[dbo].[QCRECORD]");
            sbSql.AppendFormat(" WHERE CONVERT(varchar(6),[QCDATE],112)='{0}'",dateTimePicker1.Value.ToString("yyyyMM") );
            sbSql.Append(" ");

            adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
            sqlCmdBuilder = new SqlCommandBuilder(adapter);

            sqlConn.Open();
            ds.Clear();
          
            adapter.Fill(ds, "TEMP1");
            dataGridView1.DataSource = ds.Tables["TEMP1"];

            sqlConn.Close();
        }
        #endregion

        #region BUTTON   

        private void button1_Click(object sender, EventArgs e)
        {
            Search();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            //ExportExcel(ds, NowTable);
        }

        #endregion
    }
}
