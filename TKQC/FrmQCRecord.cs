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
using DocumentFormat.OpenXml;
using System.Collections.Generic;
using Aspose.Words;

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
        int result;
        int rownum=0;

        public void SetValue()
        {
            this.label1.Text = shareArea.UserName;
        }

        public FrmQCRecord()
        {
            InitializeComponent();
            SetValue();
            button4.Visible = false;
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
            sbSql.Append("SELECT  [QCNO] AS '編號' ,[QCDATE] AS '受理日期',[CLIENT] AS '客戶',[FACTORY] AS '廠商',[DEP] AS '部門',[TEL] AS '電話',[Address] AS '地址',[MB001] AS '品號',[MB002] AS '品名',[MB003] AS '規格',[LOTNO] AS '料號/批號',[MANU] AS '線別',[TYPE] AS '使用範圍',[STATUS] AS '異常情況/應急對策',[PROCESS] AS '處理方式',[REASON] AS '原因分析',[PROTECT] AS '預防對策',[RESULT] AS '確認結果'");
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

        public void PRINTDOC()
        {
            // 首先把建立的範本檔案讀入MemoryStream
            //首先把建立的範本檔案讀入MemoryStream
            System.IO.MemoryStream _memoryStream = new System.IO.MemoryStream(Properties.Resources.品質異常處理單);

            //建立一個Document物件
            //並傳入MemoryStream
            Aspose.Words.Document doc = new Aspose.Words.Document(_memoryStream);

            //新增一個DataTable
            DataTable table = new DataTable();
            //建立Column
            table.Columns.Add("QCNO");
            table.Columns.Add("name");
            table.Columns.Add("QCDATE");
            table.Columns.Add("TEL");
            table.Columns.Add("Address");
            table.Columns.Add("MB001");
            table.Columns.Add("LOTNO");
            table.Columns.Add("MANU");
            table.Columns.Add("TYPE");
            table.Columns.Add("STATUS");
            table.Columns.Add("PROCESS");
            table.Columns.Add("REASON");
            table.Columns.Add("PROTECT");
            table.Columns.Add("RESULT");



            //透過建立的DataTable物件來New一個儲存資料的Row
            DataRow row = table.NewRow();
            //這些Row具有上面所建立相同的Column欄位
            //因此可以直接指定欄位名稱將資料填入裡面       
            row["QCNO"] = textBox1.Text.ToString();
            row["name"] = textBox2.Text.ToString() + textBox3.Text.ToString() + textBox4.Text.ToString();
            row["QCDATE"] = dateTimePicker1.Value.ToString("yyyy/MM/dd");
            row["TEL"] = textBox5.Text.ToString();
            row["Address"] = textBox6.Text.ToString();
            row["MB001"] = textBox7.Text.ToString() + textBox8.Text.ToString() + textBox9.Text.ToString();
            row["LOTNO"] = textBox10.Text.ToString();
            row["MANU"] = textBox11.Text.ToString();
            row["TYPE"] = textBox12.Text.ToString();
            row["STATUS"] = textBox13.Text.ToString();
            row["PROCESS"] = textBox14.Text.ToString();
            row["REASON"] = textBox15.Text.ToString();
            row["PROTECT"] = textBox16.Text.ToString();
            row["RESULT"] = textBox17.Text.ToString();

            //把所建立的資料行加入Table的Row清單內
            table.Rows.Add(row);


            //將DataTable傳入Document的MailMerge.Execute()方法
            doc.MailMerge.Execute(table);
            //清空所有未被合併的功能變數
            doc.MailMerge.DeleteFields();

            if (Directory.Exists(@"c:\temp\"))
            {
                //資料夾存在
            }
            else
            {
                //新增資料夾
                Directory.CreateDirectory(@"c:\temp\");
            }
            //將檔案儲存至c:\
            StringBuilder filename = new StringBuilder();
            filename.AppendFormat(@"c:\temp\品質異常處理單{0}.doc", DateTime.Now.ToString("yyyyMMdd"));
            doc.Save(filename.ToString());

            MessageBox.Show("匯出完成-文件放在-" + filename.ToString());
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
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if(dataGridView1.Rows.Count>=1)
            {
                textBox1.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
                textBox2.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                textBox3.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                textBox4.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
                textBox5.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
                textBox6.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();
                textBox7.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString();
                textBox8.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString();
                textBox9.Text = dataGridView1.CurrentRow.Cells[9].Value.ToString();
                textBox10.Text = dataGridView1.CurrentRow.Cells[10].Value.ToString();
                textBox11.Text = dataGridView1.CurrentRow.Cells[11].Value.ToString();
                textBox12.Text = dataGridView1.CurrentRow.Cells[12].Value.ToString();
                textBox13.Text = dataGridView1.CurrentRow.Cells[13].Value.ToString();
                textBox14.Text = dataGridView1.CurrentRow.Cells[14].Value.ToString();
                textBox15.Text = dataGridView1.CurrentRow.Cells[15].Value.ToString();
                textBox16.Text = dataGridView1.CurrentRow.Cells[16].Value.ToString();
                textBox17.Text = dataGridView1.CurrentRow.Cells[17].Value.ToString();
                dateTimePicker2.Value = Convert.ToDateTime(dataGridView1.CurrentRow.Cells[1].Value.ToString());
            }
            
        }

        public void ADDtoDB()
        {
            try
            {

                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" INSERT INTO  [{0}].[dbo].[QCRECORD] ([QCNO],[QCDATE],[CLIENT],[FACTORY],[DEP],[TEL],[Address],[MB001],[MB002],[MB003],[LOTNO],[MANU],[TYPE],[STATUS],[PROCESS],[REASON],[PROTECT],[RESULT])  VALUES ('{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}') ", sqlConn.Database.ToString(), textBox1.Text.ToString(), dateTimePicker2.Value.ToString("yyyy/MM/dd"), textBox2.Text.ToString(), textBox3.Text.ToString(), textBox4.Text.ToString(), textBox5.Text.ToString(), textBox6.Text.ToString(), textBox7.Text.ToString(), textBox8.Text.ToString(), textBox9.Text.ToString(), textBox10.Text.ToString(), textBox11.Text.ToString(), textBox12.Text.ToString(), textBox13.Text.ToString(), textBox14.Text.ToString(), textBox15.Text.ToString(), textBox16.Text.ToString(), textBox17.Text.ToString());
                //sbSql.AppendFormat("  UPDATE Member SET Cname='{1}',Mobile1='{2}' WHERE ID='{0}' ", list_Member[0].ID.ToString(), list_Member[0].Cname.ToString(), list_Member[0].Mobile1.ToString());

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易

                }

                sqlConn.Close();

                rownum = dataGridView1.RowCount;
                Search();


            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }

        }

        public void UpdateDB()
        {
            textBox1.ReadOnly = true;
            try
            {
                DialogResult dialogResult = MessageBox.Show("是否真的要更新", "UPDATE?", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();
                   

                    sbSql.AppendFormat("UPDATE [{0}].dbo.[QCRECORD]   SET [QCDATE]='{2}',[CLIENT]='{3}',[FACTORY]='{4}',[DEP]='{5}',[TEL]='{6}',[Address]='{7}',[MB001]='{8}',[MB002]='{9}',[MB003]='{10}',[LOTNO]='{11}',[MANU]='{12}',[TYPE]='{13}',[STATUS]='{14}',[PROCESS]='{15}',[REASON]='{16}',[PROTECT]='{17}',[RESULT]='{18}' WHERE [QCNO]='{1}' ", sqlConn.Database.ToString(), textBox1.Text.ToString(), dateTimePicker1.Value.ToString("yyyy/MM/dd"), textBox2.Text.ToString(), textBox3.Text.ToString(), textBox4.Text.ToString(), textBox5.Text.ToString(), textBox6.Text.ToString(), textBox7.Text.ToString(), textBox8.Text.ToString(), textBox9.Text.ToString(), textBox10.Text.ToString(), textBox11.Text.ToString(), textBox12.Text.ToString(), textBox13.Text.ToString(), textBox14.Text.ToString(), textBox15.Text.ToString(), textBox16.Text.ToString(), textBox17.Text.ToString());                  

                    cmd.Connection = sqlConn;
                    cmd.CommandTimeout = 60;
                    cmd.CommandText = sbSql.ToString();
                    cmd.Transaction = tran;
                    result = cmd.ExecuteNonQuery();

                    if (result == 0)
                    {
                        tran.Rollback();    //交易取消
                    }
                    else
                    {
                        tran.Commit();      //執行交易
                    }

                    sqlConn.Close();

                    Search();
                }

            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }

        public void ClearText()
        {
            textBox1.ReadOnly = false;
            textBox1.Text = null;
            textBox2.Text = null;
            textBox3.Text = null;
            textBox4.Text = null;
            textBox5.Text = null;
            textBox6.Text = null;
            textBox7.Text = null;
            textBox8.Text = null;
            textBox9.Text = null;
            textBox10.Text = null;
            textBox11.Text = null;
            textBox12.Text = null;
            textBox13.Text = null;
            textBox14.Text = null;
            textBox15.Text = null;
            textBox16.Text = null;
            textBox17.Text = null;
            dateTimePicker2.Value = DateTime.Now;

        }
        public void DelDB()
        {
            textBox1.ReadOnly = true;
            try
            {
                textBox1.Text = dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value.ToString();
                DialogResult dialogResult = MessageBox.Show("是否真的要刪除", "del?", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();
                    //sbSql.Append("UPDATE Member SET Cname='009999',Mobile1='009999',Telphone='',Email='',Address='',Sex='',Birthday='' WHERE ID='009999'");

                    sbSql.AppendFormat("DELETE [{0}].dbo.[QCRECORD] WHERE [QCNO]='{1}' ", sqlConn.Database.ToString(), textBox1.Text.ToString());
                    //sbSql.AppendFormat("  UPDATE Member SET Cname='{1}',Mobile1='{2}' WHERE ID='{0}' ", list_Member[0].ID.ToString(), list_Member[0].Cname.ToString(), list_Member[0].Mobile1.ToString());

                    cmd.Connection = sqlConn;
                    cmd.CommandTimeout = 60;
                    cmd.CommandText = sbSql.ToString();
                    cmd.Transaction = tran;
                    result = cmd.ExecuteNonQuery();

                    if (result == 0)
                    {
                        tran.Rollback();    //交易取消
                    }
                    else
                    {
                        tran.Commit();      //執行交易
                    }

                    sqlConn.Close();

                    Search();
                }

            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }

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

        private void button3_Click(object sender, EventArgs e)
        {
            PRINTDOC();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            ClearText();
            button4.Visible = true;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ADDtoDB();
            button4.Visible = false;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            UpdateDB();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            DelDB();
        }

        #endregion


    }
}
