using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraBars;
using System.IO;
using System.Diagnostics;
using DX_QMS.Common;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraEditors;

namespace DX_QMS
{
    public partial class TestListSearch : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        IQC ic = new IQC();
        //Solder so = new Solder();
        System.Data.DataTable dt = null;
        public TestListSearch()
        {
            InitializeComponent();           
        }

        private void TestListSearch_Load(object sender, EventArgs e)
        {            
            bindDeviceType();
            bindTestItem(txttesttype.Text);
            bindTestTools();
            gridView.PrintExportProgress += new ProgressChangedEventHandler(gridView_PrintExportProgress);

        }

        private void bindDeviceType()
        {
            ////System.Data.DataTable dt = ic.SelectTestTypeRecord("查询", "", "测试类别", "").Tables[0];

            string sql = @"  select TestType from IQC_TestType where TestType= TestType and TTypes='测试类别'  ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            txttesttype.Properties.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
                txttesttype.Properties.Items.Add(row["TestType"]);
            }
             /////txttesttype.SelectedIndex = 0;
        }
        private void bindTestTools()
        {
           //// System.Data.DataTable dt = ic.SelectTestTypeRecord("查询", "", "测试工具", "").Tables[0];
            string sql = @"  select TestType from IQC_TestType where TestType= TestType and TTypes='测试工具'  ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            if (dt == null || dt.Rows.Count < 1)
                return;
            txttesttools.Properties.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
               // string s = row["TestType"].ToString();
                txttesttools.Properties.Items.Add(row["TestType"]);
            }
           //////// txttesttools.SelectedIndex = 0;
        }
        private void bindTestItem(string TestType)
        {
            ///// DataSet ds = ic.SelectTestItemRecord("查询", TestType, "", 0, "");

            //string sql = " select TestItem from IQC_TestItem where TestType='"+TestType+ "' order by Item ";
            //DataSet ds = Common.DbAccess.SelectBySql(sql);
            //if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            //{
            //    System.Data.DataTable dt = ds.Tables[0];               
            //    txttestitem.Properties.Items.Clear();
            //    foreach (DataRow row in dt.Rows)
            //    {
            //        txttestitem.Properties.Items.Add(row["TestItem"]);
            //    }
            //    txttestitem.SelectedIndex = 0;
            //}

            DataSet ds = ic.SelectTestItemRecord("查询", TestType, "", 0, "");
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                DataTable dt = ds.Tables[0];
                txttestitem.Properties.Items.Clear();
                foreach (DataRow row in dt.Rows)
                {
                    txttestitem.Properties.Items.Add(row["TestItem"]);
                }
                 txttestitem.SelectedIndex = 0;
            }

        }
        public DateTime beforeTime, afterTime;
        private void TestListSearch_FormClosing(object sender, FormClosingEventArgs e)
        {
            Process[] myProcesses;
            DateTime startTime;
            myProcesses = Process.GetProcessesByName("Excel");

            // 得不到Excel进程ID，暂时只能判断进程启动时间 
            foreach (Process myProcess in myProcesses)
            {
                startTime = myProcess.StartTime;

                if (startTime > beforeTime && startTime < afterTime)
                {
                    myProcess.Kill();
                }
            }
        }
        private void txttesttype_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txttesttype.Text == "")
                return;
            bindTestItem(txttesttype.Text);
        }
        private void bindTestSubItem(string testtype, string Testitem)
        {
            string ssql = "select * from IQC_TestSubItemSet where TestType=case '" +testtype +"' when '' then TestType else '" +testtype +"' end and TestItem= case '"+Testitem+"' when '' then TestItem else '"+Testitem+"' end";
            DataSet ds = DbAccess.SelectBySql(ssql);
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                DataTable dt = ds.Tables[0];               
                txttestsubitem.Properties.Items.Clear();
                foreach (DataRow row in dt.Rows)
                {
                    txttestsubitem.Properties.Items.Add(row["TestSubItem"]);
                }
                 txttestsubitem.SelectedIndex = 0;
            }
        }
        private void txttestitem_SelectedIndexChanged(object sender, EventArgs e)
        {
                if (txttestitem.Text != "")
                bindTestSubItem(txttesttype.Text == "" ? "" : txttesttype.Text, txttestitem.Text == "" ? "" : txttestitem.Text);
           
        }

        private void btnreset_Click(object sender, EventArgs e)
        {
            txttesttype.Text = "";
            txttestitem.Text = "";
            txttestsubitem.Text = "";
            txttesttools.Text = "";
            txttestuser.Text = "";
            txtmaterialcode.Text = "";
            txtmaterialname.Text = "";
            txtvendorcode.Text = "";
            txtvendorname.Text = "";
            txtlotno.Text = "";
            bdate.Text = "";
            edate.Text = "";
            txtcusvendor.Text = "";
            databind.DataSource = null;

        }
        public void DataToExcel(DataGridView m_DataView)
        {
            SaveFileDialog kk = new SaveFileDialog();
            kk.Title = "保存EXECL文件";
            kk.Filter = "EXECL文件(*.xls) |*.xls";
            kk.FilterIndex = 1;

            if (kk.ShowDialog() == DialogResult.OK)
            {
                string FileName = kk.FileName;
                if (File.Exists(FileName))
                    File.Delete(FileName);
                FileStream objFileStream;
                StreamWriter objStreamWriter;
                string strLine = "";
                objFileStream = new FileStream(FileName, FileMode.OpenOrCreate, FileAccess.Write);
                objStreamWriter = new StreamWriter(objFileStream, System.Text.Encoding.Unicode);

                for (int i = 0; i < m_DataView.Columns.Count; i++)
                {
                    if (m_DataView.Columns[i].Visible == true)
                    {
                        strLine = strLine + m_DataView.Columns[i].HeaderText.ToString() + Convert.ToChar(9);
                    }
                }
                objStreamWriter.WriteLine(strLine);
                strLine = "";

                for (int i = 0; i < m_DataView.Rows.Count; i++)
                {
                    if (m_DataView.Columns[0].Visible == true)
                    {
                        if (m_DataView.Rows[i].Cells[0].Value == null)
                            strLine = strLine + " " + Convert.ToChar(9);
                        else
                            strLine = strLine + m_DataView.Rows[i].Cells[0].Value.ToString() + Convert.ToChar(9);
                    }
                    for (int j = 1; j < m_DataView.Columns.Count; j++)
                    {

                        if (m_DataView.Columns[j].Visible == true)
                        {
                            if (m_DataView.Rows[i].Cells[j].Value == null)
                                strLine = strLine + " " + Convert.ToChar(9);
                            else
                            {
                                string rowstr = "";
                                rowstr = m_DataView.Rows[i].Cells[j].Value.ToString();
                                if (rowstr.IndexOf("\r\n") > 0)
                                    rowstr = rowstr.Replace("\r\n", " ");
                                if (rowstr.IndexOf("\t") > 0)
                                    rowstr = rowstr.Replace("\t", " ");
                                if (rowstr.IndexOf("\n") > 0)
                                    rowstr = rowstr.Replace("\n", " ");

                                strLine = strLine + rowstr + Convert.ToChar(9);
                            }
                        }
                    }
                    objStreamWriter.WriteLine(strLine);
                    strLine = "";
                }
                objStreamWriter.Close();
                objFileStream.Close();

            }
        }
        private string ShowSaveFileDialog(string title, string filter)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            string name = "IQC测试记录信息";
            int n = name.LastIndexOf(".") + 1;
            if (n > 0) name = name.Substring(n, name.Length - n);
            dlg.Title = "Export To " + title;
            dlg.FileName = name;
            dlg.Filter = filter;
            if (dlg.ShowDialog() == DialogResult.OK) return dlg.FileName;
            return "";
        }
        private void ExportToEx(String filename, string ext, BaseView exportView)
        {
            Cursor currentCursor = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;
            if (ext == "rtf") exportView.ExportToRtf(filename);
            if (ext == "pdf") exportView.ExportToPdf(filename);
            if (ext == "mht") exportView.ExportToMht(filename);
            if (ext == "htm") exportView.ExportToHtml(filename);
            if (ext == "txt") exportView.ExportToText(filename);
            if (ext == "xls") exportView.ExportToXls(filename);
            if (ext == "xlsx") exportView.ExportToXlsx(filename);
            Cursor.Current = currentCursor;
        }
        private void OpenFile(string fileName)
        {
            if (XtraMessageBox.Show("Do you want to open this file?", "Export To...", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    System.Diagnostics.Process process = new System.Diagnostics.Process();
                    process.StartInfo.FileName = fileName;
                    process.StartInfo.Verb = "Open";
                    process.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal;
                    process.Start();
                }
                catch
                {
                    DevExpress.XtraEditors.XtraMessageBox.Show(this, "Cannot find an application on your system suitable for openning the file with exported data.", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            progressBarControl1.Position = 0;
        }
        void gridView_PrintExportProgress(object sender, ProgressChangedEventArgs e)
        {
            SetPosition(e.ProgressPercentage);
        }
        void SetPosition(int pos)
        {
            progressBarControl1.Position = pos;
            this.Update();
        }
        private void btnToExcel_Click(object sender, EventArgs e)
        {
            //if (databind.Rows.Count <= 0) return;
            //DataToExcel(databind);
            DataTable dt = databind.DataSource as DataTable;
            if (dt == null)
                return;
            if (dt.Rows.Count <= 0) return;
            string fileName = ShowSaveFileDialog("Microsoft Excel 2007 Document", "Microsoft Excel|*.xlsx");
            if (fileName == string.Empty) return;
            ExportToEx(fileName, "xlsx", gridView);
            OpenFile(fileName);
        }
        private string GetReAc(string Sampleqty, string AQL)
        {
            string AR = "";
            string ssql = "select Sampleqty,AQL,cast(Ac as varchar(10))+'/'+cast(Re as varchar(10)) AR from IQC_TestSTD105ECode s left join  IQC_TestAQLRcSet r on s.Code=r.Code where Sampleqty='" + Sampleqty + "' and AQL='" + AQL + "'";
            DataSet ds = DbAccess.SelectBySql(ssql);
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                AR = ds.Tables[0].Rows[0]["AR"].ToString();
            }
            return AR;
        }

        private string GetISO2859ReAc(string materialstate, int qty ,string AQL)
        {

            string AR = "0/1";
            string ssql = "";
            DataTable dt = null;

            if (materialstate == "放宽检验")
            {
                ssql = " select cast(a.Ac as varchar(10))+'/'+cast(a.Re as varchar(10)) AR from IHPS_QUALITY_SPC_AQLIS02859 a left join IQC_TestSTD105ECheckSet i on i.Code1=a.Code left join IQC_TestSTD105ECode c on c.Code = a.Code where Type = '放宽'  and(LotSizemin <= "+qty+"  and LotSizemax >="+qty+ " and CheckLevel = 'II') and AQLValue ="+ AQL+" ";
                dt = DbAccess.SelectBySql(ssql).Tables[0];
                if (dt != null  && dt.Rows.Count > 0)
                {
                    AR = dt.Rows[0]["AR"].ToString();
                }
            }
            else if (materialstate == "正常检验")
            {
                ssql = " select cast(a.Ac as varchar(10))+'/'+cast(a.Re as varchar(10)) AR from IHPS_QUALITY_SPC_AQLIS02859 a left join IQC_TestSTD105ECheckSet i on i.Code=a.Code left join IQC_TestSTD105ECode c on c.Code = i.Code  where Type = '正常检验'  and(LotSizemin <="+qty+ " and LotSizemax >="+qty+ " and CheckLevel = 'II') and AQLValue = "+AQL+" ";
                DataSet ds = DbAccess.SelectBySql(ssql);
                dt = DbAccess.SelectBySql(ssql).Tables[0];
                if (dt != null && dt.Rows.Count > 0)
                {
                    AR = dt.Rows[0]["AR"].ToString();
                }
            }
            else if (materialstate == "加严检验")
            {
                ssql = " select cast(a.Ac as varchar(10))+'/'+cast(a.Re as varchar(10)) AR from IHPS_QUALITY_SPC_AQLIS02859 a left join IQC_TestSTD105ECheckSet i on i.Code=a.Code left join IQC_TestSTD105ECode c on c.Code = i.Code  where Type = '加严'  and(LotSizemin <=" + qty + " and LotSizemax >=" + qty + " and CheckLevel = 'II') and AQLValue = " + AQL + " ";
                DataSet ds = DbAccess.SelectBySql(ssql);
                dt = DbAccess.SelectBySql(ssql).Tables[0];
                if (dt != null && dt.Rows.Count > 0)
                {
                    AR = dt.Rows[0]["AR"].ToString();
                }
            }
            return AR;

        }

        string cusvendor(string productcode,string receptid)
        {
            string sql = " select distinct manufacturer 供应商代码  from IQC_TestList where  Productcode = '"+productcode+ "'  and receptid = '"+receptid+ "' and manufacturer <>'' and manufacturer is not null   ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            if (dt != null && dt.Rows.Count > 0)
            {
                string vendor = "";
                for (int i = 0;i<dt.Rows.Count;i++)
                {                 
                    if (dt.Rows[i]["供应商代码"].ToString() != "")
                    {
                        vendor += dt.Rows[i]["供应商代码"].ToString();
                    }
                }
                return vendor;
            }
            return "";  
        }
        public void Copy(string sheetPrefixName, string sup, string dcode, string orderno,string materialstate, System.Data.DataTable tb, int qty)
        {
            string finaltotalstate = "", testuser = "", AQLValue = "";
            int sheetCount = 1;
            if (sheetPrefixName == null || sheetPrefixName.Trim() == "")
                sheetPrefixName = " Sheet ";
            beforeTime = DateTime.Now;

            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = true;
            afterTime = DateTime.Now;
            object missing = System.Reflection.Missing.Value;
            string templetFile = Environment.CurrentDirectory + @"\Resources\Report.xls";
            Microsoft.Office.Interop.Excel.Workbook workBook = app.Workbooks.Open(templetFile, missing, true, missing, missing, missing,
                                                          missing, missing, missing, missing, missing, missing, missing, missing, missing);

            Microsoft.Office.Interop.Excel.Worksheet workSheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.Sheets.get_Item(1);

            for (int i = 1; i < sheetCount; i++)
            {
                ((Microsoft.Office.Interop.Excel.Worksheet)workBook.Worksheets.get_Item(i)).Copy(missing, workBook.Worksheets[i]);
            }

            #region  将源DataTable数据写入Excel

            Microsoft.Office.Interop.Excel.Worksheet sheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.Worksheets.get_Item(1);
            sheet.Name = sheetPrefixName.Replace("/", "");

            string lotnos = "";
            System.Data.DataTable dtlotno = null;

            for (int j = 0; j < tb.Rows.Count; j++)
            {
                string finalstate = "";
                sheet.Cells[2, 1] = "IQC检验报表(" + tb.Rows[0 + j]["TestType"].ToString() + ")";
                sheet.Cells[4, 2] = sup;
                sheet.Cells[4, 18] = dcode;
                sheet.Cells[3, 2] = tb.Rows[0 + j]["TestTime"].ToString();
                sheet.Cells[3, 6] = tb.Rows[0 + j]["lot_number"].ToString();

                if (qty == 0)
                    sheet.Cells[4, 12] = tb.Rows[0 + j]["Qty"].ToString();
                else
                    sheet.Cells[4, 12] = qty.ToString();

                sheet.Cells[3, 17] = tb.Rows[0 + j]["LotNo"].ToString();
                sheet.Cells[46, 19] = tb.Rows[0 + j]["TestUser"].ToString();
                sheet.Cells[5, 2] = tb.Rows[0 + j]["Productcode"].ToString();
                sheet.Cells[5, 7] = tb.Rows[0 + j]["materialenname"].ToString();
                sheet.Cells[10 + j + 1, 1] = tb.Rows[0 + j]["TestItem"].ToString();
                sheet.Cells[10 + j + 1, 2] = tb.Rows[0 + j]["TestSubItem"].ToString();
                sheet.Cells[10 + j + 1, 3] = tb.Rows[0 + j]["TestDes"].ToString();
                sheet.Cells[10 + j + 1, 9] = tb.Rows[0 + j]["TestTools"].ToString();
                sheet.Cells[10 + j + 1, 11] = tb.Rows[0 + j]["SampleType"].ToString();
                if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("MI"))
                {
                    sheet.Cells[10 + j + 1, 23] = "√";
                    AQLValue = "1.5";
                }
                else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("MA"))
                {
                    sheet.Cells[10 + j + 1, 21] = "√";
                    AQLValue = "0.65";
                }
                else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("CR=0.1"))
                {
                    sheet.Cells[10 + j + 1, 19] = "√";
                    AQLValue = "0.1";
                }
                else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=0.40"))
                {
                    sheet.Cells[10 + j + 1, 20] = "√";
                    AQLValue = "0.4";
                }
                else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=1.0"))
                {
                    sheet.Cells[10 + j + 1, 22] = "√";
                    AQLValue = "1";
                }
                else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("CR=0.01"))
                {
                    sheet.Cells[10 + j + 1, 18] = "√";
                    AQLValue = "0.01";
                }
                else if (tb.Rows[0 + j]["SampleType"].ToString().Contains("ISO2859-1"))
                {
                    if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=1.5"))
                    {
                        sheet.Cells[10 + j + 1, 23] = "√";
                        AQLValue = "1.5";
                    }
                    else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=0.65"))
                    {
                        sheet.Cells[10 + j + 1, 21] = "√";
                        AQLValue = "0.65";
                    }
                    else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=0.10"))
                    {
                        sheet.Cells[10 + j + 1, 19] = "√";
                        AQLValue = "0.1";
                    }
                    else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=0.40"))
                    {
                        sheet.Cells[10 + j + 1, 20] = "√";
                        AQLValue = "0.4";
                    }
                    else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=1.0"))
                    {
                        sheet.Cells[10 + j + 1, 22] = "√";
                        AQLValue = "1";
                    }
                    else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=0.010"))
                    {
                        sheet.Cells[10 + j + 1, 18] = "√";
                        AQLValue = "0.01";
                    }
                }
                else if (tb.Rows[0 + j]["SampleType"].ToString().Contains("C=0"))
                {
                    if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=1.5"))
                    {
                        sheet.Cells[10 + j + 1, 23] = "√";
                    }
                    else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=0.65"))
                    {
                        sheet.Cells[10 + j + 1, 21] = "√";
                    }
                    else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=0.1"))
                    {
                        sheet.Cells[10 + j + 1, 19] = "√";
                    }
                    else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=0.4"))
                    {
                        sheet.Cells[10 + j + 1, 20] = "√";
                    }
                    else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=1"))
                    {
                        sheet.Cells[10 + j + 1, 22] = "√";
                    }
                    else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=0.01"))
                    {
                        sheet.Cells[10 + j + 1, 18] = "√";
                    }
                }
              
                if (tb.Rows[0 + j]["SampleType"].ToString().Contains("ISO2859-1"))
                {                    
                   sheet.Cells[10 + j + 1, 17] = GetISO2859ReAc(materialstate, qty, AQLValue);
                }
                else if (tb.Rows[0 + j]["SampleType"].ToString().Contains("C=0"))
                {
                    sheet.Cells[10 + j + 1, 17] = "0/1";
                }
                else
                {
                    sheet.Cells[10 + j + 1, 17] = GetReAc(tb.Rows[0 + j]["Sampleqty"].ToString(), tb.Rows[0 + j]["AQL"].ToString());
                }


                if (materialstate == "正常检验")
                {
                    sheet.Cells.get_Range("F7").Value = "√正常检验";
                    sheet.Cells.get_Range("F7").Font.Size = 11;
                    sheet.Cells.get_Range("F7").Font.Bold = true;
                    sheet.Cells.get_Range("F7").EntireColumn.AutoFit();
                    sheet.Cells.get_Range("F7").EntireRow.AutoFit();
                    sheet.Cells.get_Range("F7").Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.LightGreen);
                }
                else if (materialstate == "加严检验")
                {
                    sheet.Cells.get_Range("B7").Value = "√加严检验";
                    sheet.Cells.get_Range("B7").Font.Size = 11;
                    sheet.Cells.get_Range("B7").Font.Bold = true;
                    sheet.Cells.get_Range("B7").EntireColumn.AutoFit();
                    sheet.Cells.get_Range("B7").EntireRow.AutoFit();
                    sheet.Cells.get_Range("B7").Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Red);

                }
                else if (materialstate == "放宽检验")
                {
                    sheet.Cells.get_Range("K7").Value = "√放宽检验";
                    sheet.Cells.get_Range("K7").Font.Size = 11;
                    sheet.Cells.get_Range("K7").Font.Bold = true;
                    sheet.Cells.get_Range("K7").EntireColumn.AutoFit();
                    sheet.Cells.get_Range("K7").EntireRow.AutoFit();
                    sheet.Cells.get_Range("K7").Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.LightYellow);

                }
                //else if (dt.Rows[0 + j]["CheckType"].ToString() == "全检")
                //{
                //    sheet.Cells.get_Range("P7").Value = "√全检";
                //    sheet.Cells.get_Range("P7").Font.Size = 11;
                //    sheet.Cells.get_Range("P7").Font.Bold = true;
                //    sheet.Cells.get_Range("P7").EntireColumn.AutoFit();
                //    sheet.Cells.get_Range("P7").EntireRow.AutoFit();
                //}
                System.Data.DataTable dlist = ic.SelectTestListRecord("测试明细", "", tb.Rows[j]["TestItem"].ToString(), tb.Rows[j]["TestSubItem"].ToString(), "", tb.Rows[j]["lotno"].ToString(), "", tb.Rows[j]["Productcode"].ToString(), "", "", "", "", "").Tables[0];
               
                dtlotno = dlist;
                string s = "";
                int testqty = 0;
                for (int m = 0; m < dlist.Rows.Count; m++)
                {
                    testqty = testqty + 1;
                    s = s + (m + 1).ToString() + ".测试值:" + dlist.Rows[0 + m]["Testvalue"].ToString() + ",检验结果:" + dlist.Rows[0 + m]["TestResult"].ToString() + ",备注:" + dlist.Rows[0 + m]["Remarks"].ToString() + ";";
                    finalstate = finalstate + dlist.Rows[0 + m]["TestFinalResult"].ToString();
                }
                finaltotalstate = finaltotalstate + finalstate;
                testuser = tb.Rows[j]["TestUser"].ToString();

                sheet.Cells[10 + j + 1, 12] = s.TrimEnd(';').TrimStart(';');
                if (finalstate.Contains("拒收"))

                    sheet.Cells[10 + j + 1, 24] = "NG";
                else

                    sheet.Cells[10 + j + 1, 24] = "OK";
            }
            System.Data.DataTable newtb = dtlotno.Clone();
            newtb.Rows.Add(dtlotno.Rows[0].ItemArray);
            lotnos = lotnos + "," + dtlotno.Rows[0]["LotNos"].ToString();
            for (int i = 1; i < dtlotno.Rows.Count; i++)
            {
                bool flag = true;
                foreach (DataRow dr in newtb.Rows)
                {
                    if (dtlotno.Rows[i]["LotNos"].ToString() == dr["LotNos"].ToString())
                    {
                        flag = false;
                        continue;
                    }
                }
                if (flag)
                {
                    newtb.Rows.Add(dtlotno.Rows[i].ItemArray);
                    lotnos = lotnos + "," + dtlotno.Rows[i]["LotNos"].ToString();
                }
            }
            #endregion

            sheet.Cells[4, 7] = lotnos.Trim(',').ToUpper();

            string ssql = "select approvalman, Auditingman from  IQC_Approval";
            DataSet ds = Common.DbAccess.SelectBySql(ssql);
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                sheet.Cells[46, 9] = ds.Tables[0].Rows[0]["approvalman"].ToString();
                sheet.Cells[46, 13] = ds.Tables[0].Rows[0]["Auditingman"].ToString();
            }
            sheet.Cells[46, 20] = testuser;

            if (finaltotalstate.Contains("拒收"))
            {
                sheet.Cells[44, 12] = "√ 判退";
                sheet.Cells[44, 2] = "  合格";
            }
            else
            {
                sheet.Cells[44, 2] = "√ 合格";
                sheet.Cells[44, 12] = "  判退";
            }

        }
        public void OpenUrl(string url)
        {
            Process pro = new Process();
            pro.StartInfo.FileName = "iexplore.exe";
            pro.StartInfo.Arguments = url;
            pro.Start();
        }
        private void databind_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            //DataGridView dgv = (DataGridView)sender;
            //if (dgv.Columns[e.ColumnIndex].Name == "订单号")
            //{
            //    dt = ic.SelectTestListRecord("测试汇总", "", "", "", "", databind.CurrentRow.Cells["订单号"].Value.ToString(), "", databind.CurrentRow.Cells["料号"].Value.ToString(), "", "", "", "", "").Tables[0];
            //    Copy(dt.Rows[0]["TestType"].ToString(), databind.CurrentRow.Cells["供应商名称"].Value.ToString(), databind.CurrentRow.Cells["生成时间"].Value.ToString(), databind.CurrentRow.Cells["订单号"].Value.ToString(), dt);
            //}
            //else if (dgv.Columns[e.ColumnIndex].Name == "全尺寸数据")
            //{
            //    string s = databind.CurrentRow.Cells["料号"].Value.ToString();
            //    OpenUrl("http://webapp2.hytera.com/weboa/app/SC/supplierReport.nsf/文档按分类?SearchView&Query=" + s + "&SearchMax=9999&SearchOrder=3&Start=1&Count=9999&ViewName=vw_AllDocByCategory&ViewAliases=文档按分类");
            //}
        }

        private void databind_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //TestListRecord TLR = new TestListRecord(databind.CurrentRow.Cells["订单号"].Value.ToString(), databind.CurrentRow.Cells["料号"].Value.ToString());
            //TLR.ShowDialog();
        }

        private void databind_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            //using (SolidBrush b = new SolidBrush(databind.RowHeadersDefaultCellStyle.ForeColor))
            //{
            //    string linenum = e.RowIndex.ToString();
            //    int linen = 0;
            //    linen = Convert.ToInt32(linenum) + 1;
            //    string line = linen.ToString();
            //    e.Graphics.DrawString(line, e.InheritedRowStyle.Font, b, e.RowBounds.Location.X + 20, e.RowBounds.Location.Y + 5);
            //    SolidBrush B = new SolidBrush(Color.Red);
            //}
        }

        private void databind_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            //DataGridViewRow dgr = databind.Rows[e.RowIndex];
            //try
            //{
            //    if (dgr.Cells["检验结果"].Value.ToString() == "NG")
            //    {
            //        dgr.DefaultCellStyle.BackColor = Color.Red;
            //    }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}
        }


        string MaterialState(string Materialcode,string receptid)
        {
            string sql, State = "正常检验";
            sql = @"  select  distinct CheckType  from IQC_TestList  where Productcode = '"+Materialcode+ "'  and receptid = '"+receptid+ "' and NGtype is not null ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            if (dt != null && dt.Rows.Count > 0)
            {
                State = dt.Rows[0]["CheckType"].ToString();
                if (!State.Contains("检验"))
                {
                    State = "正常检验";
                }
            }
            return State;
        }


        private void gridView_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {
            //DataGridView dgv = (DataGridView)sender;
            //if (dgv.Columns[e.ColumnIndex].Name == "订单号")
            //{
            //    dt = ic.SelectTestListRecord("测试汇总", "", "", "", "", databind.CurrentRow.Cells["订单号"].Value.ToString(), "", databind.CurrentRow.Cells["料号"].Value.ToString(), "", "", "", "", "").Tables[0];
            //    Copy(dt.Rows[0]["TestType"].ToString(), databind.CurrentRow.Cells["供应商名称"].Value.ToString(), databind.CurrentRow.Cells["生成时间"].Value.ToString(), databind.CurrentRow.Cells["订单号"].Value.ToString(), dt);
            //}
            //else if (dgv.Columns[e.ColumnIndex].Name == "全尺寸数据")
            //{
            //    string s = databind.CurrentRow.Cells["料号"].Value.ToString();
            //    OpenUrl("http://webapp2.hytera.com/weboa/app/SC/supplierReport.nsf/文档按分类?SearchView&Query=" + s + "&SearchMax=9999&SearchOrder=3&Start=1&Count=9999&ViewName=vw_AllDocByCategory&ViewAliases=文档按分类");
            //}  gridView.GetFocusedRowCellValue("料号").ToString();

            if (e.Column.FieldName == "接收单号")
            {
                int qty = 0;
                Cursor.Current = Cursors.WaitCursor;
                if (gridView.GetFocusedRowCellValue("名称").ToString().Contains("客供】"))
                {
                    string sql = " select  round(sum(qty),0) qty from deliveryEMSOtherRec where deliveryid = '"+gridView.GetFocusedRowCellValue("接收单号").ToString()+ "' and  materialcode = '"+gridView.GetFocusedRowCellValue("料号").ToString()+"'  group by deliveryid,materialcode  ";
                    DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        if (dt.Rows[0][0].ToString() != "")
                            qty = int.Parse(dt.Rows[0][0].ToString());
                    }
                    else
                    {
                        string ssql = " select sum(qty) qty from delivery where deliveryid = '"+gridView.GetFocusedRowCellValue("接收单号").ToString()+ "' and materialcode = '"+gridView.GetFocusedRowCellValue("料号").ToString()+ "'  group by deliveryid,materialcode  ";
                        DataTable dtt = DbAccess.SelectBySql(ssql).Tables[0];
                        if (dtt.Rows[0][0].ToString() != "")
                            qty = int.Parse(dtt.Rows[0][0].ToString());
                    }
                }
                else
                {
                    //string sqloracle = "select sum(接收数量) qty from apps.CUX_INV_RECEIVE_DATE where 接收单号='" + gridView.GetFocusedRowCellValue("接收单号").ToString() + "' and 物料='" + gridView.GetFocusedRowCellValue("料号").ToString() + "'";
                    string sqloracle = " select sum(数量) qty from APPS.CUX_INVIQC_DAY_V where  物料编码='" + gridView.GetFocusedRowCellValue("料号").ToString() + "'  and  接收单号='" + gridView.GetFocusedRowCellValue("接收单号").ToString() + "' and  处理类型 ='接收'";
                    DataTable dtoracle = DbAccess.SelectByOracle(sqloracle).Tables[0];
                    if (dtoracle.Rows[0][0].ToString() != "")
                        qty = int.Parse(dtoracle.Rows[0][0].ToString());
                }
                     
                dt = ic.SelectTestListRecord("测试汇总", "", "", "", "", gridView.GetFocusedRowCellValue("接收单号").ToString(), "", gridView.GetFocusedRowCellValue("料号").ToString(), "", "", "", "", "").Tables[0];
               
                string materialcode = gridView.GetFocusedRowCellValue("料号").ToString();
                string materialstate = MaterialState(materialcode, gridView.GetFocusedRowCellValue("接收单号").ToString().Trim());


                //// Copy(dt.Rows[0]["TestType"].ToString(), gridView.GetFocusedRowCellValue("供应商名称").ToString(), gridView.GetFocusedRowCellValue("生成时间").ToString(), gridView.GetFocusedRowCellValue("接收单号").ToString(), materialstate, dt, qtyerp);
                try
                {
                    Copy(dt.Rows[0]["TestType"].ToString(), gridView.GetFocusedRowCellValue("供应商名称").ToString(), gridView.GetFocusedRowCellValue("生成时间").ToString(), gridView.GetFocusedRowCellValue("接收单号").ToString(), materialstate, dt, qty);
                }
              catch
               {
                MessageBox.Show("生成完整报表失败", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Cursor.Current = Cursors.Default;
                return;
               }
                Cursor.Current = Cursors.Default;
            }
            else if (e.Column.FieldName == "全尺寸数据")
            {
                string s = gridView.GetFocusedRowCellValue("料号").ToString();
                OpenUrl("http://webapp2.hytera.com/weboa/app/SC/supplierReport.nsf/文档按分类?SearchView&Query=" + s + "&SearchMax=9999&SearchOrder=3&Start=1&Count=9999&ViewName=vw_AllDocByCategory&ViewAliases=文档按分类");
            }

        }

        private void gridView_DoubleClick(object sender, EventArgs e)
        {
            TestListRecord TLR = new TestListRecord(gridView.GetFocusedRowCellValue("接收单号").ToString(), gridView.GetFocusedRowCellValue("料号").ToString());
            TLR.ShowDialog();
        }

        private void gridView_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
        {        
            try
            {
                DataTable dt = databind.DataSource as DataTable;
                if (dt == null || dt.Rows.Count < 1)
                {
                    return;
                }
                if (gridView.GetDataRow(e.RowHandle)["检验结果"].ToString() == "NG")
                {
                    e.Appearance.BackColor = Color.Red;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void gridView_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            gridView.IndicatorWidth = 60;

            if (e.Info.IsRowIndicator && e.RowHandle >= 0)
            {
                e.Info.DisplayText = (e.RowHandle + 1).ToString();
            }
        }

        void selectdata(object obj)
        {

            string sqlstr = "", sqlwhere1 = " where 1=1", sqlwhere2 = " where 1=1";
            string materialcode = "", materialname = "", vendorcode = "", vendorname = "";
            string testtype = "", testitem = "", testsubitem = "", testtools = "", testuser = "", lotno = "", cusvendor = "";

            materialcode = txtmaterialcode.Text.Trim();
            materialname = txtmaterialname.Text.Trim();
            vendorcode = txtvendorcode.Text.Trim();
            vendorname = txtvendorname.Text.Trim();


            testtype = txttesttype.Text.Trim();
            testitem = txttestitem.Text.Trim();
            testsubitem = txttestsubitem.Text.Trim();
            testtools = txttesttools.Text.Trim();
            testuser = txttestuser.Text.Trim();
            lotno = txtlotno.Text.Trim();
            cusvendor = txtcusvendor.Text;


            string sbdate = "", sedate = "";
            sbdate = bdate.Text;
            sedate = edate.Text;

            databind.DataSource = null;
            DataSet ds = null;

            if (!string.IsNullOrEmpty(materialcode))
            {
                sqlwhere1 += " and d.materialcode like '%" + materialcode + "%' ";
            }
            if (!string.IsNullOrEmpty(materialname))
            {
                sqlwhere1 += "  and materialname like '%" + materialname + "%' ";
            }
            if (!string.IsNullOrEmpty(vendorcode))
            {
                sqlwhere1 += " and vendorcode like '%" + vendorcode + "%' ";
            }
            if (!string.IsNullOrEmpty(vendorname))
            {
                sqlwhere1 += "  and vendorname like '%" + vendorname + "%' ";
            }


            if (!string.IsNullOrEmpty(testtype))
            {
                sqlwhere2 += " and testtype='" + testtype + "' ";
            }
            if (!string.IsNullOrEmpty(testitem))
            {
                sqlwhere2 += " and testitem='" + testitem + "' ";
            }
            if (!string.IsNullOrEmpty(testsubitem))
            {
                sqlwhere2 += " and testsubitem='" + testsubitem + "' ";
            }
            if (!string.IsNullOrEmpty(testtools))
            {
                sqlwhere2 += "  and TestTools like '%" + testtools + "%' ";
            }
            if (!string.IsNullOrEmpty(testuser))
            {
                sqlwhere2 += " and TestUser like '%" + testuser + "%' ";
            }
            if (!string.IsNullOrEmpty(lotno))
            {
                sqlwhere2 += "  and Lotno like '%" + lotno + "%' ";
            }
            if (!string.IsNullOrEmpty(cusvendor))
            {
                sqlwhere2 += "  and manufacturer = '"+ cusvendor + "' ";
            }
            if (!string.IsNullOrEmpty(sbdate))
            {
                sqlwhere2 += " and testtime >= '" + sbdate + " 00:00:00 ' ";
            }
            if (!string.IsNullOrEmpty(sedate))
            {
                sqlwhere2 += " and testtime <= '" + sedate + " 23:59:59 ' ";
            }
            sqlwhere2 = sqlwhere2 + "   group by receptid,Productcode,Lotno) x on d.deliveryid=x.receptid and d.materialcode=x.Productcode and d.lotno=x.LotNo ";

            sqlstr = @"  select * from
                        ( select d.deliveryid 接收单号,d.materialcode 料号,max(materialname) 名称,case when max(testqty)>1 then max(testqty) else sum(cast(qty as bigint)) end 数量,max(vendorcode) 供应商代码,max(vendorname) 供应商名称,min(transactiondate) 生成时间,
			             case max(TestFinalResult) when '拒收' then 'NG' else 'OK' end 检验结果,
                         max(testtime) 测试日期,max(TestType) 测试类别,Max(Remarks) 备注,max(TestUser) 检验员,d.materialcode 全尺寸数据 
				         from delivery d left join MaterialSpec m on d.materialcode=m.materialcode  ";
            sqlstr = sqlstr + "   right join (select receptid,Productcode,LotNo,max(testqty) testqty,max(TestFinalResult) TestFinalResult,min(testtime) testtime,max(TestType) TestType,MAX(Remarks) Remarks,max(TestUser) TestUser from IQC_TestList ";

            sqlstr = sqlstr + sqlwhere2 + sqlwhere1 + "  group by d.deliveryid,d.materialcode ) z order by z.测试日期 desc   ";



          // if (!cb.Checked)
          // {
                DataTable dt = DbAccess.SelectBySql(sqlstr).Tables[0];

                if (dt != null && dt.Rows.Count > 0)
                {
                    databind.DataSource = dt;
                }
                else
                {
                    MessageBox.Show("没有符合条件的记录");
                }
           // }
            //else
            //{
            //    TestListRecord TLR = new TestListRecord("退料测试详细", txttesttype.Text == "" ? "" : txttesttype.Text, txttestitem.Text == "" ? "" : txttestitem.Text, "", txttesttools.Text == "" ? "" : txttesttools.Text, txtlotno.Text, txttestuser.Text, txtmaterialcode.Text,
            //                                          txtmaterialname.Text, txtvendorcode.Text, txtvendorname.Text, sbdate, sedate);
            //    TLR.ShowDialog();
            //}

        }



        private void btnsearch_Click(object sender, EventArgs e)
        {
            string sbdate = "", sedate = "";
            sbdate = bdate.Text;
            sedate = edate.Text;

            if (txtmaterialcode.Text.Trim()=="" && txtmaterialname.Text.Trim()=="" && txtvendorcode.Text.Trim()== ""&& txtvendorname.Text.Trim()== ""
                && bdate.Text.Trim() == "" && edate.Text.Trim() == "" && txtlotno.Text.Trim() == "" && txttestuser.Text.Trim() == "" && txtcusvendor.Text.Trim() == "")
            {
                MessageBox.Show("请输入相关条件查询","提醒",MessageBoxButtons.OK ,MessageBoxIcon.Information);
                return;
            }

            if (!cb.Checked)
            {
                 BackgroundTask.BackgroundWork(selectdata, null);               
            }
            else
            {
                TestListRecord TLR = new TestListRecord("退料测试详细", txttesttype.Text == "" ? "" : txttesttype.Text, txttestitem.Text == "" ? "" : txttestitem.Text, "", txttesttools.Text == "" ? "" : txttesttools.Text, txtlotno.Text, txttestuser.Text, txtmaterialcode.Text,
                                          txtmaterialname.Text, txtvendorcode.Text, txtvendorname.Text, sbdate, sedate);
                TLR.ShowDialog();
            }

            //string sqlstr = "" , sqlwhere1 = " where 1=1", sqlwhere2= " where 1=1";
            //string materialcode = "", materialname = "", vendorcode = "", vendorname = "";
            //string testtype = "", testitem = "", testsubitem = "", testtools = "", testuser = "", lotno = "", testtime1 = "", testtime2 = "";

            //materialcode = txtmaterialcode.Text.Trim();
            //materialname = txtmaterialname.Text.Trim();
            //vendorcode = txtvendorcode.Text.Trim();
            //vendorname = txtvendorname.Text.Trim();


            //testtype = txttesttype.Text.Trim();
            //testitem = txttestitem.Text.Trim();
            //testsubitem = txttestsubitem.Text.Trim();
            //testtools = txttesttools.Text.Trim();
            //testuser = txttestuser.Text.Trim();
            //lotno = txtlotno.Text.Trim();


            //string sbdate = "", sedate = "";        
            //sbdate = bdate.Text ;           
            //sedate = edate.Text ;

            //databind.DataSource = null;
            //DataSet ds = null;

            //if (!string.IsNullOrEmpty(materialcode))
            //{
            //    sqlwhere1 += " and d.materialcode like '%" + materialcode + "%' ";
            //}
            //if (!string.IsNullOrEmpty(materialname))
            //{
            //    sqlwhere1 += "  and materialname like '%" + materialname + "%' ";
            //}
            //if (!string.IsNullOrEmpty(vendorcode))
            //{
            //    sqlwhere1 += " and vendorcode like '%" + vendorcode + "%' ";
            //}
            //if (!string.IsNullOrEmpty(vendorname))
            //{
            //    sqlwhere1 += "  and vendorname like '%" + vendorname + "%' ";
            //}


            //if (!string.IsNullOrEmpty(testtype))
            //{
            //    sqlwhere2 += " and testtype='"+ testtype + "' ";
            //}
            //if (!string.IsNullOrEmpty(testitem))
            //{
            //    sqlwhere2 += " and testitem='" + testitem + "' ";
            //}
            //if (!string.IsNullOrEmpty(testsubitem))
            //{
            //    sqlwhere2 += " and testsubitem='" + testsubitem + "' ";
            //}
            //if (!string.IsNullOrEmpty(testtools))
            //{
            //    sqlwhere2 += "  and TestTools like '%" + testtools + "%' ";
            //}
            //if (!string.IsNullOrEmpty(testuser))
            //{
            //    sqlwhere2 += " and TestUser like '%" + testuser + "%' ";
            //}
            //if (!string.IsNullOrEmpty(lotno))
            //{
            //    sqlwhere2 += "  and Lotno like '%" + lotno + "%' ";
            //}
            //if (!string.IsNullOrEmpty(sbdate))
            //{
            //    sqlwhere2 += " and testtime >= '" + sbdate + " 00:00:00 ' ";
            //}
            //if (!string.IsNullOrEmpty(sedate))
            //{
            //    sqlwhere2 += " and testtime <= '" + sedate + " 23:59:59 ' ";
            //}
            //sqlwhere2 = sqlwhere2 + "   group by receptid,Productcode,Lotno) x on d.deliveryid=x.receptid and d.materialcode=x.Productcode and d.lotno=x.LotNo ";

            //sqlstr = @"  select * from
            //            ( select d.deliveryid 接收单号,d.materialcode 料号,max(materialname) 名称,case when max(testqty)>1 then max(testqty) else sum(cast(qty as bigint)) end 数量,max(vendorcode) 供应商代码,max(vendorname) 供应商名称,min(transactiondate) 生成时间,
            //    case max(TestFinalResult) when '拒收' then 'NG' else 'OK' end 检验结果,
            //             max(testtime) 测试日期,max(TestType) 测试类别,Max(Remarks) 备注,max(TestUser) 检验员,d.materialcode 全尺寸数据 
            // from delivery d left join MaterialSpec m on d.materialcode=m.materialcode  ";
            //sqlstr = sqlstr + "   right join (select receptid,Productcode,LotNo,max(testqty) testqty,max(TestFinalResult) TestFinalResult,min(testtime) testtime,max(TestType) TestType,MAX(Remarks) Remarks,max(TestUser) TestUser from IQC_TestList ";

            //sqlstr = sqlstr + sqlwhere2 + sqlwhere1 + "  group by d.deliveryid,d.materialcode ) z order by z.测试日期 desc   ";



            //if (!cb.Checked)
            //{
            //    DataTable dt = DbAccess.SelectBySql(sqlstr).Tables[0];

            //    if (dt != null && dt.Rows .Count >0 )
            //    {
            //        databind.DataSource = dt;
            //    }
            //    else
            //    {
            //        MessageBox.Show("没有符合条件的记录");
            //    }
            //}
            //else
            //{
            //    TestListRecord TLR = new TestListRecord("退料测试详细", txttesttype.Text == "" ? "" : txttesttype.Text, txttestitem.Text == "" ? "" : txttestitem.Text, "", txttesttools.Text == "" ? "" : txttesttools.Text, txtlotno.Text, txttestuser.Text, txtmaterialcode.Text,
            //                                          txtmaterialname.Text, txtvendorcode.Text, txtvendorname.Text, sbdate, sedate);
            //    TLR.ShowDialog();
            //}


            //if (!cb.Checked)
            //{
            //    ds = ic.SelectTestListRecord("测试记录", txttesttype.SelectedValue == null ? "" : txttesttype.SelectedValue.ToString(), txttestitem.SelectedValue == null ? "" : txttestitem.SelectedValue.ToString(), "", txttesttools.SelectedValue == null ? "" : txttesttools.SelectedValue.ToString(), txtlotno.Text, txttestuser.Text, txtmaterialcode.Text,
            //                                  txtmaterialname.Text, txtvendorcode.Text, txtvendorname.Text, sbdate, sedate);
            //    if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            //    {
            //        databind.DataSource = ds.Tables[0];
            //    }
            //    else
            //    {
            //        MessageBox.Show("没有符合条件的记录");
            //    }
            //}
            //else
            //{
            //    TestListRecord TLR = new TestListRecord("退料测试详细", txttesttype.SelectedValue == null ? "" : txttesttype.SelectedValue.ToString(), txttestitem.SelectedValue == null ? "" : txttestitem.SelectedValue.ToString(), "", txttesttools.SelectedValue == null ? "" : txttesttools.SelectedValue.ToString(), txtlotno.Text, txttestuser.Text, txtmaterialcode.Text,
            //                                          txtmaterialname.Text, txtvendorcode.Text, txtvendorname.Text, sbdate, sedate);
            //    TLR.ShowDialog();
            //}
        }

   

    }
}