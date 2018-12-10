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
using DX_QMS.Common;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraEditors;
using System.Text.RegularExpressions;
using System.Collections;
using DevExpress.XtraGrid;
using DevExpress.Data;
using DevExpress.XtraGrid.Menu;
using DevExpress.Utils.Menu;

namespace DX_QMS
{
    public partial class OQCTestListSearch : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private GridGroupSummaryItemCollection  gsiMultiSummary;
        DXPopupMenu formatRulesMenu = new DXPopupMenu();
        public OQCTestListSearch()
        {
            InitializeComponent();
            gridView.PrintExportProgress += new ProgressChangedEventHandler(gridView_PrintExportProgress);
        }
        private void bindorg_id()
        {
            string sql = "select ORG_ID,ORG_NAME from ORG_INFO where ORG_TYPE='ORG' order by SORT ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            txtorg_id.Properties.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
                txtorg_id.Properties.Items.Add(row["ORG_NAME"]);
            }
            txtorg_id.SelectedIndex = 0;

        }
        private void OQCTestListSearch_Load(object sender, EventArgs e)
        {
            bindorg_id();
            setRule();
            sBtndetele.Enabled = true;
            InitSummaries();
        }

        private void setRule()
        {
            string post = "";
            if (Login.manager != "")
            {
                post = Login.manager;
            }
            else
            {
                post = Login.post;
            }
            Dictionary<string, bool> dic = GroupPermission.QMS_SelectRulesForForm(post, "OQC检验记录");
            this.sBtndetele.Enabled = bool.Parse(dic["hasDelete"].ToString());           
        }


        private void InitSummaries()
        {
            gsiMultiSummary = new GridGroupSummaryItemCollection(gridView);
            /////gsiMultiSummary.Add(SummaryItemType.Count, "客户");
           ///// gsiMultiSummary.Add(SummaryItemType.Count, "工单号");
           ///// gsiMultiSummary.Add(SummaryItemType.Count, "机型");
            ///// gsiMultiSummary.Add(SummaryItemType.Average, "工单号", null,"");
        }

        private void btnselect_Click(object sender, EventArgs e)
        {
            string where = " where 1=1 ";
            string org_id = txtorg_id.Text.Trim();
            string customer = txtcustomer.Text.Trim();
            string workno = txtworkno.Text.Trim();
            string PN = txtPN.Text.Trim();
            string Serialnumber = txtSerialnumber.Text.Trim();
            string begintime = begindate.Text;
            string endtime = enddate.Text;

            if (!string.IsNullOrEmpty(org_id))
            {
                where += " and org_id = '" + org_id + "' ";
            }
            if (!string.IsNullOrEmpty(customer))
            {
                where += " and customer = '" + customer + "' ";
            }
            if (!string.IsNullOrEmpty(workno))
            {
                where += " and workno = '" + workno + "' ";
            }
            if (!string.IsNullOrEmpty(PN))
            {
                where += " and hytcode = '" + PN + "' ";
            }
            if (!string.IsNullOrEmpty(Serialnumber))
            {
                where += " and serialnumber = '" + Serialnumber + "' ";
            }
            if (!string.IsNullOrEmpty(begintime))
            {
                where += " and checkdate >= '" + begintime + " 00:00:00 '";
            }
            if (!string.IsNullOrEmpty(endtime))
            {
                where += " and checkdate <='" + endtime + " 23:59:59 '";
            }

            string sql = @" select  checkdate 检查时间,customer 客户,model 机型,hytcode 编码,workno 工单号,serialnumber 抽样流水号,
                            sendqty 送检批量,org_id 组织,sampleqty 应抽数量,factsampleqty 实际抽样数量,sampleplan 抽样计划,checktype 检验方式,CartonNo 箱号,productionphase 生产阶段,productstate 产品状态,
                            NGQty NG数量,testresult 检查结果,testman 检验人,testremark 检验备注,latyper 拉长,lineid 线体,QC QC,QE QE工程师,masters 责任主管 ,badinformation 不良信息 from OQC_TestListNew  ";
            sql += where + " order by checkdate desc ";

            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];

            if (dt != null && dt.Rows.Count > 0)
            {


                gridControl.DataSource = dt;
                gridColumn2.GroupIndex = 0;
               ////// gridColumn5.GroupIndex = 1;
                gridView.ExpandAllGroups();

               //// gridView.GroupSummary.Assign(gsiMultiSummary);


            }
            else
            {
                MessageBox.Show("没有符合条件的记录");
                gridControl.DataSource = null;
            }

        }

    private void btnreset_Click(object sender, EventArgs e)
        {
           txtorg_id.Text ="";
           txtcustomer.Text = "";
           txtworkno.Text = "";
           txtPN.Text = "";
           txtSerialnumber.Text = "";
           begindate.Text = "";
           enddate.Text = "";
           gridColumn2.GroupIndex = -1;
          //// gridColumn5.GroupIndex = -1;
           gridControl.DataSource = null;

        }


        private string ShowSaveFileDialog(string title, string filter)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            string name = "OQC测试记录信息";
            int n = name.LastIndexOf(".") + 1;
            if (n > 0) name = name.Substring(n, name.Length - n);
            dlg.Title = "导出 " + title;
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
        private void btnexport_Click(object sender, EventArgs e)
        {
            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null)
                return;
            if (dt.Rows.Count <= 0) return;
            string fileName = ShowSaveFileDialog("Microsoft Excel 2007 Document", "Microsoft Excel|*.xlsx");
            if (fileName == string.Empty) return;
            ExportToEx(fileName, "xlsx", gridView);
            OpenFile(fileName);
        }

        void export(string Serialnumber)
        {
            string teststandard = "",SNsql= "";
            int sheetCount = 1;
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = true;
            object missing = System.Reflection.Missing.Value;
            string templetFile = Environment.CurrentDirectory + @"\ReportFolder\OQCreport.xlsx";
            Microsoft.Office.Interop.Excel.Workbook workBook = app.Workbooks.Open(templetFile, missing, true, missing, missing, missing,
                                                          missing, missing, missing, missing, missing, missing, missing, missing, missing);
            Microsoft.Office.Interop.Excel.Worksheet workSheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.Sheets.get_Item(1);

            for (int i = 1; i < sheetCount; i++)
            {
                ((Microsoft.Office.Interop.Excel.Worksheet)workBook.Worksheets.get_Item(i)).Copy(missing, workBook.Worksheets[i]);
            }
            Microsoft.Office.Interop.Excel.Worksheet sheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.Worksheets.get_Item(1);
            if (sheet == null)
                return;


            string testlistsql = @"select * from OQC_TestListNew where serialnumber='"+ Serialnumber + "'";
            DataTable testdt = DbAccess.SelectBySql(testlistsql).Tables[0];
            if (testdt == null || testdt.Rows.Count < 1)
                return;
            string customer = "", PN = "", model="", productionphase= "", deliveryID="", sendqty="", sampleqty="", NGQty="", checkdate="", item="",ECN = "", worknosendqty= "", testtestresult="", testremark = "", testman= "",QE = "";
            string TestListsampleplan = "";
            customer = testdt.Rows[0]["customer"].ToString();
            PN = testdt.Rows[0]["hytcode"].ToString();
            model = testdt.Rows[0]["model"].ToString();
            productionphase = testdt.Rows[0]["productionphase"].ToString();
            deliveryID = testdt.Rows[0]["deliveryID"].ToString();
            sendqty = testdt.Rows[0]["sendqty"].ToString();
            sampleqty = testdt.Rows[0]["sampleqty"].ToString();
            NGQty = testdt.Rows[0]["NGQty"].ToString();
            checkdate = testdt.Rows[0]["checkdate"].ToString();
            item= testdt.Rows[0]["item"].ToString();
            ECN = testdt.Rows[0]["ECNnumber"].ToString();
            worknosendqty = testdt.Rows[0]["sendqty"].ToString();
            testtestresult = testdt.Rows[0]["testresult"].ToString();
            teststandard = testdt.Rows[0]["teststandard"].ToString();
            testremark = testdt.Rows[0]["testremark"].ToString();
            testman = testdt.Rows[0]["testman"].ToString();
            QE = testdt.Rows[0]["QE"].ToString();
            TestListsampleplan = testdt.Rows[0]["sampleplan"].ToString();

            sheet.Cells.get_Range("C4").Value = customer;
            sheet.Cells.get_Range("C5").Value = model;
            sheet.Cells.get_Range("C6").Value = PN;
            sheet.Cells.get_Range("C7").Value = ECN;
            sheet.Cells.get_Range("C8").Value = Serialnumber;
            sheet.Cells.get_Range("S3").Value = deliveryID;
            sheet.Cells.get_Range("S4").Value = sendqty;
            sheet.Cells.get_Range("S5").Value = sampleqty;
            sheet.Cells.get_Range("S6").Value = NGQty;
            sheet.Cells.get_Range("S7").Value = checkdate.Substring(0, 9);
            sheet.Cells.get_Range("A43").Value = "出货明细：工单号数量为：" + worknosendqty;
            sheet.Cells.get_Range("Q38").Value = "备注："+ testremark;
            sheet.Cells.get_Range("C44").Value = testman;
            sheet.Cells.get_Range("L44").Value = QE;
            sheet.Cells.get_Range("R44").Value = DateTime.Now.ToString("yyyy-MM-dd");
            if (productionphase.Contains("样机/试产"))
            {
                sheet.Cells.get_Range("I5").Value = "√样机/试产";
            }
            else if (productionphase.Contains("工程变更"))
            {
                sheet.Cells.get_Range("I6").Value = "√工程变更";
            }
            else if (productionphase.Contains("正常量产"))
            {
                sheet.Cells.get_Range("I7").Value = "√正常量产";
            }
            if (testtestresult.Contains("OK"))
            {
                sheet.Cells.get_Range("N43").Value = "√允收";
            }
            if (testtestresult.Contains("NG"))
            {
                sheet.Cells.get_Range("O43").Value = "√退货";
            }

            SNsql = @" select * from OQC_SampleNewList where items ='"+item+ "'";
            DataTable SNtestdt = DbAccess.SelectBySql(SNsql).Tables[0];
            if (SNtestdt != null && SNtestdt.Rows.Count > 0)
            {
                string SNbaddescribe = "";
                int SNbadcount = 1;
                for (int n=0;n< SNtestdt.Rows.Count;n++)   //// SNtestdt.Rows[n]["SNnumber"].ToString() != "" 
                {
                    if ( SNtestdt.Rows[n]["badnumber"].ToString() !="" && SNtestdt.Rows[n]["badclass"].ToString() !="" && SNtestdt.Rows[n]["baddescribe"].ToString() !="")
                    {
                        SNbaddescribe = SNbaddescribe +"["+SNbadcount.ToString()+ "]【" + SNtestdt.Rows[n]["SNnumber"].ToString() + "；" + SNtestdt.Rows[n]["badnumber"].ToString() + "；" + SNtestdt.Rows[n]["badclass"].ToString() + "；" + SNtestdt.Rows[n]["baddescribe"].ToString() + "】"+ "\r\n";
                        SNbadcount = SNbadcount + 1;
                    }
                }

                sheet.Cells.get_Range("Q11").Value = SNbaddescribe;
            }
            else
            {
                sheet.Cells.get_Range("Q11").Value = "NA";
            }

            string testitem = "", checkmethod = "", checkMA = "否", checkMI = "否";
            string sampleplan = "", MA = "", MI = "", IPCA610F = "", customerstandard = "", otherstandard = "";
            double MAvalue, MIvalue;
            int count = 0;
            string sql = "";
            sql = @" select * from OQCTestProgSet where customer ='" + customer + "' and PN='" + PN + "' order by testitem ,standardsequence ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            if (dt != null && dt.Rows.Count > 0)
            {
                count = dt.Rows.Count;
                sampleplan = dt.Rows[0]["sampleplan"].ToString();
                MA = dt.Rows[0]["MA"].ToString();
                MI = dt.Rows[0]["MI"].ToString();
                IPCA610F = dt.Rows[0]["IPCA610F"].ToString();
                customerstandard = dt.Rows[0]["customerstandard"].ToString();
                otherstandard = dt.Rows[0]["otherstandard"].ToString();
                for (int n = 0, m = 11; n < dt.Rows.Count; n++, m++)
                {
                    sheet.Cells[m, 1] = dt.Rows[n]["testitem"].ToString();
                    sheet.Cells[m, 3] = dt.Rows[n]["standardsequence"].ToString();
                    sheet.Cells[m, 4] = dt.Rows[n]["teststandard"].ToString();
                    sheet.Cells[m, 13] = dt.Rows[n]["checkmethod"].ToString();

                    string Testitem = "", Testnumber = "", Testitemresult = "";
                    string[] resultString = Regex.Split(teststandard, "；", RegexOptions.IgnoreCase);
                    foreach (string testresult in resultString)
                    {
                        string[] result = Regex.Split(testresult, "，", RegexOptions.IgnoreCase);
                        if (result.Length == 3)
                        {
                            Testitem = result[0].Trim();
                            Testnumber = result[1].Trim();
                            Testitemresult = result[2].Trim();

                            if (dt.Rows[n]["testitem"].ToString() == Testitem && dt.Rows[n]["standardsequence"].ToString() == Testnumber)
                            {
                                if (Testitemresult.Contains("OK"))
                                {
                                    sheet.Cells[m, 15] = Testitemresult;
                                }
                                else
                                {
                                    sheet.Cells[m, 16] = Testitemresult;
                                }

                            }
                        }                        
                    }

                }

            }
            else
            {
                sql = @" select * from OQCTestProgSet where customer ='" + customer + "' and PN=''  order by testitem ,standardsequence ";
                DataTable dts = DbAccess.SelectBySql(sql).Tables[0];
                if (dts != null && dts.Rows.Count > 0)
                {
                    count = dts.Rows.Count;
                    sampleplan = dts.Rows[0]["sampleplan"].ToString();
                    MA = dts.Rows[0]["MA"].ToString();
                    MI = dts.Rows[0]["MI"].ToString();
                    IPCA610F = dts.Rows[0]["IPCA610F"].ToString();
                    customerstandard = dts.Rows[0]["customerstandard"].ToString();
                    otherstandard = dts.Rows[0]["otherstandard"].ToString();
                    for (int n = 0, m = 11; n < dts.Rows.Count; n++, m++)
                    {
                        sheet.Cells[m, 1] = dts.Rows[n]["testitem"].ToString();
                        sheet.Cells[m, 3] = dts.Rows[n]["standardsequence"].ToString();
                        sheet.Cells[m, 4] = dts.Rows[n]["teststandard"].ToString();
                        sheet.Cells[m, 13] = dts.Rows[n]["checkmethod"].ToString();
                        string Testitem = "", Testnumber = "", Testitemresult = "";
                        string[] resultString = Regex.Split(teststandard, "；", RegexOptions.IgnoreCase);
                        foreach (string testresult in resultString)
                        {
                            string[] result = Regex.Split(testresult, "，", RegexOptions.IgnoreCase);
                            if (result.Length == 3)
                            {
                                Testitem = result[0].Trim();
                                Testnumber = result[1].Trim();
                                Testitemresult = result[2].Trim();

                                if (dts.Rows[n]["testitem"].ToString() == Testitem && dts.Rows[n]["standardsequence"].ToString() == Testnumber)
                                {
                                    if (Testitemresult.Contains("OK"))
                                    {
                                        sheet.Cells[m, 15] = Testitemresult;
                                    }
                                    else
                                    {
                                        sheet.Cells[m, 16] = Testitemresult;
                                    }

                                }
                            }
                        }


                    }
                }
                else
                {
                    sql = @" select * from OQCTestProgSet where customer ='' and PN='' order by testitem ,standardsequence ";
                    DataTable dss = DbAccess.SelectBySql(sql).Tables[0];
                    if (dss != null && dss.Rows.Count > 0)
                    {
                        count = dss.Rows.Count;
                        sampleplan = dss.Rows[0]["sampleplan"].ToString();
                        MA = dss.Rows[0]["MA"].ToString();
                        MI = dss.Rows[0]["MI"].ToString();
                        IPCA610F = dss.Rows[0]["IPCA610F"].ToString();
                        customerstandard = dss.Rows[0]["customerstandard"].ToString();
                        otherstandard = dss.Rows[0]["otherstandard"].ToString();
                        for (int n = 0, m = 11; n < dss.Rows.Count; n++, m++)
                        {
                            sheet.Cells[m, 1] = dss.Rows[n]["testitem"].ToString().Trim ();
                            sheet.Cells[m, 3] = dss.Rows[n]["standardsequence"].ToString();
                            sheet.Cells[m, 4] = dss.Rows[n]["teststandard"].ToString().Trim();
                            sheet.Cells[m, 13] = dss.Rows[n]["checkmethod"].ToString();
                            string Testitem = "", Testnumber = "", Testitemresult = "";
                            string[] resultString = Regex.Split(teststandard, "；", RegexOptions.IgnoreCase);
                            foreach (string testresult in resultString)
                            {
                                string[] result = Regex.Split(testresult, "，", RegexOptions.IgnoreCase);
                                if (result.Length == 3)
                                {
                                    Testitem = result[0].Trim();
                                    Testnumber = result[1].Trim();
                                    Testitemresult = result[2].Trim();

                                    if (dss.Rows[n]["testitem"].ToString() == Testitem && dss.Rows[n]["standardsequence"].ToString() == Testnumber)
                                    {
                                        if (Testitemresult.Contains("OK"))
                                        {
                                            sheet.Cells[m, 15] = Testitemresult;
                                        }
                                        else
                                        {
                                            sheet.Cells[m, 16] = Testitemresult;
                                        }

                                    }
                                }
                            }
                        }

                    }
                    else
                    {
                        MessageBox.Show("没有维护检验项目", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }

            {
                if (TestListsampleplan == "C=0")
                    sheet.Cells.get_Range("N3").Value = "√C=0";
                else if (TestListsampleplan == "ISO2859-1")
                    sheet.Cells.get_Range("O3").Value = "√ISO2859-1一般II级";
                else
                    sheet.Cells.get_Range("Q3").Value = "√全检";

                if (MA == "AQL=0.65")
                    sheet.Cells.get_Range("N4").Value = "√0.65";
                else if (MA == "AQL=0.40")
                    sheet.Cells.get_Range("O4").Value = "√0.4";
                else if (MA == "AQL=0.01")
                    sheet.Cells.get_Range("P4").Value = "√0.01";
                else
                    sheet.Cells.get_Range("Q4").Value = "√其他";

                if (MI == "AQL=0.65")
                    sheet.Cells.get_Range("N5").Value = "√0.65";
                else if (MI == "AQL=0.40")
                    sheet.Cells.get_Range("O5").Value = "√0.4";
                else if (MI == "AQL=0.01")
                    sheet.Cells.get_Range("P5").Value = "√0.01";
                else
                    sheet.Cells.get_Range("Q5").Value = "√其他";

                if (IPCA610F == "I级")
                    sheet.Cells.get_Range("O6").Value = "√I级";
                if (IPCA610F == "II级")
                    sheet.Cells.get_Range("P6").Value = "√II级";
                if (IPCA610F == "III级")
                    sheet.Cells.get_Range("Q6").Value = "√III级";
                if (customerstandard == "客户检验标准规范")
                    sheet.Cells.get_Range("N7").Value = "√客户检验标准规范";
                if (otherstandard == "其他")
                    sheet.Cells.get_Range("N8").Value = "√其他";

            }

        }


        void exportF1(string Serialnumber)
        {
            string sql = "", item = "";
            int sheetCount = 1;
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = true;
            object missing = System.Reflection.Missing.Value;
            string templetFile = Environment.CurrentDirectory + @"\ReportFolder\OQCF1report.xlsx";
            Microsoft.Office.Interop.Excel.Workbook workBook = app.Workbooks.Open(templetFile, missing, true, missing, missing, missing,
                                                          missing, missing, missing, missing, missing, missing, missing, missing, missing);
            Microsoft.Office.Interop.Excel.Worksheet workSheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.Sheets.get_Item(1);

            for (int i = 1; i < sheetCount; i++)
            {
                ((Microsoft.Office.Interop.Excel.Worksheet)workBook.Worksheets.get_Item(i)).Copy(missing, workBook.Worksheets[i]);
            }
            Microsoft.Office.Interop.Excel.Worksheet sheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.Worksheets.get_Item(1);
            if (sheet == null)
                return;

            sql = @" select item, checkdate, model,sendqty,sampleqty from OQC_TestListNew where serialnumber='" + Serialnumber + "'";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            if (dt == null || dt.Rows.Count < 0)
                return;

            sheet.Cells.get_Range("O3").Value = dt.Rows[0]["checkdate"].ToString();
            sheet.Cells.get_Range("F62").Value = dt.Rows[0]["checkdate"].ToString();
            sheet.Cells.get_Range("P62").Value = dt.Rows[0]["checkdate"].ToString();
            sheet.Cells.get_Range("B5").Value = dt.Rows[0]["model"].ToString();
            sheet.Cells.get_Range("O4").Value = dt.Rows[0]["sendqty"].ToString();
            sheet.Cells.get_Range("O5").Value = dt.Rows[0]["sampleqty"].ToString();


            item = dt.Rows[0]["item"].ToString();

            sql = @" select SNnumber ,max(CartonNo) CartonNo, max(VersionNO) VersionNO from OQC_SampleNewList where items = '" + item + "' group by SNnumber order by SNnumber asc  ";

           // sql = @" select SNnumber , CartonNo ,VersionNO from OQC_SampleNewList where items = '20180211111146803'   ";
            DataTable itemdt = DbAccess.SelectBySql(sql).Tables[0];
            if (itemdt == null && itemdt.Rows.Count < 0)
                return;
            sheet.Cells.get_Range("B3").Value = itemdt.Rows [0]["CartonNo"].ToString();
            sheet.Cells.get_Range("B4").Value = "";
            sheet.Cells.get_Range("F5").Value = itemdt.Rows[0]["VersionNO"].ToString();

            int n = itemdt.Rows.Count / 50;
            int a = (int)'A';
            for (int i = 1,j=1,m= 8; i <= itemdt.Rows.Count; i++)
            {                        
                //sheet.Cells[m, j] = i;
                sheet.Cells.get_Range((char)a + m.ToString()).Value = i;
               // sheet.Cells.get_Range((char)a+m.ToString()).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Yellow);    
                         
               // sheet.Cells[m, j + 1] = itemdt.Rows[i-1]["SNnumber"].ToString().Trim();
                sheet.Cells.get_Range((char)(a + 1) + m.ToString()).Value= itemdt.Rows[i - 1]["SNnumber"].ToString().Trim();
               // sheet.Cells.get_Range((char)(a+1)+ m.ToString()).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Yellow);

               // sheet.Cells[m, j + 2] = itemdt.Rows[i-1]["CartonNo"].ToString().Trim();
                sheet.Cells.get_Range((char)(a + 2) + m.ToString()).Value = itemdt.Rows[i - 1]["CartonNo"].ToString().Trim();
               // sheet.Cells.get_Range((char)(a + 2)+ m.ToString()).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Yellow);

                //sheet.Cells[m, j + 3] = "ACC";
                sheet.Cells.get_Range((char)(a + 3) + m.ToString()).Value = "ACC";
               // sheet.Cells.get_Range((char)(a + 3)+ m.ToString()).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Yellow);

                m++;
                if (i % 50 == 0  && i != 0)
                {
                    m = 8;
                    a = a + 4;
                    j= j+4;
                }

            }

        }


        private string NunToChar(int number)
        {
            if (65 <= number && 90 >= number)
            {
                System.Text.ASCIIEncoding asciiEncoding = new System.Text.ASCIIEncoding();
                byte[] btNumber = new byte[] { (byte)number };
                return asciiEncoding.GetString(btNumber);
            }
            return "数字不在转换范围内";
        }


        private string NunberToChar(int number)
        {
            if (1 <= number && 36 >= number)
            {
                int num = number + 64;
                System.Text.ASCIIEncoding asciiEncoding = new System.Text.ASCIIEncoding();
                byte[] btNumber = new byte[] { (byte)num };
                return asciiEncoding.GetString(btNumber);
            }
            return "数字不在转换范围内";
        }


        private void gridView_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
        {
            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
            {
                return;
            }
            if (gridView.GetDataRow(e.RowHandle)["检查结果"].ToString() == "NG")
            {
                e.Appearance.BackColor = Color.Red;
            }
        }

        private void gridView_Click(object sender, EventArgs e)
        {
            // gridView.GetFocusedRowCellValue("产品编码").ToString();
        }



        private void gridView_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {
            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
            {
                return;
            }
            try
            {
                txtcustomer.Text = gridView.GetFocusedRowCellValue("客户").ToString();
                txtworkno.Text = gridView.GetFocusedRowCellValue("工单号").ToString();
                txtPN.Text = gridView.GetFocusedRowCellValue("编码").ToString();
            }
            catch
            {

            }
        }

        private void gridView_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {
            string Serialnumber, customer, item = "", sql = "";
            if (e.Column.FieldName == "抽样流水号")
            {
                Serialnumber = gridView.GetFocusedRowCellValue("抽样流水号").ToString();
                customer = gridView.GetFocusedRowCellValue("客户").ToString();
                try
                {
                    //if (customer.Contains("F1"))
                    //{
                    //    exportF1(Serialnumber);
                    //}
                    //else
                    //{
                    //    export(Serialnumber);
                    //}

                    export(Serialnumber);

                }
                catch (Exception ex)
                {
                   
                   MessageBox.Show("还没生成完整报表", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else if (e.Column.FieldName == "实际抽样数量")
            {

                Serialnumber = gridView.GetFocusedRowCellValue("抽样流水号").ToString();
                sql = @" select item from OQC_TestListNew where serialnumber = '" + Serialnumber + "' ";
                DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
                if (dt != null && dt.Rows.Count > 0)
                {
                    item = dt.Rows[0]["item"].ToString();
                    OQCSampleList OL = new OQCSampleList(item);
                    DialogResult dr = OL.ShowDialog();
                }
                else
                {
                    MessageBox.Show("没有抽样数据！", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }

            else if (e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                GridFormatRuleMenuItems items = new GridFormatRuleMenuItems(gridView, e.Column, formatRulesMenu.Items);
                if (items.Count > 0)
                    MenuManagerHelper.ShowMenu(formatRulesMenu, gridControl.LookAndFeel, gridControl.MenuManager, gridControl, new Point(e.X, e.Y));
            }

        }

        private void sBtndetele_Click(object sender, EventArgs e)
        {

            int i = gridView.FocusedRowHandle;
            if (i < 0)
            {
                MessageBox.Show("请选中需要删除的信息", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
            {
                MessageBox.Show("没有数据可删除", "停止", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            if (detelereason.Text.Trim()=="")
            {
                MessageBox.Show("请输入删除原因", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            string sql = "", item = "";
            ArrayList list = new ArrayList() ;
            list.Clear();
            string serialnumber = gridView.GetFocusedRowCellValue("抽样流水号").ToString();
            string org_id = gridView.GetFocusedRowCellValue("组织").ToString();
            string workno = gridView.GetFocusedRowCellValue("工单号").ToString();

            sql = @"  select item from OQC_TestListNew where serialnumber='" + serialnumber + "'";
            DataTable itemdt = DbAccess.SelectBySql(sql).Tables[0];
            if (itemdt == null || itemdt.Rows.Count < 1)
            {
                MessageBox.Show("删除失败", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else
            {
                item = itemdt.Rows[0]["item"].ToString();
            }
            sql = " insert into OQC_TestListNewBackup select  ";
            sql += " checkdate,item,customer,cuscode,model,deliveryID,hytcode,workno,serialnumber,sendqty,org_id,sampleqty,sampleplan,MA,MI,";
            sql += " checktype,allowqty,lineid,latyper,QC,masters,QE,CartonNo,productionphase,ECNnumber,factsampleqty,NGQty,NGpoint,rsno,productstate,teststandard,badinformation,testremark,testresult,testman,checkman,Auditman, ";
            sql += " CauseAnalysis,improvemeasures  from OQC_TestListNew where  org_id='" + org_id + "' and workno='" + workno + "' and item='" + item + "' " ;
            list.Add(sql);
            sql = @" delete OQC_SampleNewList where items='"+ item + "' and org_id='"+ org_id + "' and workno='"+workno+ "' ";
            list.Add(sql);
            sql = @" delete OQC_TestListNew where org_id='"+ org_id + "' and workno='"+ workno + "' and item='"+ item + "' ";
            list.Add(sql);
            sql = @" update OQC_TestListNewBackup set testremark = '"+ detelereason.Text + "' where org_id='" + org_id + "' and workno='" + workno + "' and item='" + item + "' ";
            list.Add(sql);

            bool flag = false;
            try
            {
                flag = DbAccess.ExecutSqlTran(list);
            }
            catch
            {
                MessageBox.Show("删除失败", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (flag == true )
            {
                MessageBox.Show("删除成功", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                gridControl.DataSource = null;
            }
            else
            {
                MessageBox.Show("删除失败", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}