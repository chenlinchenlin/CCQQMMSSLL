using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Windows.Forms;
using DevExpress.XtraBars;
using DX_QMS.Common;
using DevExpress.XtraCharts;
using DevExpress.Utils;
using DevExpress.XtraPrinting;
using DevExpress.XtraPrintingLinks;
using System.IO;
using DevExpress.XtraGrid.Views.Grid;
using System.Diagnostics;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraEditors;

namespace DX_QMS.IPQC
{
    public partial class IPQCExceptionReportIPQC : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private string serverFilePath = "192.168.0.204\\FilePath$";
        public IPQCExceptionReportIPQC()
        {
            InitializeComponent();
        }

        private void IPQCExceptionReport_Load(object sender, EventArgs e)
        {
            txtstartdate.Text = DateTime.Now.ToString("yyyy-MM-dd");
            txtenddate.Text = DateTime.Now.ToString("yyyy-MM-dd");
            selecttype.SelectedIndex = 0;
            selecttype_SelectedIndexChanged(null, null);
        }

        void bindCucustomer()
        {
            DataTable dt = null;
            string sql = "   select distinct customer from IPQCExceptionList where checkdate >='"+txtstartdate.DateTime.ToString("yyyy-MM-dd")+ " 00:00:00'  and  checkdate <='"+txtenddate.DateTime.ToString("yyyy-MM-dd")+" 23:59:59'  ";
            dt = DbAccess.SelectBySql(sql).Tables[0];     
            customercheck.Properties.Items.Clear();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                customercheck.Properties.Items.Add(dt.Rows[i]["customer"].ToString());
            }
        }

        void bindbadclass()
        {
            DataTable dt = null;
            string sql = "   select distinct badclass from IPQCExceptionList where checkdate >='" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00'  and  checkdate <='" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59' and badclass<>''   ";
            dt = DbAccess.SelectBySql(sql).Tables[0];
            checkbadclass.Properties.Items.Clear();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                checkbadclass.Properties.Items.Add(dt.Rows[i]["badclass"].ToString());
            }
        }
        private void txtstartdate_EditValueChanged(object sender, EventArgs e)
        {
            bindCucustomer();
            bindbadclass();
        }

        private void selecttype_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (selecttype.SelectedIndex == 0)
            {
                txtworkno.Text = "";
                txtworkno.Enabled = true;
                txtworkno.Focus();
                txthytcode.Text = "";
                txthytcode.Enabled = false;
                customercheck.Properties.Items.Clear();
                customercheck.Enabled = false;
                checkbadclass.Properties.Items.Clear();
                checkbadclass.Enabled = false;
            }
            else if (selecttype.SelectedIndex == 1)
            {
                txtworkno.Text = "";
                txtworkno.Enabled = false;
                txthytcode.Text = "";
                txthytcode.Enabled = true;
                txthytcode.Focus();
                customercheck.Properties.Items.Clear();
                customercheck.Enabled = false;
                checkbadclass.Properties.Items.Clear();
                checkbadclass.Enabled = false;
            }
            else if (selecttype.SelectedIndex == 2)
            {
                txtworkno.Text = "";
                txtworkno.Enabled =false ;
                txthytcode.Text = "";
                txthytcode.Enabled = false;         
                customercheck.Properties.Items.Clear();
                customercheck.Enabled = true;
                bindCucustomer();
                checkbadclass.Properties.Items.Clear();
                checkbadclass.Enabled = false;

            }
            else if (selecttype.SelectedIndex == 3)
            {
                txtworkno.Text = "";
                txtworkno.Enabled = false;
                txthytcode.Text = "";
                txthytcode.Enabled = false;
                customercheck.Properties.Items.Clear();
                customercheck.Enabled = false;             
                checkbadclass.Properties.Items.Clear();
                checkbadclass.Enabled = true;
                bindbadclass();
            }
        }
        DataTable querydetail()
        {
            string where = " where 1=1 ";
            string workno = "";
            string[] worknoArray = txtworkno.Text.Trim().Split(new char[4] { ';', '；', ',', '，' });
            for (int i = 0; i < worknoArray.Length; i++)
            {
                workno += ",'" + worknoArray[i].ToString() + "'";
            }
            workno = workno.TrimStart(',');

            string hytcode = "";
            string[] hytcodeArray = txthytcode.Text.Trim().Split(new char[4] { ';', '；', ',', '，' });
            for (int i = 0; i < hytcodeArray.Length; i++)
            {
                hytcode += ",'" + hytcodeArray[i].ToString() + "'";
            }
            hytcode = hytcode.TrimStart(',');

            string customer = "";
            string[] customerArray = customercheck.Text.Split(',');
            for (int i = 0; i < customerArray.Length; i++)
            {
                customer += ",'"+customerArray[i].ToString().Trim()+"'";
            }
            customer = customer.TrimStart(',');

            string badclass = "";
            string[] badclassArray = checkbadclass.Text.Split(',');
            for (int i = 0; i < badclassArray.Length; i++)
            {
                badclass += ",'" + badclassArray[i].ToString().Trim()+ "'";
            }
            badclass = badclass.TrimStart(',');

            string begintime = txtstartdate.DateTime.ToString("yyyy-MM-dd");
            string endtime = txtenddate.DateTime.ToString("yyyy-MM-dd");

            if (selecttype.SelectedIndex == 0)
            {
                where += " and workno in (" + workno + ") ";
            }
            if (selecttype.SelectedIndex == 2)
            {
                where += " and customer in (" + customer + ") ";
            }
            if (selecttype.SelectedIndex == 1)
            {
                where += " and hytcode in (" + hytcode + ") ";
            }
            if (selecttype.SelectedIndex == 3)
            {
                where += " and badclass in (" + badclass + ")  ";
            }
            if (!string.IsNullOrEmpty(begintime))
            {
                where += " and checkdate >= '" + begintime + " 00:00:00 '";
            }
            if (!string.IsNullOrEmpty(endtime))
            {
                where += " and checkdate <='" + endtime + " 23:59:59 '";
            }

            string sql = @"  select  ifrepeat 是否重复发生,standid 站别,item 表单编号,checkdate 日期,productlineid 生产线别,customer 客户,org_id 组织,
		        workno 工单号,worknoqty 工单数量,hytcode Hytera编码,model 客户机型,qty 投入数,NGQty 不良数,CASE WHEN NGQty < worknoqty  then  convert( varchar,convert(numeric(3,1),(NGQty+0.0)/(isnull(worknoqty,qty))*100 ))+'%' ELSE '100%' end 不良率,
				case when  worknoqty >200 and convert(numeric(4,2),(NGQty+0.0)/(worknoqty)) >0.1 then '是' when  worknoqty <= 200  and NGQty>20  then '是'  else  '否'  end 是否批量异常,
                badclass 不良类别,baddescribe 问题描述,problemtype 问题分类,temporaryhandle 临时处理方法,CauseAnalysis 原因分析,improvemeasures 改善计划,
				dutyDepartment 责任部门,updateMan 记录人,chargeMan 责任人,ifClose 是否关闭,ifOvertime 是否超期,overdutyDepartment 延时责任部门,badImage 不良图片
		        from IPQCExceptionList  ";
            sql += where;

            if (batchexception.Checked == true)
            {
                sql += "  and  (( worknoqty >200 and convert(numeric(4,2),(NGQty+0.0)/(worknoqty)) >0.1 ) or ( worknoqty <= 200  and NGQty>20 ))   ";
            }
            sql += "  order by checkdate desc ";

            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            return dt;
        }

        private void sBtnselect_Click(object sender, EventArgs e)
        {
            if (txtstartdate.Text == "" || txtenddate.Text == "")
                return;
          
            if (selecttype.SelectedIndex == 0)
            {
                if (txtworkno.Text.Trim() == "")
                    return;

                gridControl.DataSource = null;
                chartControlBar.DataSource = null;
                            
                gridControl.DataSource = querydetail();              
                chartControlBar.DataSource = WorknoReport();

                ChartTitle chartTitle = new ChartTitle();
                chartTitle.Text = txtstartdate.DateTime.ToString("yyyy-MM-dd") + "至" + txtenddate.DateTime.ToString("yyyy-MM-dd") + txtworkno.Text + " 工单 报表";
                chartTitle.TextColor = System.Drawing.Color.Black; //颜色设置
                chartTitle.Font = new Font("微软雅黑", 10);
                chartTitle.Alignment = StringAlignment.Center;
                chartControlBar.Titles.Clear();
                chartControlBar.Titles.Add(chartTitle);

                XYDiagram diagram = chartControlBar.Diagram as XYDiagram;
                chartControlBar.SeriesTemplate.ArgumentDataMember = "类型";
                diagram.AxisX.Title.Visibility = DefaultBoolean.True;
                diagram.AxisX.Title.Text = "工单号";

                chartControlBar.Series["批数"].ArgumentDataMember = "类型";
                chartControlBar.Series["批数"].ArgumentScaleType = ScaleType.Qualitative;
                chartControlBar.Series["批数"].ValueDataMembers[0] = "批数";
                chartControlBar.Series["批数"].LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
                chartControlBar.Series["批数"].LegendText = "批数";

            }
            else  if (selecttype.SelectedIndex == 1)
            {
                if (txthytcode.Text.Trim() == "")
                    return;

                gridControl.DataSource = null;
                chartControlBar.DataSource = null;
                        
                gridControl.DataSource = querydetail();
               
                chartControlBar.DataSource = HytcodeReport();


                ChartTitle chartTitle = new ChartTitle();
                chartTitle.Text = txtstartdate.DateTime.ToString("yyyy-MM-dd") + "至" + txtenddate.DateTime.ToString("yyyy-MM-dd") +txthytcode.Text+ " 机型 报表";
                chartTitle.TextColor = System.Drawing.Color.Black; //颜色设置
                chartTitle.Font = new Font("微软雅黑", 10);
                chartTitle.Alignment = StringAlignment.Center;
                chartControlBar.Titles.Clear();
                chartControlBar.Titles.Add(chartTitle);


                XYDiagram diagram = chartControlBar.Diagram as XYDiagram;
                chartControlBar.SeriesTemplate.ArgumentDataMember = "类型";
                diagram.AxisX.Title.Visibility = DefaultBoolean.True;
                diagram.AxisX.Title.Text = "物料编码";


                chartControlBar.Series["批数"].ArgumentDataMember = "类型";
                chartControlBar.Series["批数"].ArgumentScaleType = ScaleType.Qualitative;
                chartControlBar.Series["批数"].ValueDataMembers[0] = "批数";
                chartControlBar.Series["批数"].LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
                chartControlBar.Series["批数"].LegendText = "批数";

           
            }
            else if (selecttype.SelectedIndex == 2)
            {
                if (customercheck.Text == "")
                    return;

                gridControl.DataSource = null;
                chartControlBar.DataSource = null;
         
                gridControl.DataSource = querydetail();            
                chartControlBar.DataSource = CustomerReport();

                ChartTitle chartTitle = new ChartTitle();
                chartTitle.Text = txtstartdate.DateTime.ToString("yyyy-MM-dd") + "至" + txtenddate.DateTime.ToString("yyyy-MM-dd") +customercheck.Text+ " 客户 报表";
                chartTitle.TextColor = System.Drawing.Color.Black; //颜色设置
                chartTitle.Font = new Font("微软雅黑", 10);
                chartTitle.Alignment = StringAlignment.Center;
                chartControlBar.Titles.Clear();
                chartControlBar.Titles.Add(chartTitle);

                XYDiagram diagram = chartControlBar.Diagram as XYDiagram;
                chartControlBar.SeriesTemplate.ArgumentDataMember = "类型";
                diagram.AxisX.Title.Visibility = DefaultBoolean.True;
                diagram.AxisX.Title.Text = "客户";

                chartControlBar.Series["批数"].ArgumentDataMember = "类型";
                chartControlBar.Series["批数"].ArgumentScaleType = ScaleType.Qualitative;
                chartControlBar.Series["批数"].ValueDataMembers[0] = "批数";
                chartControlBar.Series["批数"].LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
                chartControlBar.Series["批数"].LegendText = "批数";

            }
            else if (selecttype.SelectedIndex == 3)
            {

                if (checkbadclass.Text == "")
                    return;

                gridControl.DataSource = null;
                chartControlBar.DataSource = null;


                gridControl.DataSource = querydetail();
             
                chartControlBar.DataSource = BadClassReport();

                ChartTitle chartTitle = new ChartTitle();
                chartTitle.Text = txtstartdate.DateTime.ToString("yyyy-MM-dd") + "至" + txtenddate.DateTime.ToString("yyyy-MM-dd") + checkbadclass.Text+ " 不良类别 报表";
                chartTitle.TextColor = System.Drawing.Color.Black; //颜色设置
                chartTitle.Font = new Font("微软雅黑", 10);
                chartTitle.Alignment = StringAlignment.Center;
                chartControlBar.Titles.Clear();
                chartControlBar.Titles.Add(chartTitle);

                XYDiagram diagram = chartControlBar.Diagram as XYDiagram;
                chartControlBar.SeriesTemplate.ArgumentDataMember = "类型";
                diagram.AxisX.Title.Visibility = DefaultBoolean.True;
                diagram.AxisX.Title.Text = "不良类别";

                chartControlBar.Series["批数"].ArgumentDataMember = "类型";
                chartControlBar.Series["批数"].ArgumentScaleType = ScaleType.Qualitative;
                chartControlBar.Series["批数"].ValueDataMembers[0] = "批数";
                chartControlBar.Series["批数"].LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
                chartControlBar.Series["批数"].LegendText = "批数";

            }

            repositoryItemMemoEdit1.LinesCount = 4;

        }


        DataTable WorknoReport()
        {
            DataTable dt = null;      
            string workno = "";
            string[] worknoArray = txtworkno.Text.Trim().Split(new char[4] { ';', '；', ',', '，' });
            for (int i = 0; i < worknoArray.Length; i++)
            {
                workno += ",'" + worknoArray[i].ToString() + "'";
            }
            workno = workno.TrimStart(',');
            string sql = "  select workno 类型,COUNT (workno) 批数  from IPQCExceptionList where checkdate >='" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and checkdate <= '" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59' and workno in (" + workno + ")  ";
            if (batchexception.Checked == true)
            {
                sql += "   and  (( worknoqty >200 and convert(numeric(4,2),(NGQty+0.0)/(worknoqty)) >0.1 ) or ( worknoqty <= 200  and NGQty>20 ))  ";
            }
            sql += "   group by workno ";
            dt = DbAccess.SelectBySql(sql).Tables[0];
            return dt;
        }
        DataTable HytcodeReport()
        {
            DataTable dt = null;

            string hytcode = "";
            string[] hytcodeArray = txthytcode.Text.Trim().Split(new char[4] { ';', '；', ',', '，' });
            for (int i = 0; i < hytcodeArray.Length; i++)
            {
                hytcode += ",'" + hytcodeArray[i].ToString() + "'";
            }
            hytcode = hytcode.TrimStart(',');

            string sql = "  select hytcode 类型,COUNT (hytcode) 批数  from IPQCExceptionList where checkdate >='" + txtstartdate.DateTime.ToString("yyyy-MM-dd")+ " 00:00:00' and checkdate <= '"+txtenddate.DateTime.ToString("yyyy-MM-dd")+ " 23:59:59'  and  hytcode in (" + hytcode + ")   ";

            if (batchexception.Checked == true)
            {
                sql += "   and  (( worknoqty >200 and convert(numeric(4,2),(NGQty+0.0)/(worknoqty)) >0.1 ) or ( worknoqty <= 200  and NGQty>20 ))  ";
            }
            sql += "  group by hytcode ";
            dt = DbAccess.SelectBySql(sql).Tables[0];
            return dt;
        }

        DataTable CustomerReport()
        {          
            string customer = "";
            string[] customerArray = customercheck.Text.Split(',');
            for (int i = 0; i < customerArray.Length; i++)
            {
                customer += ",'" + customerArray[i].ToString().Trim()+"'";
            }
            customer = customer.TrimStart(',');

            DataTable dt = null;
            string sql = " select customer 类型,COUNT (customer) 批数  from ";
            sql += " IPQCExceptionList where checkdate >= '" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00'  and checkdate <= '" + txtenddate.DateTime.ToString("yyyy-MM-dd")+ " 23:59:59' and customer in (" + customer + ")  ";
            if (batchexception.Checked == true)
            {
                sql += "   and  (( worknoqty >200 and convert(numeric(4,2),(NGQty+0.0)/(worknoqty)) >0.1 ) or ( worknoqty <= 200  and NGQty>20 ))  ";
            }
            sql += "  group by customer ";
            dt = DbAccess.SelectBySql(sql).Tables[0];
            return dt;
        }

        DataTable BadClassReport()
        {
            DataTable dt = null;
            string badclass = "";
            string[] badclassArray = checkbadclass.Text.Split(',');
            for (int i = 0; i < badclassArray.Length; i++)
            {
                badclass += ",'" + badclassArray[i].ToString().Trim()+ "'";
            }
            badclass = badclass.TrimStart(',');
            string sql = "  select badclass 类型,COUNT (badclass) 批数  from IPQCExceptionList where checkdate >='" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and checkdate <= '" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59'  and badclass in (" + badclass + ")    ";
            if (batchexception.Checked == true)
            {
                sql += "   and  (( worknoqty >200 and convert(numeric(4,2),(NGQty+0.0)/(worknoqty)) >0.1 ) or ( worknoqty <= 200  and NGQty>20 ))  ";
            }
            sql += "  group by badclass ";
            dt = DbAccess.SelectBySql(sql).Tables[0];
            return dt;
        }


        private void sBtnreset_Click(object sender, EventArgs e)
        {
            gridControl.DataSource = null;
            chartControlBar.DataSource = null;
        }

        public void ExportToExcel(string title, params IPrintable[] panels)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.FileName = title;
            saveFileDialog.Title = "导出Excel";
            saveFileDialog.Filter = "Excel文件(*.xlsx)|*.xlsx|Excel文件(*.xls)|*.xls";
            DialogResult dialogResult = saveFileDialog.ShowDialog();
            if (dialogResult == DialogResult.Cancel)
                return;
            string FileName = saveFileDialog.FileName;
            PrintingSystem ps = new PrintingSystem();
            CompositeLink link = new CompositeLink(ps);
            ps.Links.Add(link);
            foreach (IPrintable panel in panels)
            {
                link.Links.Add(CreatePrintableLink(panel));
            }
            link.Landscape = true;  //横向           
            //判断是否有标题，有则设置         
            try
            {
                int count = 1;
                //在重复名称后加（序号）
                while (File.Exists(FileName))
                {
                    if (FileName.Contains(")."))
                    {
                        int start = FileName.LastIndexOf("(");
                        int end = FileName.LastIndexOf(").") - FileName.LastIndexOf("(") + 2;
                        FileName = FileName.Replace(FileName.Substring(start, end), string.Format("({0}).", count));
                    }
                    else
                    {
                        FileName = FileName.Replace(".", string.Format("({0}).", count));
                    }
                    count++;
                }
                if (FileName.LastIndexOf(".xlsx") >= FileName.Length - 5)
                {
                    XlsxExportOptions options = new XlsxExportOptions();
                    link.ExportToXlsx(FileName, options);
                }
                else
                {
                    XlsExportOptions options = new XlsExportOptions();
                    link.ExportToXls(FileName, options);
                }
                if (DevExpress.XtraEditors.XtraMessageBox.Show("保存成功，是否打开文件？", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                    System.Diagnostics.Process.Start(FileName);

                progressBarControl1.Position = 0;
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message);
            }
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

        PrintableComponentLink CreatePrintableLink(IPrintable printable)
        {
            ChartControl chart = printable as ChartControl;
            if (chart != null)
                chart.OptionsPrint.SizeMode = DevExpress.XtraCharts.Printing.PrintSizeMode.Stretch;
            PrintableComponentLink printableLink = new PrintableComponentLink()
            {
                Component = printable
            };
            return printableLink;
        }


        private void sBtnreport_Click(object sender, EventArgs e)
        {
            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
                return;
            if (gridView.RowCount < 1)
                return;
            string filename = chartControlBar.Titles[0].ToString();
            ExportToExcel(filename, gridControl, chartControlBar);
        }

        private void gridView_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
            {
                return;
            }

            if (gridView.GetDataRow(e.RowHandle)["是否批量异常"].ToString() == "是")
            {
                e.Appearance.BackColor = Color.Red;
            }

        }
        public bool Connect(string remoteHost)
        {
            bool Flag = true;
            Process proc = new Process();
            try
            {
                proc.StartInfo.FileName = "cmd.exe";
                proc.StartInfo.UseShellExecute = false;
                proc.StartInfo.RedirectStandardInput = true;
                proc.StartInfo.RedirectStandardOutput = true;
                proc.StartInfo.RedirectStandardError = true;
                proc.StartInfo.CreateNoWindow = true;
                proc.Start();
                string dosdel = @"net use \\" + remoteHost + " /del";
                proc.StandardInput.WriteLine(dosdel);
                //string dosLine = @"net use \\" + remoteHost + " hytera;2012" + " /user:" + "Upload";
                string dosLine = @"net use \\" + remoteHost + " hytera;2012" + " /user:" + this.serverFilePath.Split('\\')[0] + "\\Upload";
                proc.StandardInput.WriteLine(dosLine);
                proc.StandardInput.WriteLine("exit");
                while (proc.HasExited == false)
                {
                    proc.WaitForExit(1000);
                }
                string errormsg = proc.StandardError.ReadToEnd();
                if (errormsg != "")
                {
                    Flag = false;
                }
                proc.StandardError.Close();
            }
            catch (Exception ex)
            {
                Flag = false;
            }
            finally
            {
                try
                {
                    proc.Close();
                    proc.Dispose();
                }
                catch
                {
                }
            }
            return Flag;
        }

        private void OpenFile(string filepath, string pdffile)
        {
            string filename = "";
            //filename = pdffile + ".pdf";
            filename = pdffile;
            //定义一个ProcessStartInfo实例
            System.Diagnostics.ProcessStartInfo info = new System.Diagnostics.ProcessStartInfo();
            //设置启动进程的初始目录
            //string[] fileName = pdffile;

            //filename = fileName[fileName.Length - 1];
            info.FileName = @filename;
            info.WorkingDirectory = filepath;
            //设置启动进程的参数
            info.Arguments = "";
            //启动由包含进程启动信息的进程资源
            try
            {
                System.Diagnostics.Process.Start(info);
                //System.Diagnostics.Process.Start(Application.StartupPath +"\\"+info);

            }

            catch (System.ComponentModel.Win32Exception we)
            {

                MessageBox.Show(this, we.Message);
                return;
            }

        }

        void export(string item)
        {
            string sql = "";
            int sheetCount = 1;
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = true;
            object missing = System.Reflection.Missing.Value;
            string templetFile = Environment.CurrentDirectory + @"\ReportFolder\制程异常处理单表格.xlsx";
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

            sql = @"  select  ifrepeat 是否重复发生,standid 站别,item 表单编号,checkdate 日期,productlineid 生产线别,customer 客户,
		        workno 工单号,hytcode Hytera编码,model 客户机型,qty 投入数,NGQty 不良数,CASE WHEN NGQty < worknoqty  then  convert( varchar,convert(numeric(3,1),(NGQty+0.0)/(isnull(worknoqty,qty))*100 ))+'%' ELSE '100%' end 不良率,
				badclass 不良类别,baddescribe 问题描述,problemtype 问题分类,temporaryhandle 临时处理方法,CauseAnalysis 原因分析,improvemeasures 改善计划,
				dutyDepartment 责任部门,chargeMan 责任人,ifClose 是否关闭,ifOvertime 是否超期,overdutyDepartment 延时责任部门,badImage 不良图片
		        from IPQCExceptionList  where item = '" + item + "'";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            if (dt == null || dt.Rows.Count < 1)
                return;

            string ifrepeat = dt.Rows[0]["是否重复发生"].ToString();
            string problemtype = dt.Rows[0]["问题分类"].ToString();
            if (ifrepeat == "是")
            {
                sheet.Cells.get_Range("D22").Value = "√是";
            }
            else
            {
                sheet.Cells.get_Range("E22").Value = "√否";
            }
            sheet.Cells.get_Range("H2").Value = dt.Rows[0]["表单编号"].ToString();
            sheet.Cells.get_Range("C3").Value = dt.Rows[0]["客户机型"].ToString();
            sheet.Cells.get_Range("F3").Value = dt.Rows[0]["Hytera编码"].ToString();
            sheet.Cells.get_Range("J3").Value = dt.Rows[0]["生产线别"].ToString();
            sheet.Cells.get_Range("C4").Value = dt.Rows[0]["投入数"].ToString();
            sheet.Cells.get_Range("F4").Value = dt.Rows[0]["不良数"].ToString();
            sheet.Cells.get_Range("J4").Value = dt.Rows[0]["不良率"].ToString();
            sheet.Cells.get_Range("A5").Value = "异常描述:" + "\r\n" + dt.Rows[0]["问题描述"].ToString();
            sheet.Cells.get_Range("D10").Value = dt.Rows[0]["临时处理方法"].ToString();
            sheet.Cells.get_Range("B15").Value = dt.Rows[0]["原因分析"].ToString();
            sheet.Cells.get_Range("B21").Value = "问题分类：" + dt.Rows[0]["问题分类"].ToString();
            sheet.Cells.get_Range("B24").Value = dt.Rows[0]["改善计划"].ToString();

        }


        private void gridView_RowCellClick(object sender, RowCellClickEventArgs e)
        {
            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
                return;
            if (gridView.RowCount < 1)
                return;
            if (gridView.FocusedRowHandle < 0)
                return;
            string workno = gridView.GetFocusedRowCellValue("工单号").ToString();
            string item = gridView.GetFocusedRowCellValue("表单编号").ToString();
            string hytcode = gridView.GetFocusedRowCellValue("Hytera编码").ToString();

            if (e.Column.FieldName == "表单编号")
            {

                string floerPath = "\\\\" + serverFilePath + "\\制程异常处理报告" + "\\" + workno + "\\" + item;

                if (Connect(serverFilePath))
                {
                    if (Directory.Exists(floerPath) == false)
                    {
                        MessageBox.Show("还没有上传相应的处理报告!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        return;
                    }

                    string[] str = Directory.GetFiles(floerPath);
                    if (str.Length > 0)
                    {
                        for (int i = 0; i < str.Length; i++)
                        {
                            if (str[i].ToString().Contains(".db"))
                                continue;
                            string filename = Path.GetFileName(str[i].ToString());

                            FileStream fs = new FileStream(str[i].ToString(), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                            BinaryReader reader = new BinaryReader(fs);
                            string fileclass = "";
                            try
                            {
                                for (int j = 0; j < 2; j++)
                                {
                                    fileclass += reader.ReadByte().ToString();
                                }
                            }
                            catch (Exception)
                            {

                                throw;
                            }
                            {
                                OpenFile(floerPath, filename);
                                fs.Close();
                                fs.Dispose();
                                reader.Close();
                                return;
                            }
                        }

                    }
                }
            }
            if (e.Column.FieldName == "工单号")
            {
                try
                {
                    export(item);

                }
                catch (Exception ex)
                {

                    MessageBox.Show("还没生成完整报表", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

            }
            if (e.Column.FieldName == "Hytera编码")
            {
                try
                {
                    string badfilePath = "\\\\" + serverFilePath + "\\制程不良图片" + "\\" + workno + "\\" + item;
                    Cursor.Current = Cursors.WaitCursor;
                    if (Connect(serverFilePath))
                    {
                        if (Directory.Exists(badfilePath))
                        {
                            string[] pt = Directory.GetFiles(badfilePath);
                            if (pt.Length == 0)
                                return;
                            if (pt.Length > 0)
                            {
                                for (int i = 0; i < pt.Length; i++)
                                {
                                    if (pt[i].ToString().Contains(".db"))
                                        continue;

                                    string filename = Path.GetFileName(pt[0].ToString());
                                    OpenFile(badfilePath, filename);

                                    //FileStream ft = new FileStream(pt[i].ToString(), FileMode.Open);
                                    //Image badpic = Image.FromStream(ft);
                                    //ft.Close();
                                    //ft.Dispose();
                                    //picbadImage.Image = badpic;
                                    //picbadImage.Properties.SizeMode = PictureSizeMode.Stretch;
                                }
                            }
                        }
                    }
                    Cursor.Current = Cursors.Default;

                }
                catch (Exception ex)
                {

                    MessageBox.Show("还没上传不良图片", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

            }
        }
        private string ShowSaveFileDialog(string title, string filter)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            string name = "制程异常信息记录";
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
 
        private void sBtnreportdetail_Click(object sender, EventArgs e)
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
    }
}