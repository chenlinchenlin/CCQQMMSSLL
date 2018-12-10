using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Windows.Forms;
using DevExpress.XtraBars;
using System.Data.SqlClient;
using System.Data.OleDb;
using DevExpress.XtraEditors;
using DevExpress.XtraGrid.Menu;
using DevExpress.Utils.Menu;
using DX_QMS.Common;
using System.IO;
using DX_QMS.KPI.KPIdata;
using System.Diagnostics;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraEditors;
using DevExpress.XtraGrid;
using DevExpress.Data;

namespace DX_QMS.KPI
{
    public partial class KPIitem : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        DXPopupMenu formatRulesMenu = new DXPopupMenu();
        private GridGroupSummaryItemCollection gsiMultiSummary;
        DataSet ds = new DataSet();
        public KPIitem()
        {
            InitializeComponent();

            //this.sBtnAdd.Enabled = false;
            //this.sBtndelete.Enabled = false;

            //setRule();

            //InitMDBData();

        }

        private void setRule()
        {
            string post = "";
            if (Login.manager != null && Login.manager != "")
            {
                post = Login.manager;
            }
            else
            {
                post = Login.post;
            }
            Dictionary<string, bool> dic = GroupPermission.QMS_SelectRulesForForm(post, "KPI");
            this.sBtnAdd.Enabled = bool.Parse(dic["hasInsert"].ToString());
            this.sBtndelete.Enabled = bool.Parse(dic["hasDelete"].ToString());
        }

        private void bindBusinessType()
        {
            string sql = "  select distinct businessType from QMS_KPIindicators  ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            txtbusinessType.Properties.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
                txtbusinessType.Properties.Items.Add(row["businessType"]);
            }
            txtbusinessType.SelectedIndex = 0;
        }



        protected void InitMDBData()
        {   
            SqlConnection conText = new SqlConnection(DbAccess.connSql);

            try
            {
                conText.Open();
                 
                string adsql = @"  select businessType 业务分类,max(items) 顺序,indicatorsName 指标名称,max(checkDepartment) 考核部门,max(undertakeLevel) 承担级别,
		          max(qualityMan) 质量部责任人,max(ifComplete) 是否达成,max(dataSource) 数据来源,max(mathFormula) 计算公式,max(Remark) 备注,max(KpiYeartime) KPI年份 from QMS_KPIDataSource  group by businessType ,indicatorsName ";
                SqlDataAdapter adda = new SqlDataAdapter(adsql, conText);
                adda.Fill(ds, "adtable");

                string desql = @" select indicatorsName 指标名称,kpiData  KPI数据,lastMonthkpiData 上月KPI数据,nextMonthkipData 下月KPI数据,thisYearkpiData  今年KPI数据,lastYearkpiData 去年KPI数据,nextYearkipData 明年KPI数据,kpiMonthtime  KPI时间 
		        from QMS_KPIDataSource order by updateTime desc  ";
                SqlDataAdapter deda = new SqlDataAdapter(desql, conText);
                deda.Fill(ds, "detable");

            }
            catch
            {
                conText.Close(); 
            }

            ds.Relations.Add("adtable",
                ds.Tables["adtable"].Columns["指标名称"],
                ds.Tables["detable"].Columns["指标名称"], false);


            gridControl.DataSource = ds.Tables["adtable"];

            gridControl.MainView.PopulateColumns();

        }


        private void KPIitem_Load(object sender, EventArgs e)
        {
            bindBusinessType();
            gridView.PrintExportProgress += new ProgressChangedEventHandler(gridView_PrintExportProgress);
            InitSummaries();
        }

        private void businessType_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txtbusinessType.Text.Trim() == "")
                return;
            string sql = " select indicatorsName from QMS_KPIindicators where businessType = '"+ txtbusinessType.Text + "' ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            txtindicatorsName.Properties.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
                txtindicatorsName.Properties.Items.Add(row["indicatorsName"]);
            }
            txtindicatorsName.SelectedIndex = 0;


        }

    

        public string KPIitemList( string  opertype ,string KpiYeartime ,  string kpiMonthtime ,string businessType , string items ,  string indicatorsName , string checkDepartment ,
             string undertakeLevel , string qualityMan , string dataSource , string kpiData , string lastYearMonthkpiData ,  string lastYearkpiData , string thisYearkpiData ,
             string ifComplete ,string mathFormula ,string improvementMan ,string emailID ,string reasonAnalysis ,string improvementMeasures ,string problemState ,string Remark , string updateMan)
        {

            SqlParameter[] para = new SqlParameter[24];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@KpiYeartime", KpiYeartime);
            para[2] = new SqlParameter("@kpiMonthtime", kpiMonthtime);
            para[3] = new SqlParameter("@businessType", businessType);
            para[4] = new SqlParameter("@items", items);
            para[5] = new SqlParameter("@indicatorsName", indicatorsName);
            para[6] = new SqlParameter("@checkDepartment", checkDepartment);
            para[7] = new SqlParameter("@undertakeLevel", undertakeLevel);
            para[8] = new SqlParameter("@qualityMan", qualityMan);
            para[9] = new SqlParameter("@dataSource", dataSource);
            para[10] = new SqlParameter("@kpiData", kpiData);
            para[11] = new SqlParameter("@lastYearMonthkpiData", lastYearMonthkpiData);
            para[12] = new SqlParameter("@lastYearkpiData", lastYearkpiData);
            para[13] = new SqlParameter("@thisYearkpiData", thisYearkpiData);
            para[14] = new SqlParameter("@ifComplete", ifComplete);
            para[15] = new SqlParameter("@mathFormula", mathFormula);
            para[16] = new SqlParameter("@improvementMan", improvementMan);
            para[17] = new SqlParameter("@emailID", emailID);
            para[18] = new SqlParameter("@reasonAnalysis", reasonAnalysis);
            para[19] = new SqlParameter("@improvementMeasures", improvementMeasures);
            para[20] = new SqlParameter("@problemState", problemState);
            para[21] = new SqlParameter("@Remark", Remark);
            para[22] = new SqlParameter("@updateMan", updateMan);
            para[23] = new SqlParameter("@msg", SqlDbType.VarChar, 50);
            para[23].Direction = ParameterDirection.Output;
            DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "QMS_OperKPIitem", para);
            return para[23].Value.ToString();
        }




        private Microsoft.Office.Interop.Excel.ApplicationClass appClsExcel = null;
        Microsoft.Office.Interop.Excel.Workbooks wbs = null;
        Microsoft.Office.Interop.Excel.Workbook wb = null;
        Microsoft.Office.Interop.Excel.Worksheet ws = null;
        Microsoft.Office.Interop.Excel.Range range = null;
        private void OpenExcel(string strFileName)
        {
            object objMissing = System.Reflection.Missing.Value;
            this.appClsExcel = new Microsoft.Office.Interop.Excel.ApplicationClass();
            appClsExcel.UserControl = true;
            appClsExcel.DisplayAlerts = false;

            Microsoft.Office.Interop.Excel.Workbook wBookExcel = appClsExcel.Workbooks.Open(strFileName, objMissing, true, objMissing, objMissing, objMissing
                          , objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing);

            Microsoft.Office.Interop.Excel.Worksheet wSheetExcel = (Microsoft.Office.Interop.Excel.Worksheet)wBookExcel.ActiveSheet;
            Microsoft.Office.Interop.Excel.Range rCell = wSheetExcel.UsedRange;

            object[,] objList = (object[,])rCell.Value2;

            int excelcounts = objList.GetLength(0);

            wbs = appClsExcel.Workbooks;
            wb = wbs[1];
            ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets["质量KPI监控数据"];
            int rowCount = ws.UsedRange.Rows.Count;
            int colCount = ws.UsedRange.Columns.Count;
            if (rowCount <= 0)
            {
                MessageBox.Show("文件中没有数据记录", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (colCount < 21)
            {
                MessageBox.Show("字段个数不对", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            DataTable KPIDataTable = new KPIDataSet().Tables["QMS_KPIDataSource"];

            try
            {
                for (int iRow = 1; iRow < rowCount; iRow++)
                {
                    if (objList[iRow + 1, 1].ToString() != "" && objList[iRow + 1, 2].ToString() != "" && objList[iRow + 1, 3].ToString() != "" && objList[iRow + 1, 4].ToString() != "")
                    {
                        DataRow row;
                        row = KPIDataTable.NewRow();

                        row["KpiYeartime"] = objList[iRow + 1, 1].ToString();
                        row["kpiMonthtime"] = objList[iRow + 1, 2].ToString();
                        row["businessType"] = objList[iRow + 1, 3].ToString();                    
                        row["indicatorsName"] = objList[iRow + 1, 4].ToString();
                        row["checkDepartment"] = objList[iRow + 1, 5] == null ? "" : objList[iRow + 1, 5].ToString();
                        row["undertakeLevel"] = objList[iRow + 1, 6] == null ? "" : objList[iRow + 1, 6].ToString();
                        row["qualityMan"] = objList[iRow + 1, 7] == null ? "" : objList[iRow + 1, 7].ToString();
                        row["dataSource"] = objList[iRow + 1, 8] == null ? "" : objList[iRow + 1, 8].ToString();
                        row["lastYearkpiData"] = objList[iRow + 1, 9] == null ? "" : objList[iRow + 1, 9].ToString();
                        row["lastYearMonthkpiData"] = objList[iRow + 1, 10] == null ? "" : objList[iRow + 1, 10].ToString();
                        row["thisYearkpiData"] = objList[iRow + 1, 11] == null ? "" : objList[iRow + 1, 11].ToString();
                        row["kpiData"] = objList[iRow + 1, 12] == null ? "" : objList[iRow + 1, 12].ToString(); ;
                        row["mathFormula"] = objList[iRow + 1, 13] == null ? "" : objList[iRow + 1, 13].ToString();
                        row["ifComplete"] = objList[iRow + 1, 14] == null ? "" : objList[iRow + 1, 14].ToString();
                        row["improvementMan"] = objList[iRow + 1, 15] == null ? "" : objList[iRow + 1, 15].ToString();
                        row["emailID"] = objList[iRow + 1, 16] == null ? "" : objList[iRow + 1, 16].ToString();
                        row["reasonAnalysis"] = objList[iRow + 1, 17] == null ? "" : objList[iRow + 1, 17].ToString();
                        row["improvementMeasures"] = objList[iRow + 1, 18] == null ? "" : objList[iRow + 1, 18].ToString();
                        row["problemState"] = objList[iRow + 1, 19] == null ? "" : objList[iRow + 1, 19].ToString();
                        row["Remark"] = objList[iRow + 1, 21] == null ? "" : objList[iRow + 1, 21].ToString();
                        row["updateMan"] = objList[iRow + 1, 20] == null ? "" : objList[iRow + 1, 20].ToString();


                        KPIDataTable.Rows.Add(row);
                    }
                    else
                    {
                        break;
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message , "停止", MessageBoxButtons.OK, MessageBoxIcon.Stop);

            }

            gridColumn3.GroupIndex = -1;
            gridColumn4.GroupIndex = -1;
            gridColumn1.GroupIndex = -1;
            this.gridControl.DataSource = KPIDataTable;

            appClsExcel.Quit();
            appClsExcel = null;
            Process[] procs = Process.GetProcessesByName("Excel");
            foreach (Process pro in procs)
            {
                pro.Kill();
            }
            GC.Collect();
        }

        private void sBtnimport_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.OpenFileDialog fd = new OpenFileDialog();
            if (fd.ShowDialog() == DialogResult.OK)
            {
                string fileExtenSion;
                fileExtenSion = Path.GetExtension(fd.FileName);
                if (fileExtenSion.ToLower() != ".xls" && fileExtenSion.ToLower() != ".xlsx")
                {
                    MessageBox.Show("文件格式不正确", "停止", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }

                try
                {
                    OpenExcel(fd.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("导入失败", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

            }
        }

        private void sBtnsubmit_Click(object sender, EventArgs e)
        {
            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
                return;

            int countsuccess = 0;

            for (int i = 0; i < dt.Rows.Count; i++ )
            {
                DataRow dr = dt.Rows[i];

                string flag = KPIitemList("新增", dr["KpiYeartime"].ToString(), dr["kpiMonthtime"].ToString(), dr["businessType"].ToString(), "", dr["indicatorsName"].ToString(), dr["checkDepartment"].ToString(),
    dr["undertakeLevel"].ToString(), dr["qualityMan"].ToString(), dr["dataSource"].ToString(), dr["kpiData"].ToString(), dr["lastYearMonthkpiData"].ToString(), dr["lastYearkpiData"].ToString(), dr["thisYearkpiData"].ToString(),
    dr["ifComplete"].ToString(), dr["mathFormula"].ToString(), dr["improvementMan"].ToString(), dr["emailID"].ToString(), dr["reasonAnalysis"].ToString(), dr["improvementMeasures"].ToString(), dr["problemState"].ToString(), dr["Remark"].ToString(), Login.username); 
                if (flag == "新增成功")
                {
                    countsuccess++;
                }
            }

            MessageBox.Show("提交成功" + countsuccess + "条,失败" + (dt.Rows.Count - countsuccess) + "条", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);

        }

        private void sBtnAdd_Click(object sender, EventArgs e)
        {
            if (txtbusinessType.Text.Trim() == "" || txtindicatorsName.Text.Trim() == "" || txtcheckDepartment.Text.Trim() == "")
            {
                MessageBox.Show("请输入业务类型、指标名称和考核部门", "提醒",MessageBoxButtons.OK,MessageBoxIcon.Information);
                return;
            }

            if (txtkpiMonthtime.Text.Trim() == "")
            {
                MessageBox.Show("请选择时间", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string kpitime = txtkpiMonthtime.DateTime.ToString("yyyy-MM-dd");

            string kpiMonthtime = txtkpiMonthtime.DateTime.ToString("yyyyMM");
            string KpiYeartime = txtkpiMonthtime.DateTime.ToString("yyyy");

            string falg = KPIitemList("新增", KpiYeartime, kpiMonthtime, txtbusinessType.Text,"", txtindicatorsName.Text, txtcheckDepartment.Text,
             txtundertakeLevel.Text, txtqualityMan.Text, txtdataSource.Text, txtkpiData.Text, txtlastYearMonthkpiData.Text, txtlastYearkpiData.Text, txtthisYearkpiData.Text,
             txtifComplete.Text, txtmathFormula.Text, txtimprovementMan.Text, txtemailID.Text, txtreasonAnalysis.Text, txtimprovementMeasures.Text, txtproblemState.Text, txtRemark.Text, Login.username);

            if (falg.Contains("新增成功"))
            {
                MessageBox.Show(falg, "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("新增失败", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }

        private void InitSummaries()
        {
            gsiMultiSummary = new GridGroupSummaryItemCollection(gridView);
            gsiMultiSummary.Add(SummaryItemType.Count, "businessType");
            gsiMultiSummary.Add(SummaryItemType.Count, "indicatorsName");
            gsiMultiSummary.Add(SummaryItemType.Count, "KpiYeartime");
            ///// gsiMultiSummary.Add(SummaryItemType.Average, "工单号", null,"");
        }

        private void sBtnselect_Click(object sender, EventArgs e)
        {

            //if (txtbusinessType.Text.Trim() == "" || txtindicatorsName.Text.Trim() == "" || txtcheckDepartment.Text.Trim() == "")
            //{
            //    MessageBox.Show("请输入业务类型、指标名称和考核部门", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //    return;
            //}
            string where = " where 1=1 ";
            string businessType = txtbusinessType.Text;
            string indicatorsName = txtindicatorsName.Text;
            string checkDepartment = txtcheckDepartment.Text;

            if (!string.IsNullOrEmpty(businessType))
            {
                where += " and businessType = '" + businessType + "' ";
            }
            if (!string.IsNullOrEmpty(indicatorsName))
            {
                where += " and indicatorsName = '" + indicatorsName + "' ";
            }
            if (!string.IsNullOrEmpty(checkDepartment))
            {
                where += " and checkDepartment = '" + checkDepartment + "' ";
            }

            string sql = @"  select * from QMS_KPIDataSource  ";
            sql += where + " order by updateTime desc ";

            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];

            if (dt != null && dt.Rows.Count > 0)
            {
                gridControl.DataSource = dt;
                gridColumn3.GroupIndex = 0;
                gridColumn4.GroupIndex = 1;
                gridColumn1.GroupIndex = 2;                            
                gridView.ExpandAllGroups();
                gridView.GroupSummary.Assign(gsiMultiSummary);



            }
            else
            {
                MessageBox.Show("没有符合条件的记录");
                gridControl.DataSource = null;
            }


        }

        private void sBtndelete_Click(object sender, EventArgs e)
        {

            int i = gridView.FocusedRowHandle;
            if (i < 0)
            {
                MessageBox.Show("请选中要删除的信息", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            if (gridView.GetSelectedRows().Length <= 0)
                return;
            int count = 0;
            for (int k = gridView.GetSelectedRows().Length; k > 0; k--)
            {
                DataRow dr = gridView.GetDataRow(gridView.GetSelectedRows()[k - 1]);
                string KpiYeartime = dr["KpiYeartime"].ToString();
                string kpiMonthtime = dr["kpiMonthtime"].ToString();
                string businessType = dr["businessType"].ToString();
                string indicatorsName = dr["indicatorsName"].ToString();
                string checkDepartment = dr["checkDepartment"].ToString();

                string falg = KPIitemList("删除", KpiYeartime, kpiMonthtime, businessType,"", indicatorsName, checkDepartment,
                  "", "", "", "", "", "", "","", "", "", "", "", "", "", "", "");

                if (falg.Contains("删除成功"))
                {
                    count = count + 1;
                }

            }

            MessageBox.Show("成功删除" + count + "条", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            sBtnselect_Click(null, null);

        }


        private string ShowSaveFileDialog(string title, string filter)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            string name = "质量KPI监控数据";
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
        private void sBtnexport_Click(object sender, EventArgs e)
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

        private void sBtnreset_Click(object sender, EventArgs e)
        {
            txtbusinessType.Text = "";
            txtindicatorsName.Text = "";
            txtcheckDepartment.Text = "";
            txtundertakeLevel.Text = "";
            txtqualityMan.Text = "";
            txtdataSource.Text = "";
            txtkpiData.Text = "";
            txtlastYearMonthkpiData.Text = "";
            txtlastYearkpiData.Text = "";
            txtthisYearkpiData.Text = "";
            txtifComplete.Text = "";
            txtmathFormula.Text = "";
            txtimprovementMan.Text = "";
            txtemailID.Text = "";
            txtreasonAnalysis.Text = "";
            txtimprovementMeasures.Text = "";
            txtproblemState.Text = "";
            txtRemark.Text = "";
            gridControl.DataSource = null;
        }

        private void gridView_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            gridView.IndicatorWidth = 50;

            if (e.Info.IsRowIndicator && e.RowHandle >= 0)
            {
                e.Info.DisplayText = (e.RowHandle + 1).ToString();
            }
        }

        private void gridView_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {
            try
            {

                txtbusinessType.Text = gridView.GetFocusedRowCellValue("businessType").ToString();
                txtindicatorsName.Text = gridView.GetFocusedRowCellValue("indicatorsName").ToString();
                txtcheckDepartment.Text = gridView.GetFocusedRowCellValue("checkDepartment").ToString();
                txtundertakeLevel.Text = gridView.GetFocusedRowCellValue("undertakeLevel").ToString();
                txtqualityMan.Text = gridView.GetFocusedRowCellValue("qualityMan").ToString();
                txtdataSource.Text = gridView.GetFocusedRowCellValue("dataSource").ToString();
                txtkpiData.Text = gridView.GetFocusedRowCellValue("kpiData").ToString();
                txtlastYearMonthkpiData.Text = gridView.GetFocusedRowCellValue("lastYearMonthkpiData").ToString();
                txtlastYearkpiData.Text = gridView.GetFocusedRowCellValue("lastYearkpiData").ToString();
                txtthisYearkpiData.Text = gridView.GetFocusedRowCellValue("thisYearkpiData").ToString();
                txtifComplete.Text = gridView.GetFocusedRowCellValue("ifComplete").ToString();
                txtmathFormula.Text = gridView.GetFocusedRowCellValue("mathFormula").ToString();
                txtimprovementMan.Text = gridView.GetFocusedRowCellValue("improvementMan").ToString();
                txtemailID.Text = gridView.GetFocusedRowCellValue("emailID").ToString();
                txtreasonAnalysis.Text = gridView.GetFocusedRowCellValue("reasonAnalysis").ToString();
                txtimprovementMeasures.Text = gridView.GetFocusedRowCellValue("improvementMeasures").ToString();
                txtproblemState.Text = gridView.GetFocusedRowCellValue("problemState").ToString();
                txtRemark.Text = gridView.GetFocusedRowCellValue("Remark").ToString();

            }
            catch
            {

            }
        }

        private void gridView_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                GridFormatRuleMenuItems items = new GridFormatRuleMenuItems(gridView, e.Column, formatRulesMenu.Items);
                if (items.Count > 0)
                    MenuManagerHelper.ShowMenu(formatRulesMenu, gridControl.LookAndFeel, gridControl.MenuManager, gridControl, new Point(e.X, e.Y));
            }
        }
    }
}