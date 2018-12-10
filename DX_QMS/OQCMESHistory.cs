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
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraEditors;

namespace DX_QMS
{
    public partial class OQCMESHistory : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public string oldtestitem = "";
        private IQC ic = new IQC();
        public OQCMESHistory()
        {
            InitializeComponent();
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
        private void OQCMESHistory_Load(object sender, EventArgs e)
        {
            bindorg_id();
            gridView.PrintExportProgress += new ProgressChangedEventHandler(gridView_PrintExportProgress);
            bindDeviceType();

            if (Login.post == "OQC管理员" || Login.manager == "IPQC&OQC管理员" || Login.manager == "IT管理员")
            {
                Add.Enabled = true;
                Delete.Enabled = true;
                update.Enabled = true;
            }
            else
            {
                Add.Enabled = false;
                Delete.Enabled = false;
                update.Enabled = false;
            }

        }
        private void btnselect_Click(object sender, EventArgs e)
        {
            string where = " where 1=1 ";
            string org_id = txtorg_id.Text.Trim();
            string customer = txtcustomer.Text.Trim();
            string workno = txtworkno.Text.Trim();
            string PN = txtPN.Text.Trim();
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
            if (!string.IsNullOrEmpty(begintime))
            {
                where += " and checkdate >= '" + begintime + " 00:00:00 '";
            }
            if (!string.IsNullOrEmpty(endtime))
            {
                where += " and checkdate <='" + endtime + " 23:59:59 '";
            }
            string sql = @" select o.checkdate 检验日期, hytcode Hytera编码, cuscode 客户型号, customer 客户, org_id 组织, workno 工单号, lineid 线别, testitem 测试项目,code 检验内容,tools 测试工具,AQL 抽样水准,sendqty 送检数量, 
	          sampleqty 抽检数量,fqty 实抽数量,NGQty NG数量,testresult 检验结果, testremark 备注, testman 检验员, QC 产线QC,latyper 责任拉长, QE 责任QE, Masters 责任主管, states 产品状态,checkman 审核人,Auditman 批准人,item 序号,CartonNo 箱号
	          from OQC_TestList o left join (select count(*) fqty,items from OQC_SampleFactList group by items) t on o.item=t.items left join OQC_TypeDefine d on o.testitem=d.Definevalue  ";

            sql += where + " order by checkdate desc ";

            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];

            if (dt != null && dt.Rows.Count > 0)
            {
                gridControl.DataSource = dt;
            }
            else
            {
                MessageBox.Show("没有符合条件的记录");
                gridControl.DataSource = null;
            }
        }

        private void btnreset_Click(object sender, EventArgs e)
        {
            txtorg_id.Text = "";
            txtcustomer.Text = "";
            txtworkno.Text = "";
            txtPN.Text = "";           
            begindate.Text = "";
            enddate.Text = "";
            gridControl.DataSource = null;
        }

        private void gridView_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
        {
            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 0)
            {
                return;
            }
            if (gridView.GetDataRow(e.RowHandle)["检验结果"].ToString() == "NG")
            {
                e.Appearance.BackColor = Color.Red;
            }
        }
        private string ShowSaveFileDialog(string title, string filter)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            string name = "MES系统OQC测试记录信息";
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
            if (XtraMessageBox.Show("打开此文件?", "导 出...", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
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
           //progressBarControl1.Position = 0;
        }
        void gridView_PrintExportProgress(object sender, ProgressChangedEventArgs e)
        {
            SetPosition(e.ProgressPercentage);
        }
        void SetPosition(int pos)
        {
            //progressBarControl1.Position = pos;
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




        private void bindDeviceType()
        {
            string sql = "select checkType from OQC_CheckType";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            cbTestType.Properties.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
                cbTestType.Properties.Items.Add(row["checkType"]);
            }
            cbTestType.SelectedIndex = 0;
            Select_Click(null, null);
        }

        private void Reset_Click(object sender, EventArgs e)
        {
            dgvDefect.DataSource = null;
            if (cbTestType.Enabled)
                cbTestType.SelectedIndex = -1;
            txtTestItem.Text = "";
            txtid.Text = "";
            oldtestitem = "";
            txtconment.Text = "";
            this.txtTestItem.BackColor = Color.White;
            Add.Text = "新 增";
        }


        private DataTable GetType(string checktype, string stype)
        {
            string sql = "select Definetype,Definevalue,sort,code from OQC_TypeDefine where Definetype='" + checktype + "' and Definevalue like '%" + stype + "%'";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            return dt;
        }
        private void Select_Click(object sender, EventArgs e)
        {
            if (cbTestType.Text == "")
                return;
            DataTable dt = GetType(cbTestType.Text, txtTestItem.Text.Trim());
            if (dt.Rows.Count > 0)
                dgvDefect.DataSource = dt;
            else
                dgvDefect.DataSource = null;
        }

        private void Add_Click(object sender, EventArgs e)
        {
            if (txtTestItem.Text.Trim() == "")
                return;
            int m = 0;
            if (!int.TryParse(txtid.Text, out m))
            {
                MessageBox.Show("顺序请输入数字");
                return;
            }
            int i = ic.AddOQCCheckItem("新增", cbTestType.Text, txtTestItem.Text.Trim(), int.Parse(txtid.Text), oldtestitem, txtconment.Text);
            if (i > 0)
                dgvDefect.DataSource = GetType(cbTestType.Text, txtTestItem.Text.Trim());
            oldtestitem = "";
            txtTestItem.BackColor = Color.White;
            Add.Text = "新 增";
        }

        private void Delete_Click(object sender, EventArgs e)
        {
            DataTable de = dgvDefect.DataSource as DataTable;
            if (de == null || de.Rows.Count < 1)
                return;
            if (gridOQCView.FocusedRowHandle < 0)
                return;
            if (MessageBox.Show("确定删除选中的测试类别？", "删除提示！", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                int temp = ic.AddOQCCheckItem("删除", gridOQCView.GetFocusedRowCellValue("Definetype").ToString(), gridOQCView.GetFocusedRowCellValue("Definevalue").ToString(), 0, "", "");
                if (temp > 0)
                    dgvDefect.DataSource = GetType(cbTestType.Text, txtTestItem.Text.Trim());
            }
        }

        private void update_Click(object sender, EventArgs e)
        {
            if (txtTestItem.Text.Trim() == "") return;
            int m = 0;
            if (!int.TryParse(txtid.Text, out m))
            {
                MessageBox.Show("顺序请输入数字");
                return;
            }
            int i = ic.AddOQCCheckItem("更新", cbTestType.Text, txtTestItem.Text.Trim(), int.Parse(txtid.Text), oldtestitem, txtconment.Text);
            if (i > 0)
                dgvDefect.DataSource = GetType(cbTestType.Text, txtTestItem.Text.Trim());
            oldtestitem = "";
            txtTestItem.BackColor = Color.White;
            update.Enabled = false;
        }

        private void gridOQCView_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {
            if (gridOQCView.RowCount<1)
            {
                return;
            }
            cbTestType.Text = gridOQCView.GetFocusedRowCellValue("Definetype").ToString();
            txtTestItem.Text = gridOQCView.GetFocusedRowCellValue("Definevalue").ToString();
            txtid.Text = gridOQCView.GetFocusedRowCellValue("sort").ToString();
            txtconment.Text = gridOQCView.GetFocusedRowCellValue("code").ToString();
            oldtestitem = gridOQCView.GetFocusedRowCellValue("Definevalue").ToString();
            this.txtTestItem.BackColor = Color.Yellow;
            update.Enabled = true;
        }

        private void cbTestType_SelectedIndexChanged(object sender, EventArgs e)
        {
            Select_Click(sender, e);
        }
    }
}