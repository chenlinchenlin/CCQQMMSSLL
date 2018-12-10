using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using DevExpress.XtraBars;
using DX_QMS.Common;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraEditors;

namespace DX_QMS
{
    public partial class MaterialWaitForQC : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public MaterialWaitForQC()
        {
            InitializeComponent();
        }

        private void txtMaterialCode_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13 && txtMaterialCode.Text !="")
            {

            }
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            txtMaterialCode.Text = "";
            dateTimePickerBegin.Text = "";
            dateTimePickerEnd.Text = "";
            gridControl.DataSource = null;
        }

        private void btnQuery_Click(object sender, EventArgs e)
        {
            if (txtMaterialCode.Text == "" && dateTimePickerBegin.Text == "")
            {
                MessageBox.Show("物料编码和日期不能都为空");
                return;
            }

            if (txtMaterialCode.Text.Trim() != "")
            {
                string sql = @"select deliveryid 接收单号,d.materialcode 料号,max(materialname) 名称,sum(cast(qty as bigint)) 数量,vendorcode 供应商代码,vendorname 供应商名称,max(pushtimestamp) 生成条码时间,
                   cast(round(DATEDIFF(minute,max(pushtimestamp),getdate())/60.00,2) as numeric(18,2)) 周期,org_id 组织 from delivery d left join MaterialSpec m on d.materialcode=m.materialcode
                     where lotno like 'Z%' and not exists(select 1 from deliveryCheck c where productcode = '" + txtMaterialCode.Text + "' and d.lotno = c.lotno)and d.materialcode = '" + txtMaterialCode.Text + "'";
                string sqlwhere = " and 1=1";
                if (dateTimePickerBegin.Text  != "")//使用物料编码查询，可附带日期查询
                {
                    sqlwhere = " and pushtimestamp>='" + dateTimePickerBegin.DateTime.ToString("yyyy-MM-dd") + " 00:00:00" + "' and pushtimestamp<='" + dateTimePickerEnd.DateTime.ToString("yyyy-MM-dd") + " 23:59:59" + "' ";
                }
                sqlwhere += " group by deliveryid,d.materialcode,vendorcode,vendorname,org_id order by 生成条码时间";
                sql = sql + sqlwhere;
                DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
                gridControl.DataSource = dt;
            }
            else
            {
                if (dateTimePickerBegin.Text  != "")//只有日期查询
                {
                    string sql = @"select deliveryid 接收单号,d.materialcode 料号,max(materialname) 名称,sum(cast(qty as bigint)) 数量,vendorcode 供应商代码,vendorname 供应商名称,max(pushtimestamp) 生成条码时间,
                    cast(round(DATEDIFF(minute,max(pushtimestamp),getdate())/60.00,2) as numeric(18,2)) 周期,org_id 组织 from delivery d left join MaterialSpec m on d.materialcode=m.materialcode
                           where lotno like 'Z%' and not exists(select 1 from deliveryCheck c where  d.lotno = c.lotno)";
                    string sqlwhere = " and pushtimestamp>='" + dateTimePickerBegin.DateTime.ToString("yyyy-MM-dd") + " 00:00:00" + "' and pushtimestamp<='" + dateTimePickerEnd.DateTime.ToString("yyyy-MM-dd") + " 23:59:59" + "' ";
                    sqlwhere += " group by deliveryid,d.materialcode,vendorcode,vendorname,org_id order by 生成条码时间 ";
                    sql = sql + sqlwhere;
                    DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
                    gridControl.DataSource = dt;
                }

            }

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

        private void btnToExcel_Click(object sender, EventArgs e)
        {
            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null)
                return;
            if (dt.Rows.Count <= 0)
                return;
            string fileName = ShowSaveFileDialog("Microsoft Excel 2007 Document", "Microsoft Excel|*.xlsx");
            if (fileName == string.Empty) return;
            ExportToEx(fileName, "xlsx", gridView);
            OpenFile(fileName);
        }

        private void MaterialWaitForQC_Load(object sender, EventArgs e)
        {
            gridView.PrintExportProgress += new ProgressChangedEventHandler(gridView_PrintExportProgress);
        }

        private void gridView_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
        {
        
            try
            {
                DataTable dt = gridControl.DataSource as DataTable;
                if (dt == null || dt.Rows.Count < 1)
                {
                    return;
                }

                if (double.Parse(gridView.GetDataRow(e.RowHandle)["周期"].ToString()) > 24)
                {
                    e.Appearance.BackColor = Color.Red;
                }
                else if (double.Parse(gridView.GetDataRow(e.RowHandle)["周期"].ToString()) > 12 && double.Parse(gridView.GetDataRow(e.RowHandle)["周期"].ToString()) <= 24)
                {
                    e.Appearance.BackColor = Color.YellowGreen;
                }
                else if (double.Parse(gridView.GetDataRow(e.RowHandle)["周期"].ToString()) > 8 && double.Parse(gridView.GetDataRow(e.RowHandle)["周期"].ToString()) <= 12)
                {
                    e.Appearance.BackColor = Color.Yellow;
                }
                else if (double.Parse(gridView.GetDataRow(e.RowHandle)["周期"].ToString()) > 0 && double.Parse(gridView.GetDataRow(e.RowHandle)["周期"].ToString()) <= 8)
                {
                    e.Appearance.BackColor = Color.White;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}