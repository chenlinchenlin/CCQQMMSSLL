using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
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
    public partial class EMSOtherRecList : DevExpress.XtraEditors.XtraForm
    {
        public EMSOtherRecList()
        {
            InitializeComponent();
        }
        public EMSOtherRecList(string receptid, string materialcode, string Reporttype, string id)
        {

            InitializeComponent();
            databind.DataSource = null;
            //lblinfo.Text = "接收单号:" + receptid + "料号:" + materialcode + "," + Reporttype;
            this.Text = "接收单号:" + receptid + "料号:" + materialcode + "," + Reporttype;
            if (Reporttype == "入库明细")
            {
                string SOracle = "select ORGANIZATION_ID 组织id,TRANSACTION_REFERENCE 单号,SUBINVENTORY_CODE 子仓库,TRANSACTION_QUANTITY 数量,BARCODE_LOT 批次号,TRANSACTION_DATE 入库时间,BARCODE_MAN 入库人 from apps.CUX_MTL_TRANSACTIONS_V where TRANSACTION_REFERENCE='" + receptid + "' and INVENTORY_ITEM_ID='" + id + "'";
                databind.DataSource = DbAccess.SelectByOracle(SOracle).Tables[0];
            }
            else
                databind.DataSource =delivery_emsotherrecReport(Reporttype, receptid, "", "", "", "", materialcode, "", "", "", "").Tables[0];
        }

        public static DataSet delivery_emsotherrecReport(string opertype, string receptid, string date1, string date2, string vendorcode, string vendorname, string materialcode, string materialname, string states, string remarks, string po)
        {
            SqlParameter[] para = new SqlParameter[11];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@receptid", receptid);
            para[2] = new SqlParameter("@date1", date1);
            para[3] = new SqlParameter("@date2", date2);
            para[4] = new SqlParameter("@vendorcode", vendorcode);
            para[5] = new SqlParameter("@vendorname", vendorname);
            para[6] = new SqlParameter("@materialcode", materialcode);
            para[7] = new SqlParameter("@materialname", materialname);
            para[8] = new SqlParameter("@states", states);
            para[9] = new SqlParameter("@remarks", remarks);
            para[10] = new SqlParameter("@po", po);
            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "Delivery_EMSOtherRecReport", para);
        }


        private string ShowSaveFileDialog(string title, string filter)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            string name = "EMS明细列表";
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

            DataTable dt = databind.DataSource as DataTable;
            if (dt == null)
                return;
            if (dt.Rows.Count <= 0)
                return;
            string fileName = ShowSaveFileDialog("Microsoft Excel 2007 Document", "Microsoft Excel|*.xlsx");
            if (fileName == string.Empty) return;
            ExportToEx(fileName, "xlsx", gridView);
            OpenFile(fileName);
        }

        private void EMSOtherRecList_Load(object sender, EventArgs e)
        {
            gridView.PrintExportProgress += new ProgressChangedEventHandler(gridView_PrintExportProgress);
        }
    }
}