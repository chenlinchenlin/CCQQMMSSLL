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
using LabelManager2;
using System.Reflection;
using System.Runtime.InteropServices;

namespace DX_QMS
{
    public partial class ESDTestLabel : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        DataTable dtsubitem = null;
        public ESDTestLabel()
        {
            InitializeComponent();

            DataTable dt;
            try
            {
                dt = bindTypesByName("楼层");
                txtblock.Properties.Items.Clear();
                foreach (DataRow row in dt.Rows)
                {
                    txtblock.Properties.Items.Add(row["Dvalue"]);
                }
                txtblock.SelectedIndex = 0;
            }
            catch
            {
            }
           ///// setRule();

        }

        private DataTable bindTypesByName(string types)
        {
            string sql = "select Dvalue from ESD_TypeDefine where Dtype='" + types + "'";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            return dt;
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
            Dictionary<string, bool> dic = GroupPermission.QMS_SelectRulesForForm(post, "标签打印");
            this.btnsave.Enabled = bool.Parse(dic["hasInsert"].ToString());
            this.btnprint.Enabled = bool.Parse(dic["hasPrint"].ToString());
            this.btndel.Enabled = bool.Parse(dic["hasDelete"].ToString());
        }

        private void btnMB1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofdImport = new OpenFileDialog();
            ofdImport.Filter = "标签文件(*.lab)|*.lab";
            ofdImport.Multiselect = false;
            DialogResult dr = ofdImport.ShowDialog();
            if (dr == DialogResult.Cancel) return;
            this.txtMB1.Text = ofdImport.FileName;
        }

        private void bindSubType(string block, string line, string sType)
        {
            string sql = "select DsubType,UpValue, LowValue from ESD_TestProgSet where line='" + line + "' and block='" + block + "' and Dtype='" + sType + "'";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            if (dt.Rows.Count > 0)
            {
                dtsubitem = dt;
                txttestsubitem.Properties.Items.Clear();
                foreach (DataRow row in dt.Rows)
                {
                    txttestsubitem.Properties.Items.Add(row["DsubType"]);
                }
               ////////// txttestsubitem.SelectedIndex = 0;

                if (txttestsubitem.Properties.Items.Count == 1)
                {
                    txttestsubitem.SelectedIndex = 0;
                    DataRow[] rw = dtsubitem.Select("DsubType='" + this.txttestsubitem.Text.Trim() + "'");
                }
                else
                {
                    txttestsubitem.Focus();
                    return;
                }
            }
            else
            {
                txttestsubitem.Properties.Items.Clear();
            }
        }

        private void bindType(string block, string line)
        {
            string sql = "select Dtype from ESD_TestProgSet where line='" + line + "' and block='" + block + "'";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            if (dt.Rows.Count > 0)
            {
                DataTable newdt = dt.Clone();
                newdt.Rows.Add(dt.Rows[0].ItemArray);
                for (int i = 1; i < dt.Rows.Count; i++)
                {
                    bool flag = true;
                    foreach (DataRow dr in newdt.Rows)
                    {
                        if (dt.Rows[i]["Dtype"].ToString() == dr["Dtype"].ToString())
                        {
                            flag = false;
                            continue;
                        }
                    }
                    if (flag)
                        newdt.Rows.Add(dt.Rows[i].ItemArray);
                }
                txttesttype.Properties.Items.Clear();
                foreach (DataRow row in newdt.Rows)
                {
                    txttesttype.Properties.Items.Add(row["Dtype"]);
                }
                txttesttype.SelectedIndex = 0;

                bindSubType(txtblock.Text, txtline.Text, txttesttype.Text);
            }
            else
            {
                txttesttype.Properties.Items.Clear();
                txttestsubitem.Properties.Items.Clear();
            }
        }
        private void bindLine(string block)
        {
            string sql = "select line from ESD_TestProgSet where Block='" + block + "'";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            if (dt.Rows.Count > 0)
            {
                DataTable newdt = dt.Clone();
                newdt.Rows.Add(dt.Rows[0].ItemArray);
                for (int i = 1; i < dt.Rows.Count; i++)
                {
                    bool flag = true;
                    foreach (DataRow dr in newdt.Rows)
                    {
                        if (dt.Rows[i]["line"].ToString() == dr["line"].ToString())
                        {
                            flag = false;
                            continue;
                        }
                    }
                    if (flag)
                        newdt.Rows.Add(dt.Rows[i].ItemArray);
                }

                txtline.Properties.Items.Clear();
                foreach (DataRow row in newdt.Rows)
                {
                    txtline.Properties.Items.Add(row["line"]);
                }
                txtline.SelectedIndex = 0;

                bindType(txtblock.Text.Trim(), txtline.Text.Trim());
            }
            else
            {
                txtline.Properties.Items.Clear();
                txttesttype.Properties.Items.Clear();
                txttestsubitem.Properties.Items.Clear();
            }
        }
        private void txtblock_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txtblock.Text .Trim () == "")
                return;
            bindLine(txtblock.Text.Trim());
        }

        private void txtline_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txtline.Text == "")
                return;
            bindType(txtblock.Text, txtline.Text);
        }

        private void txttesttype_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txttesttype.Text == "")
                return;
            bindSubType(txtblock.Text, txtline.Text, txttesttype.Text);
        }

        private void txttestsubitem_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txttestsubitem.Text == "")
                return;
            try
            {
                DataRow[] rw = dtsubitem.Select("DsubType='" + this.txttestsubitem.Text.Trim() + "'");
            }
            catch
            {
            }
        }

        private void btnsave_Click(object sender, EventArgs e)
        {
            if (txtblock.Text == "" || txtline.Text == "" || txttesttype.Text == "" || txttestsubitem.Text == "")
                return;
            string sql = "if not exists(select 1 from ESD_LotInfo where block='" + txtblock.Text + "' and line='" + txtline.Text + "'";
            sql += " and testtype='" + txttesttype.Text + "' and testsubitem='" + txttestsubitem.Text + "') ";
            sql += " insert into ESD_LotInfo(lotno, block, line, testtype, testsubitem, OperUser, OperDate) ";
            sql += " values(Convert(varchar(2),right(year(getdate()),2))+Convert(varchar(2),datepart(MM,getdate()))+Convert(varchar(2),datepart(DD,getdate()))+Convert(varchar(2),datepart(HH,getdate()))+Convert(varchar(2),datepart(MI,getdate()))+Convert(varchar(2),datepart(SS,getdate()))+Convert(varchar(3),datepart(MS,getdate())),'";
            sql += txtblock.Text + "','" + txtline.Text + "','" + txttesttype.Text + "','" + txttestsubitem.Text + "','" + Login.username + "',getdate())";

            bool F = DbAccess.ExecuteSql(sql);
            if (F)
                MessageBox.Show("生成成功!");
            btnsearch_Click(sender, e);

        }

        private void btndel_Click(object sender, EventArgs e)
        {
            if (gridView.RowCount <= 0)
                return;
            if (gridView.GetSelectedRows().Length < 1)
                return;

            for (int k = gridView.GetSelectedRows().Length; k > 0; k--)
            {
                DataRow dr = gridView.GetDataRow(gridView.GetSelectedRows()[k - 1]);

                string sql = "delete ESD_LotInfo where block='" + dr["楼层"].ToString() + "' and line='" + dr["线别"].ToString() + "'";
                sql += " and testtype='" + dr["测试类别"].ToString() + "' and testsubitem='" + dr["测试子项目"].ToString() + "'";
                DbAccess.ExecuteSql(sql);
            }
            btnsearch_Click(sender, e);
        }
        private DataTable BindESDRecord(string sblock, string sline, string testtype, string testsubitem)
        {
            string sql = "select lotno 条码号,l.Block 楼层,l.line 线别, testtype 测试类别, testsubitem 测试子项目,OperUser 操作人, OperDate 操作日期 ";
            sql += " from ESD_LotInfo l ";
            sql += " where l.Block= case '" + sblock + "' when '' then l.Block else '" + sblock + "' end";
            sql += " and l.line= case '" + sline + "' when '' then l.line else '" + sline + "' end";
            sql += " and testtype= case '" + testtype + "' when '' then testtype else '" + testtype + "' end";
            sql += " and testsubitem= case '" + testsubitem + "' when '' then testsubitem else '" + testsubitem + "' end";
            return DbAccess.SelectBySql(sql).Tables[0];
        }
        private void btnsearch_Click(object sender, EventArgs e)
        {
            string sblock, sline, stype, subtype;
            if (txtblock.Text.Trim() == "")
                sblock = "";
            else
                sblock = txtblock.Text.Trim ();
            if (txtline.Text.Trim() == "")
                sline = "";
            else
                sline = txtline.Text;
            if (txttesttype.Text == "")
                stype = "";
            else
                stype = txttesttype.Text;
            if (txttestsubitem.Text == "")
                subtype = "";
            else
                subtype = txttestsubitem.Text;

            databind.DataSource = null;
            DataTable dt = BindESDRecord(sblock, sline, stype, subtype);
            if (dt.Rows.Count > 0)
                databind.DataSource = dt;
        }
        private string ShowSaveFileDialog(string title, string filter)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            string name = "ESD信息";
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
                    DevExpress.XtraEditors.XtraMessageBox.Show(this, "Cannot find an application on your system suitable for openning the file with exported data.", System.Windows.Forms.Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
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
        private void btntoexcel_Click(object sender, EventArgs e)
        {
            DataTable dt = databind.DataSource as DataTable;
            if (dt == null)
                return;
            if (dt.Rows.Count <= 0) return;

            string fileName = ShowSaveFileDialog("Microsoft Excel 2007 Document", "Microsoft Excel|*.xlsx");
            if (fileName == string.Empty) return;
            ExportToEx(fileName, "xlsx", gridView);
            OpenFile(fileName);
        }

        private void ESDTestLabel_Load(object sender, EventArgs e)
        {
            gridView.PrintExportProgress += new ProgressChangedEventHandler(gridView_PrintExportProgress);
        }

        public  void PrintESDLabel(string block, string line, string fileName, string testtype, string testsubitem, string lotno)
        {
            ApplicationClass pLabelAppClsCarton = new ApplicationClass();
            try
            {
                System.Windows.Forms.Application.DoEvents();
                pLabelAppClsCarton.Documents.Open(fileName, false);
                Document doc = pLabelAppClsCarton.ActiveDocument;
                doc.Variables.FormVariables.Item("block").Value = block;
                doc.Variables.FormVariables.Item("line").Value = line;
                doc.Variables.FormVariables.Item("testtype").Value = testtype;
                doc.Variables.FormVariables.Item("testsubitem").Value = testsubitem;
                doc.Variables.FormVariables.Item("lotno").Value = lotno;
                doc.PrintDocument(1);
                System.Windows.Forms.Application.DoEvents();
                ((IDocument)doc).Close(false);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                System.Windows.Forms.Application.DoEvents();
                pLabelAppClsCarton.Quit();
            }
        }


        private void btnprint_Click(object sender, EventArgs e)
        {
            if (txtMB1.Text.Trim() == "")
            {
                this.lblinfo.Text = "请先选择标签模板1，无法打印!";
                lblinfo.ForeColor = Color.Red;
                return;
            }

            DataTable dt = databind.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
                return;

            if (gridView.GetSelectedRows().Length < 1)
                return;

            for (int k = gridView.GetSelectedRows().Length; k > 0; k--)
            {
                DataRow dr = gridView.GetDataRow(gridView.GetSelectedRows()[k - 1]);
                PrintESDLabel(dr["楼层"].ToString(), dr["线别"].ToString(), txtMB1.Text,
                 dr["测试类别"].ToString(), dr["测试子项目"].ToString(), dr["条码号"].ToString());
            }
        }
    }
}