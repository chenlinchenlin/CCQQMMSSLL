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
using DX_QMS.Common;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraEditors;

namespace DX_QMS
{
    public partial class ESDProgSet : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private Microsoft.Office.Interop.Excel.ApplicationClass appClsExcel = null;
        DataTable dtsubitem = null,selectsubtable = null;
        IQC ic = new IQC();
        public ESDProgSet()
        {
            InitializeComponent();
            DataTable dt, dt1, dt2, dt3;
            try
            {
                dt = bindTypesByName("楼层");
                txtblock.Properties.Items.Clear();
                foreach (DataRow row in dt.Rows)
                {
                    txtblock.Properties.Items.Add(row["Dvalue"]);
                }
                txtblock.SelectedIndex = 0;


                dt1 = bindTypesByName("线别");
                txtline.Properties.Items.Clear();
                foreach (DataRow row in dt1.Rows)
                {
                    txtline.Properties.Items.Add(row["Dvalue"]);
                }
                txtline.SelectedIndex = 0;

                //txttesttype
                dt2 = bindTypesByName("测试类别");
                txttesttype.Properties.Items.Clear();
                foreach (DataRow row in dt2.Rows)
                {
                    txttesttype.Properties.Items.Add(row["Dvalue"]);
                }
                txttesttype.SelectedIndex = 0;
         

                dt3 = bindTypesByName("测试类别子项");
                selectsubtable = dt3;
                txttestsubitem.Properties.Items.Clear();
                foreach (DataRow row in dt3.Rows)
                {
                    txttestsubitem.Properties.Items.Add(row["Dvalue"]);
                }
                txttestsubitem.SelectedIndex = 0;

            }
            catch
            {
            }
            setRule();
            gridView.PrintExportProgress += new ProgressChangedEventHandler(gridView_PrintExportProgress);
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
            Dictionary<string, bool> dic = GroupPermission.QMS_SelectRulesForForm(post, "测试设置");
            this.btnsave.Enabled = bool.Parse(dic["hasInsert"].ToString());
            this.btndel.Enabled = bool.Parse(dic["hasDelete"].ToString());
            this.btntoexcel.Enabled = bool.Parse(dic["hasPrint"].ToString());
        }
        public abstract class ScienceCount
        {
            public static string KXJSF(double num)
            {
                double bef = System.Math.Abs(num);
                int aft = 0;
                while (bef >= 10 || (bef < 1 && bef != 0))
                {
                    if (bef >= 10)
                    {
                        bef = bef / 10;
                        aft++;
                    }
                    else
                    {
                        bef = bef * 10;
                        aft--;
                    }
                }
                return string.Concat(num >= 0 ? "" : "-", ReturnBef(bef), "E", ReturnAft(aft));
            }

            public static string ReturnBef(double bef)
            {
                if (bef.ToString() != null)
                {
                    char[] arr = bef.ToString().ToCharArray();
                    switch (arr.Length)
                    {
                        case 1:
                        case 2: return string.Concat(arr[0], ".", "00"); break;
                        case 3: return string.Concat(arr[0] + "." + arr[2] + "0"); break;
                        default: return string.Concat(arr[0] + "." + arr[2] + arr[3]); break;
                    }
                }
                else
                {
                    return "000";
                }
            }
            public static string ReturnAft(int aft)
            {
                if (aft.ToString() != null)
                {
                    string end;
                    char[] arr = System.Math.Abs(aft).ToString().ToCharArray();
                    switch (arr.Length)
                    {
                        case 1: end = "00" + arr[0]; break;
                        case 2: end = "0" + arr[0] + arr[1]; break;
                        default: end = System.Math.Abs(aft).ToString(); break;
                    }
                    return string.Concat(aft >= 0 ? "+" : "-", end);
                }
                else
                {
                    return "+000";
                }
            }
        }
        private void txtLowervalue_Leave(object sender, EventArgs e)
        {
            if (txtLowervalue.Text.Trim() == "")
                return;
            float m = 0;
            if (!float.TryParse(txtLowervalue.Text == "" ? "0" : txtLowervalue.Text, out m))
            {
                MessageBox.Show("数值格式不正确", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                txtLowervalue.Text = "";
                txtLowervalue.Focus();
                return;
            }
            this.txtLowervalueSI.Text = ScienceCount.KXJSF(double.Parse(this.txtLowervalue.Text));
        }

        private void txtUppervalue_Leave(object sender, EventArgs e)
        {
            if (txtUppervalue.Text.Trim() == "")
                return;
            float m = 0;
            if (!float.TryParse(txtUppervalue.Text == "" ? "0" : txtUppervalue.Text, out m))
            {
                MessageBox.Show("数值格式不正确", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                txtUppervalue.Text = "";
                txtUppervalue.Focus();
                return;
            }
            this.txtUppervalueSI.Text = ScienceCount.KXJSF(double.Parse(this.txtUppervalue.Text));
        }

        private void btnsave_Click(object sender, EventArgs e)
        {
            float n = 0;
            if (!float.TryParse(txtLowervalue.Text == "" ? "0" : txtLowervalue.Text, out n))
            {
                MessageBox.Show("测试值范围值不正确", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                txtLowervalue.Text = "";
                txtLowervalue.Focus();
                return;
            }
            if (!float.TryParse(txtUppervalue.Text == "" ? "0" : txtUppervalue.Text, out n))
            {
                MessageBox.Show("测试值范围值不正确", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                txtUppervalue.Text = "";
                txtUppervalue.Focus();
                return;
            }
            int i = ic.AddNewESDProgRecord("新增", txtblock.Text, txtline.Text, txttesttype.Text, txttestsubitem.Text,
                txtUppervalue.Text == "" ? "0" : txtUppervalue.Text, txtLowervalue.Text == "" ? "0" : txtLowervalue.Text, (int)txtcheckcycle.Value, (int)txtleadtime.Value, Login.username);
            btnsearch_Click(sender, e);
            txtLowervalue.Text = "";
            txtUppervalue.Text = "";
        }

        private void btndel_Click(object sender, EventArgs e)
        {

            if (gridView.FocusedRowHandle < 0)
                return;
            if (gridView.GetSelectedRows().Length < 1)
                return;
            for (int k = gridView.GetSelectedRows().Length; k > 0; k--)
            {

                DataRow dr = gridView.GetDataRow(gridView.GetSelectedRows()[k - 1]);
                int i = ic.AddNewESDProgRecord("删除", dr["楼层"].ToString(), dr["线别"].ToString(), dr["测试类别"].ToString(),
                                                dr["测试子项目"].ToString(), "", "", 0, 0, "");
            }
            btnsearch_Click(sender, e);
        }
        private DataTable BindESDRecord(string sblock, string sline, string testtype, string testsubitem)
        {
            string sql = "select Block 楼层,line 线别, Dtype 测试类别, DsubType 测试子项目,UpValue 上限, LowValue 下限,checkcycle 测试周期,leadtime 提前期,AuditingByUser 审核人 ";
            sql += " from ESD_TestProgSet ";
            sql += " where Block= case '" + sblock + "' when '' then Block else '" + sblock + "' end";
            sql += " and line= case '" + sline + "' when '' then line else '" + sline + "' end";
            return DbAccess.SelectBySql(sql).Tables[0];
        }
        private void btnsearch_Click(object sender, EventArgs e)
        {
            string sblock, sline, stype, subtype;
            if (txtblock.Text == "")
                sblock = "";
            else
                sblock = txtblock.Text;
            if (txtline.Text == "")
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
            string name = "信息";
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

        private void databind_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {

                //txtblock.Text = databind.CurrentRow.Cells["楼层"].Value.ToString();
                //txtline.SelectedValue = databind.CurrentRow.Cells["线别"].Value.ToString();
                //txttesttype.SelectedValue = databind.CurrentRow.Cells["测试类别"].Value.ToString();
                //txttestsubitem.SelectedValue = databind.CurrentRow.Cells["测试子项目"].Value.ToString();
                //txtcheckcycle.Value = int.Parse(databind.CurrentRow.Cells["测试周期"].Value.ToString() == "" ? "0" : databind.CurrentRow.Cells["测试周期"].Value.ToString());
                //txtleadtime.Value = int.Parse(databind.CurrentRow.Cells["提前期"].Value.ToString() == "" ? "0" : databind.CurrentRow.Cells["提前期"].Value.ToString());
                //txtLowervalue.Text = databind.CurrentRow.Cells["下限"].Value.ToString();
                //txtUppervalue.Text = databind.CurrentRow.Cells["上限"].Value.ToString();
            }
            catch
            {
            }
        }

        private void databind_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                //txtblock.Text = databind.CurrentRow.Cells["楼层"].Value.ToString();
                //txtline.SelectedValue = databind.CurrentRow.Cells["线别"].Value.ToString();
                //txttesttype.SelectedValue = databind.CurrentRow.Cells["测试类别"].Value.ToString();
                //txttestsubitem.SelectedValue = databind.CurrentRow.Cells["测试子项目"].Value.ToString();
                //txtcheckcycle.Value = int.Parse(databind.CurrentRow.Cells["测试周期"].Value.ToString() == "" ? "0" : databind.CurrentRow.Cells["测试周期"].Value.ToString());
                //txtleadtime.Value = int.Parse(databind.CurrentRow.Cells["提前期"].Value.ToString() == "" ? "0" : databind.CurrentRow.Cells["提前期"].Value.ToString());
                //txtLowervalue.Text = databind.CurrentRow.Cells["下限"].Value.ToString();
                //txtUppervalue.Text = databind.CurrentRow.Cells["上限"].Value.ToString();
            }
            catch
            {
            }
        }

        private void gridView_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {

            try
            {
               //// gridView.GetFocusedRowCellValue("产品编码").ToString();
                txtblock.Text = gridView.GetFocusedRowCellValue("楼层").ToString();
                txtline.Text = gridView.GetFocusedRowCellValue("线别").ToString();
                txttesttype.Text = gridView.GetFocusedRowCellValue("测试类别").ToString();
                txttestsubitem.Text = gridView.GetFocusedRowCellValue("测试子项目").ToString();
                txtcheckcycle.Value = int.Parse(gridView.GetFocusedRowCellValue("测试周期").ToString() == "" ? "0" : gridView.GetFocusedRowCellValue("测试周期").ToString());
                txtleadtime.Value = int.Parse(gridView.GetFocusedRowCellValue("提前期").ToString() == "" ? "0" : gridView.GetFocusedRowCellValue("提前期").ToString());
                txtLowervalue.Text = gridView.GetFocusedRowCellValue("下限").ToString();
                txtUppervalue.Text = gridView.GetFocusedRowCellValue("下限").ToString();

            }
            catch
            {
            }


        }

        private void gridView_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {
            try
            {
                //// gridView.GetFocusedRowCellValue("产品编码").ToString();
                txtblock.Text = gridView.GetFocusedRowCellValue("楼层").ToString();
                txtline.Text = gridView.GetFocusedRowCellValue("线别").ToString();
                txttesttype.Text = gridView.GetFocusedRowCellValue("测试类别").ToString();
                txttestsubitem.Text = gridView.GetFocusedRowCellValue("测试子项目").ToString();
                txtcheckcycle.Value = int.Parse(gridView.GetFocusedRowCellValue("测试周期").ToString() == "" ? "0" : gridView.GetFocusedRowCellValue("测试周期").ToString());
                txtleadtime.Value = int.Parse(gridView.GetFocusedRowCellValue("提前期").ToString() == "" ? "0" : gridView.GetFocusedRowCellValue("提前期").ToString());
                txtLowervalue.Text = gridView.GetFocusedRowCellValue("下限").ToString();
                txtUppervalue.Text = gridView.GetFocusedRowCellValue("下限").ToString();

            }
            catch
            {
            }
        }

        private void selectsubitem_Click(object sender, EventArgs e)
        {
            //dt3 = bindTypesByName("测试类别子项");
            if (subitem.Text.Trim() == "")
                return;
            DataTable dtnew = selectsubtable.Clone();
            DataRow[] rw = selectsubtable.Select("Dvalue like '" + subitem.Text.Trim() + "%'");
            for (int i = 0; i < rw.Length; i++)
            {
                dtnew.Rows.Add(rw[i].ItemArray);
            }
            txttestsubitem.Properties.Items.Clear();
            foreach (DataRow row in dtnew.Rows)
            {
                txttestsubitem.Properties.Items.Add(row["Dvalue"]);
            }
            txttestsubitem.SelectedIndex = 0;
        }
    }
}