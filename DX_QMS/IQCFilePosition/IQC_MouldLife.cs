using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Windows.Forms;
using DevExpress.XtraBars;
using System.IO;
using System.Diagnostics;
using System.Data.SqlClient;
using DX_QMS.Common;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraEditors;
using DX_QMS.Data;
using System.Text.RegularExpressions;
using System.Net;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace DX_QMS.IQCFilePosition
{
    public partial class IQC_MouldLife : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public IQC_MouldLife()
        {
            InitializeComponent();
        }

        private void IQC_MouldLife_Load(object sender, EventArgs e)
        {
           

            if (Login.post == "SQE" || Login.manager == "IQC管理员" || Login.manager == "IT管理员")
            {
                sBtnimport.Enabled = true;
                sBtnsubmit.Enabled = true;
                sBtnsave.Enabled = true;
                sBtndelete.Enabled = true;
            }
            else
            {
                sBtnimport.Enabled = false ;
                sBtnsubmit.Enabled = false;
                sBtnsave.Enabled = false;
                sBtndelete.Enabled = false;
            }
            gridView.PrintExportProgress += new ProgressChangedEventHandler(gridView_PrintExportProgress);
        }

        private void txtCavityQty_Leave(object sender, EventArgs e)
        {
            //int outputQty = 0;
            //if (!int.TryParse(txtCavityQty.Text.Trim() == "" ? "0" : txtCavityQty.Text, out outputQty))
            //{
            //    MessageBox.Show("模穴数不正确", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //    return;
            //}
        }

        private void txtMouldLife_Leave(object sender, EventArgs e)
        {
            int outputQty = 0;
            if (!int.TryParse(txtMouldLife.Text.Trim() == "" ? "0" : txtMouldLife.Text, out outputQty))
            {
                MessageBox.Show("模具寿命格式不正确", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
        }

        private void txtRestQty_Leave(object sender, EventArgs e)
        {
            int outputQty = 0;
            if (!int.TryParse(txtRestQty.Text.Trim() == "" ? "0" : txtRestQty.Text, out outputQty))
            {
                MessageBox.Show("剩余次数不正确", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
        }

        private void txtthisTimeQty_Leave(object sender, EventArgs e)
        {
            //int outputQty = 0;
            //if (!int.TryParse(txtthisTimeQty.Text.Trim() == "" ? "0" : txtthisTimeQty.Text, out outputQty))
            //{
            //    MessageBox.Show("本次次数不正确", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //    return;
            //}
        }

        private void txtmaterialcode_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13 && txtmaterialcode.Text.Trim() != "")
            {
                sBtnselect_Click(sender,e);
            }
       }
        private void txtmaterialcode_Leave(object sender, EventArgs e)
        {
            if (txtmaterialcode.Text.Trim() != "")
            {
                sBtnselect_Click(sender, e);
            }
        }


        private void txtMouldtype_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txtMouldtype.Text == "铝壳类模具")
            {
                txtColorMode.Text = "无";
            }
            else
            {
                txtColorMode.SelectedIndex = 0;
            }
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
            ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets["信息维护"];
            int rowCount = ws.UsedRange.Rows.Count;
            int colCount = ws.UsedRange.Columns.Count;
            if (rowCount <= 0)
            {
                MessageBox.Show("文件中没有数据记录", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (colCount != 8)
            {
                MessageBox.Show("字段个数不对", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            DataTable MouldCycleTable = new IQC_MouldCycle().Tables["IQC_MouldCycle"];

            for (int iRow = 1; iRow < rowCount; iRow++)
            {
   
                if ( objList[iRow + 1, 1]!= null  && objList[iRow + 1, 2] != null)
                {
                    DataRow row;
                    row = MouldCycleTable.NewRow();
                    row["materialcode"] = objList[iRow + 1, 1].ToString();
                    row["vendorcode"] = objList[iRow + 1, 2].ToString();
                    row["Mouldcode"] = objList[iRow + 1, 3] == null ? "" : objList[iRow + 1, 3].ToString();
                    row["Mouldtype"] = objList[iRow + 1, 4] == null ? "" : objList[iRow + 1, 4].ToString();
                    row["ColorMode"] = objList[iRow + 1, 5] == null ? "" : objList[iRow + 1, 5].ToString();
                    row["MouldQty"] = objList[iRow + 1, 6] == null ? "" : objList[iRow + 1, 6].ToString();
                    row["CavityQty"] = objList[iRow + 1, 7] == null ? "" : objList[iRow + 1, 7].ToString();
                    row["MouldLife"] = objList[iRow + 1, 8] == null ? "" : objList[iRow + 1, 8].ToString();
                    //row["RestQty"] = objList[iRow + 1, 1].ToString();
                    //row["thisTimeQty"] = objList[iRow + 1, 1].ToString();
                    row["updateMan"] = Login.username;
                    MouldCycleTable.Rows.Add(row);
                }
                else
                {
                    break;
                }
            }
            gridControl.DataSource = MouldCycleTable;

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

            OpenFileDialog fd = new OpenFileDialog();
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
                    Cursor.Current = Cursors.WaitCursor;
                    OpenExcel(fd.FileName);
                    Cursor.Current = Cursors.Default;
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
            if (gridView.RowCount < 1)
                return;

            int countsuccess = 0;

            Cursor.Current = Cursors.WaitCursor;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                DataRow dr = dt.Rows[i];

                string flag = IQC_AdditemMouldLife("新增", dr["materialcode"].ToString(), dr["vendorcode"].ToString(), dr["Mouldcode"].ToString(), dr["Mouldtype"].ToString(), dr["ColorMode"].ToString(), dr["MouldQty"].ToString(),
                    dr["CavityQty"].ToString(), 
                    int.Parse(dr["MouldLife"].ToString() == "" ? "0" : dr["MouldLife"].ToString()), int.Parse(dr["MouldLife"].ToString() == "" ? "0" : dr["MouldLife"].ToString()), 0, Login.username);

                if (flag == "新增成功")
                {
                    countsuccess++;
                }
            }
            Cursor.Current = Cursors.Default;
            MessageBox.Show("提交成功" + countsuccess + "条,失败" + (dt.Rows.Count - countsuccess) + "条", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public string IQC_AdditemMouldLife(string opertype,string materialcode,string vendorcode, string Mouldcode,string Mouldtype,string ColorMode,string MouldQty,string CavityQty,int MouldLife,int RestQty,int thisTimeQty,string updateMan)
        {
            SqlParameter[] para = new SqlParameter[13];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@materialcode", materialcode);
            para[2] = new SqlParameter("@vendorcode", vendorcode);
            para[3] = new SqlParameter("@Mouldcode", Mouldcode);
            para[4] = new SqlParameter("@Mouldtype", Mouldtype);
            para[5] = new SqlParameter("@ColorMode", ColorMode);
            para[6] = new SqlParameter("@MouldQty", MouldQty);
            para[7] = new SqlParameter("@CavityQty", CavityQty);
            para[8] = new SqlParameter("@MouldLife", MouldLife);
            para[9] = new SqlParameter("@RestQty", RestQty);
            para[10] = new SqlParameter("@thisTimeQty", thisTimeQty);
            para[11] = new SqlParameter("@updateMan", updateMan);
            para[12] = new SqlParameter("@msg", SqlDbType.VarChar, 50);
            para[12].Direction = ParameterDirection.Output;
            DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "IQC_InsertMouldLife", para);
            return para[12].Value.ToString();
        }


        private void sBtnsave_Click(object sender, EventArgs e)
        {
            if (txtmaterialcode.Text.Trim() == "")
            {
                lblinfo.Text = "物料编码不能为空";
                lblinfo.ForeColor = Color.Red;
                return;
            }
            int  MouldLife = 0, RestQty= 0, thisTimeQty= 0;
            if (!int.TryParse(txtMouldLife.Text.Trim() == "" ? "0" : txtMouldLife.Text, out MouldLife))
            {
                lblinfo.Text = "模具寿命格式不正确";
                lblinfo.ForeColor = Color.Red;
                return;
            }
            string msg = IQC_AdditemMouldLife("新增",txtmaterialcode.Text.Trim(),txtvendorcode.Text,txtMouldcode.Text,txtMouldtype.Text,txtColorMode.Text,txtMouldQty.Text, txtCavityQty.Text, MouldLife,RestQty,thisTimeQty,Login.username);


            if (msg.Contains("成功"))
            {
                lblinfo.Text = "新增成功";
            }
            else
            {
                lblinfo.Text = "新增失败";
                lblinfo.ForeColor = Color.Red;
            }

            sBtnselect_Click(sender,e);

        }
        private void sBtnselect_Click(object sender, EventArgs e)
        {

            string sql = "",where = "  where 1=1 ";
            string materialcode= txtmaterialcode.Text;
            string vendorcode = txtvendorcode.Text;
            string Mouldcode = txtMouldcode.Text;
            string Mouldtype = txtMouldtype.Text;
            string MouldQty = txtMouldQty.Text;
            string ColorMode = txtColorMode.Text;

            if (!string.IsNullOrEmpty(materialcode))
            {
                where += " and materialcode = '"+materialcode+"' ";
            }
            if (!string.IsNullOrEmpty(vendorcode))
            {
                where += " and vendorcode = '" + vendorcode + "' ";
            }
            if (!string.IsNullOrEmpty(Mouldcode))
            {
                where += " and Mouldcode = '" + Mouldcode + "' ";
            }
            if (!string.IsNullOrEmpty(Mouldtype))
            {
                where += " and Mouldtype = '" + Mouldtype + "' ";
            }
            if (!string.IsNullOrEmpty(MouldQty))
            {
                where += " and MouldQty = '" + MouldQty + "' ";
            }
            if (!string.IsNullOrEmpty(ColorMode))
            {
                where += " and ColorMode = '" + ColorMode + "' ";
            }
            sql = "  select materialcode,vendorcode ,Mouldcode ,Mouldtype ,ColorMode ,MouldQty ,CavityQty ,MouldLife ,RestQty ,thisTimeQty ,updateMan ,updateTime  ";
            sql +=" from IQC_MouldCycle "+ where + " order by updateTime desc  ";

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

        private void sBtndelete_Click(object sender, EventArgs e)
        {
            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
                return;
            if (gridView.RowCount < 1)
                return;
            if (gridView.FocusedRowHandle < 0)
            {
                MessageBox.Show("请选择要删除的行记录", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            string materialcode = gridView.GetFocusedRowCellValue("materialcode").ToString();
            string vendorcode = gridView.GetFocusedRowCellValue("vendorcode").ToString();
            string Mouldcode = gridView.GetFocusedRowCellValue("Mouldcode").ToString();

            string msg = IQC_AdditemMouldLife("删除", materialcode,vendorcode, Mouldcode, "", "", "","", 0, 0, 0, "");

            if (msg.Contains("成功"))
            {
                lblinfo.Text = "删除成功";
                gridControl.DataSource = null;
            }
            else
            {
                lblinfo.Text = "删除失败";
                lblinfo.ForeColor = Color.Red;
            }

        }

        private void sBtnreset_Click(object sender, EventArgs e)
        {
            txtmaterialcode.Text = "";
            txtvendorcode.Text = "";
            txtMouldcode.Text = "";
            txtMouldtype.Text = "";
            txtMouldQty.Text = "";
            txtCavityQty.Text = "";
            txtMouldLife.Text = "";
            txtRestQty.Text = "";
            txtColorMode.Text = "";
            lblinfo.Text = "";
            gridControl.DataSource = null;
            gridControlmail.DataSource = null;
        }

        private void gridView_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {

            try
            {
                txtmaterialcode.Text = gridView.GetFocusedRowCellValue("materialcode").ToString();
                txtvendorcode.Text = gridView.GetFocusedRowCellValue("vendorcode").ToString();
                txtMouldcode.Text = gridView.GetFocusedRowCellValue("Mouldcode").ToString();
                txtMouldtype.Text = gridView.GetFocusedRowCellValue("Mouldtype").ToString();
                txtColorMode.Text = gridView.GetFocusedRowCellValue("ColorMode").ToString();
                txtMouldQty.Text = gridView.GetFocusedRowCellValue("MouldQty").ToString();
                txtCavityQty.Text = gridView.GetFocusedRowCellValue("CavityQty").ToString();
                txtMouldLife.Text = gridView.GetFocusedRowCellValue("MouldLife").ToString();
                txtRestQty.Text = gridView.GetFocusedRowCellValue("RestQty").ToString();
            }
            catch
            {

            }
        }


        private string ShowSaveFileDialog(string title, string filter)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            string name = "模具寿命信息记录";
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

        private void sBtnselectmail_Click(object sender, EventArgs e)
        {
            string sql = " select mailType ,userName ,userMail  from MailGroup where mailType = '模料管理'   ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            if (dt != null && dt.Rows.Count > 0)
            {
                gridControlmail.DataSource = dt;
            }
            else
            {
                gridControlmail.DataSource = null;
            }
        }

        private void sBtnaddmail_Click(object sender, EventArgs e)
        {
            if (mailname.Text.Trim() == "" || txtmail.Text.Trim() == "")
                return;

            Regex regex = new Regex(@"\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*");
            bool mailfalg= regex.IsMatch(txtmail.Text.Trim());
            if (!mailfalg)
            {
                MessageBox.Show("邮箱地址格式不正确","提醒",MessageBoxButtons.OK,MessageBoxIcon.Information);
                return;
            }
            string sql = " if not exists( select 1 from  MailGroup where mailType = '模料管理' and userName = '"+ mailname.Text+ "'  )  ";
                   sql += "  insert into MailGroup (deptid ,mailType ,userName ,userMail ,sendType ,operUser ,operDate)  ";
                   sql += "   values ( 'Quality','模料管理','"+mailname.Text+ "','"+ txtmail.Text+ "','to','"+Login.username+"',GETDATE())  ";
            bool falg = DbAccess.ExecuteSql(sql);
            if (falg)
            {
                MessageBox.Show("添加成功","提醒",MessageBoxButtons.OK,MessageBoxIcon.Information);
                sBtnselectmail_Click(null , null);
            }
        }

        private void sBtndeletemail_Click(object sender, EventArgs e)
        {
            DataTable dt = gridControlmail.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
                return;
            if (gridViewmail.RowCount < 1)
                return;
            if (gridViewmail.FocusedRowHandle < 0)
            {
                MessageBox.Show("请选择要删除的行记录", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            string mailType = gridViewmail.GetFocusedRowCellValue("mailType").ToString();
            string userName = gridViewmail.GetFocusedRowCellValue("userName").ToString();
            string userMail = gridViewmail.GetFocusedRowCellValue("userMail").ToString();

            string sql = " delete MailGroup where  deptid = 'Quality' and mailType = '" + mailType + "' and userName = '" + userName + "' and userMail = '" + userMail + "'   ";
            bool falg = DbAccess.ExecuteSql(sql);
            if (falg)
            {
                MessageBox.Show("删除成功", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                sBtnselectmail_Click(null, null);
            }


        }


        private void sBtnsend_Click(object sender, EventArgs e)
        {
            if (gridView.RowCount < 1)
                return;
            if (gridView.FocusedRowHandle < 0)
            {
                MessageBox.Show("请选择相应的物料", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (gridViewmail.RowCount < 1)
                return;
            if (gridViewmail.FocusedRowHandle < 0)
            {
                MessageBox.Show("请选择相应的负责人", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string materialcode = gridView.GetFocusedRowCellValue("materialcode").ToString();
            string vendorcode = gridView.GetFocusedRowCellValue("vendorcode").ToString();
            string Mouldcode = gridView.GetFocusedRowCellValue("Mouldcode").ToString();
            string Mouldtype = gridView.GetFocusedRowCellValue("Mouldtype").ToString();
            string MouldQty = gridView.GetFocusedRowCellValue("MouldQty").ToString();
            string MouldLife = gridView.GetFocusedRowCellValue("MouldLife").ToString();
            string RestQty = gridView.GetFocusedRowCellValue("RestQty").ToString();


            string userName = gridViewmail.GetFocusedRowCellValue("userName").ToString();
            string userMail = gridViewmail.GetFocusedRowCellValue("userMail").ToString();

            string content = "物料编码("+ materialcode+ ")，模具为：" + MouldQty + "；类型为：" + Mouldtype + "；寿命为：" + MouldLife + "；剩余次数为：" + RestQty;

            if ((int.Parse(RestQty) < int.Parse(MouldLife) * 0.2) && (int.Parse(RestQty) > int.Parse(MouldLife) * 0.1))
            {
                content += " ；剩余量少于20%，请提醒相应的供应商 ";
            }
            else if (int.Parse(RestQty) <= int.Parse(MouldLife) * 0.1)
            {
                content += " ；剩余量少于10%，请提醒相应的供应商 ";
            }
            try
            {
                Outlook.Application olApp = new Outlook.Application();
                Outlook.MailItem mailItem = (Outlook.MailItem)olApp.CreateItem(Outlook.OlItemType.olMailItem);
                mailItem.To = userMail;
                mailItem.Subject = "QMS系统：" + DateTime.Now.ToString("yyyyMMdd") + "_模具寿命提醒";
                mailItem.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;

                mailItem.HTMLBody = content;

                ((Outlook._MailItem)mailItem).Send();
                MessageBox.Show("发送成功", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                mailItem = null;
                olApp = null;
            }
            catch
            {
                ProdTest.SendMail(userMail,"QMS系统："+DateTime.Now.ToString("yyyyMMdd") + "_模具寿命提醒", content);
            }
        }
        private void gridView_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
        {

            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
            {
                return;
            }
            if (gridView.RowCount < 1)
                return;
            try
            {
                if (int.Parse(gridView.GetDataRow(e.RowHandle)["RestQty"].ToString()) < int.Parse(gridView.GetDataRow(e.RowHandle)["MouldLife"].ToString()) * 0.2 )
                {
                    e.Appearance.BackColor = Color.Yellow;
                }
                if (int.Parse(gridView.GetDataRow(e.RowHandle)["RestQty"].ToString()) <= int.Parse(gridView.GetDataRow(e.RowHandle)["MouldLife"].ToString()) * 0.1 )
                {
                    e.Appearance.BackColor = Color.Red;
                }
            }
            catch
            {

            }
        }



    }
}