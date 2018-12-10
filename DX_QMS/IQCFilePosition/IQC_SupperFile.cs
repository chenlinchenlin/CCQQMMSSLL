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
using DX_QMS.Common;
using Outlook = Microsoft.Office.Interop.Outlook;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraEditors;
using System.IO;
using System.Diagnostics;

namespace DX_QMS.IQCFilePosition
{
    public partial class IQC_SupperFile : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public IQC_SupperFile()
        {
            InitializeComponent();
        }

        private void IQC_SupperFile_Load(object sender, EventArgs e)
        {
            txtreporttype.SelectedIndex = 0;
            txtrohsreportdate.Text = DateTime.Now.ToString("yyyy-MM-dd");
            txtrohsExpiryDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
            txtHFreportdate.Text = DateTime.Now.ToString("yyyy-MM-dd");
            txtHFExpiryDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
            receiver();
            txtresult.SelectedIndex = -1;
        }

        private void txtreporttype_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txtreporttype.Text == "RoHS测试报告")
            {
                enabledRoHS();
                disenabledHF();
            }
            else if (txtreporttype.Text == "HF检测报告")
            {
                enabledHF();
                disenabledRoHS();
            }
        }

        void disenabledRoHS()
        {
            txtrohsreportcode.Text = "";
            txtrohsreportcode.Enabled = false;
            txtrohsreportdate.Enabled = false;
            txtrohsExpiryDate.Enabled = false;
            ROHScheckallresult.Enabled = false;
            txtresult_Cd.Text = "";
            txtresult_Cd.Enabled = false;
            txtresult_Pb.Text = "";
            txtresult_Pb.Enabled = false;
            txtresult_Hg.Text = "";
            txtresult_Hg.Enabled = false;
            txtresult_Cr6.Text = "";
            txtresult_Cr6.Enabled = false;
            txtresult_PBBs.Text = "";
            txtresult_PBBs.Enabled = false;
            txtresult_PBDEs.Text = "";
            txtresult_PBDEs.Enabled = false;
            txtresult_DBP.Text = "";
            txtresult_DBP.Enabled = false;
            txtresult_BBP.Text = "";
            txtresult_BBP.Enabled = false;
            txtresult_DEHP.Text = "";
            txtresult_DEHP.Enabled = false;
            txtresult_DIBP.Text = "";
            txtresult_DIBP.Enabled = false;
        }

        void enabledRoHS()
        {
            //txtrohsreportcode.Text = "";
            txtrohsreportcode.Enabled = true;
            txtrohsreportdate.Enabled = true;
            txtrohsExpiryDate.Enabled = true;           
            ROHScheckallresult.Enabled = true;
            //txtresult_Cd.Text = "";
            txtresult_Cd.Enabled = true;
           // txtresult_Pb.Text = "";
            txtresult_Pb.Enabled = true;
           // txtresult_Hg.Text = "";
            txtresult_Hg.Enabled = true;
           // txtresult_Cr6.Text = "";
            txtresult_Cr6.Enabled = true;
           // txtresult_PBBs.Text = "";
            txtresult_PBBs.Enabled = true;
           // txtresult_PBDEs.Text = "";
            txtresult_PBDEs.Enabled = true;
           // txtresult_DBP.Text = "";
            txtresult_DBP.Enabled = true;
           // txtresult_BBP.Text = "";
            txtresult_BBP.Enabled = true;
           // txtresult_DEHP.Text = "";
            txtresult_DEHP.Enabled = true;
           // txtresult_DIBP.Text = "";
            txtresult_DIBP.Enabled = true;
        }

        void disenabledHF()
        {
            txtHFreportcode.Text = "";
            txtHFreportcode.Enabled = false;
            txtHFreportdate.Enabled = false;
            txtHFExpiryDate.Enabled = false;
            HFcheckallresult.Enabled = false;
            txtresult_F.Text = "";
            txtresult_F.Enabled = false;
            txtresult_Cl.Text = "";
            txtresult_Cl.Enabled = false;
            txtresult_Br.Text = "";
            txtresult_Br.Enabled = false;
            txtresult_I.Text = "";
            txtresult_I.Enabled = false;
        }
        void enabledHF()
        {
            //txtHFreportcode.Text = "";
            txtHFreportcode.Enabled = true;
            txtHFreportdate.Enabled = true;
            txtHFExpiryDate.Enabled = true;
            HFcheckallresult.Enabled = true;
            //txtresult_F.Text = "";
            txtresult_F.Enabled = true;
            //txtresult_Cl.Text = "";
            txtresult_Cl.Enabled = true;
            //txtresult_Br.Text = "";
            txtresult_Br.Enabled = true;
            //txtresult_I.Text = "";
            txtresult_I.Enabled = true;
        }
        void receiver()
        {
            string sql = "   select userName +': '+ userMail  收件人,userMail 邮箱地址 from MailGroup where mailType = '检测报告'  ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];        
            ListReceiver.DataSource = dt;
            ListReceiver.DisplayMember = dt.Columns["收件人"].ToString();
            ListReceiver.ValueMember = dt.Columns["邮箱地址"].ToString();
        }
        private void Btnselect_Click(object sender, EventArgs e)
        {
            string sql = "", where = " where 1=1 ";
            string supper = txtsupper.Text;
            string productcode = txtproductcode.Text;
            string productname = txtproductname.Text;
            string productmodel = txtproductmodel.Text;

            if (!string.IsNullOrEmpty(supper))
            {
                where += " and supper = '" + supper + "' ";
            }
            if (!string.IsNullOrEmpty(productcode))
            {
                where += " and productcode = '" + productcode + "' ";
            }
            if (!string.IsNullOrEmpty(productname))
            {
                where += " and productname = '" + productname + "' ";
            }
            if (!string.IsNullOrEmpty(productmodel))
            {
                where += " and productmodel = '" + productmodel + "' ";
            }

            if (txtreporttype.Text == "RoHS测试报告")
            {
                string rohsreportcode = txtrohsreportcode.Text;
                string rohsreportdate = txtrohsreportdate.DateTime.ToString("yyyy-MM-dd");
                string rohsExpiryDate = txtrohsExpiryDate.DateTime.ToString("yyyy-MM-dd");
                if (!string.IsNullOrEmpty(rohsreportcode))
                {
                    where += " and rohsreportcode = '" + rohsreportcode + "' ";
                }
                //if (!string.IsNullOrEmpty(rohsreportdate))
                //{
                //    where += " and rohsreportdate >= '" + rohsreportdate + " 00:00:00 '";
                //}
                //if (!string.IsNullOrEmpty(rohsExpiryDate))
                //{
                //    where += " and rohsExpiryDate >= '" + rohsExpiryDate + " 00:00:00 '";
                //}
                sql = "  select supper 供应商,productcode 物料编码,productname 物料名称,productmodel 型号,signdate 签署日期,rohsreportcode 报告编码,rohsreportdate 报告日期,rohsExpiryDate 到期日期, ";
                sql += "  result_Cd Cd,result_Pb Pb,result_Hg Hg,result_Cr6 Cr6,result_PBBs PBBs,result_PBDEs PBDEs,result_DBP DBP,result_BBP BBP,result_DEHP DEHP,result_DIBP DIBP,remark 备注信息,updatetime 更新时间,updateman 记录人,states 状态,AuditingByUser 审核人  ";
                sql += "  from IQC_SupperROHS    ";
                sql += where + " order by updatetime desc   ";
                DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
                gridView.Columns.Clear();
                gridControl.DataSource = null;
                gridControl.DataSource = dt;

            }
            if (txtreporttype.Text == "HF检测报告")
            {
                string HFreportcode = txtHFreportcode.Text;
                string HFreportdate = txtHFreportdate.DateTime.ToString("yyyy-MM-dd");
                string HFExpiryDate = txtHFExpiryDate.DateTime.ToString("yyyy-MM-dd");
                if (!string.IsNullOrEmpty(HFreportcode))
                {
                    where += " and HFreportcode = '" + HFreportcode + "' ";
                }
                //if (!string.IsNullOrEmpty(HFreportdate))
                //{
                //    where += " and HFreportdate >= '" + HFreportdate + " 00:00:00 '";
                //}
                //if (!string.IsNullOrEmpty(HFExpiryDate))
                //{
                //    where += " and HFExpiryDate >= '" + HFExpiryDate + " 00:00:00 '";
                //}
                sql = "   select supper 供应商,productcode 物料编码,productname 物料名称,productmodel 型号,signdate 签署日期,HFreportcode 报告编码,HFreportdate 报告日期,HFExpiryDate 到期日期,  ";
                sql += "   result_F F,result_Cl Cl,result_Br Br,result_I I,remark 备注信息,updatetime 更新时间,updateman 记录人,states 状态,AuditingByUser 审核人  ";
                sql += "   from  IQC_SupperHF     ";
                sql += where + "  order by updatetime desc  ";
                DataTable dt = DbAccess.SelectBySql(sql).Tables[0];    
                gridView.Columns.Clear();
                gridControl.DataSource = null;
                gridControl.DataSource = dt;
                //gridView.BestFitColumns(true);
            }
        }

        public string OperSupperROHS(string opertype, string supper, string productcode, string productname, string productmodel, string signdate, string rohsreportcode,string rohsreportdate, string rohsExpiryDate, 
            string result_Cd, string result_Pb, string result_Hg,string result_Cr6, string result_PBBs, string result_PBDEs, string result_DBP, string result_BBP, string result_DEHP, string result_DIBP, string remark, string updateman)
        {
            SqlParameter[] para = new SqlParameter[22];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@supper", supper);
            para[2] = new SqlParameter("@productcode", productcode);
            para[3] = new SqlParameter("@productname", productname);
            para[4] = new SqlParameter("@productmodel", productmodel);
            para[5] = new SqlParameter("@signdate", signdate);
            para[6] = new SqlParameter("@rohsreportcode", rohsreportcode);
            para[7] = new SqlParameter("@rohsreportdate", rohsreportdate);
            para[8] = new SqlParameter("@rohsExpiryDate", rohsExpiryDate);
            para[9] = new SqlParameter("@result_Cd", result_Cd);
            para[10] = new SqlParameter("@result_Pb", result_Pb);
            para[11] = new SqlParameter("@result_Hg", result_Hg);
            para[12] = new SqlParameter("@result_Cr6", result_Cr6);
            para[13] = new SqlParameter("@result_PBBs", result_PBBs);
            para[14] = new SqlParameter("@result_PBDEs", result_PBDEs);
            para[15] = new SqlParameter("@result_DBP", result_DBP);
            para[16] = new SqlParameter("@result_BBP", result_BBP);
            para[17] = new SqlParameter("@result_DEHP", result_DEHP);
            para[18] = new SqlParameter("@result_DIBP", result_DIBP);
            para[19] = new SqlParameter("@remark", remark);
            para[20] = new SqlParameter("@updateman", updateman);
            para[21] = new SqlParameter("@msg", SqlDbType.VarChar, 50);
            para[21].Direction = ParameterDirection.Output;
            DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "IQC_OperSupperROHS", para);
            return para[21].Value.ToString();
        }


        public string OperSupperHF(string opertype, string supper, string productcode, string productname, string productmodel, string signdate, string HFreportcode, string HFreportdate, string HFExpiryDate,
    string result_F, string result_Cl, string result_Br, string result_I,string remark, string updateman)
        {
            SqlParameter[] para = new SqlParameter[16];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@supper", supper);
            para[2] = new SqlParameter("@productcode", productcode);
            para[3] = new SqlParameter("@productname", productname);
            para[4] = new SqlParameter("@productmodel", productmodel);
            para[5] = new SqlParameter("@signdate", signdate);
            para[6] = new SqlParameter("@HFreportcode", HFreportcode);
            para[7] = new SqlParameter("@HFreportdate", HFreportdate);
            para[8] = new SqlParameter("@HFExpiryDate", HFExpiryDate);
            para[9] = new SqlParameter("@result_F", result_F);
            para[10] = new SqlParameter("@result_Cl", result_Cl);
            para[11] = new SqlParameter("@result_Br", result_Br);
            para[12] = new SqlParameter("@result_I", result_I);
            para[13] = new SqlParameter("@remark", remark);
            para[14] = new SqlParameter("@updateman", updateman);
            para[15] = new SqlParameter("@msg", SqlDbType.VarChar, 50);
            para[15].Direction = ParameterDirection.Output;
            DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "IQC_OperSupperHF", para);
            return para[15].Value.ToString();
        }


        private void Btninsert_Click(object sender, EventArgs e)
        {

            if (txtsupper.Text == "" || txtproductcode.Text == "" || txtproductmodel.Text== "")
                return;
            if (txtreporttype.Text == "RoHS测试报告")
            {
                string floerPath = "\\\\QMSSVR\\rohs\\"+txtsupper.Text+"\\"+ txtproductcode.Text+"\\"+txtproductmodel.Text;
                string ROHSflag = OperSupperROHS("新增", txtsupper.Text, txtproductcode.Text, txtproductname.Text, txtproductmodel.Text, txtsigndate.Text == "" ? "" : txtsigndate.DateTime.ToString("yyyy-MM-dd"), txtrohsreportcode.Text, txtrohsreportdate.DateTime.ToString("yyyy-MM-dd"), txtrohsExpiryDate.DateTime.ToString("yyyy-MM-dd"),
                txtresult_Cd.Text, txtresult_Pb.Text, txtresult_Hg.Text, txtresult_Cr6.Text, txtresult_PBBs.Text, txtresult_PBDEs.Text, txtresult_DBP.Text, txtresult_BBP.Text, txtresult_DEHP.Text, txtresult_DIBP.Text, txtRemark.Text, Login.username);
                if (ROHSflag.Contains("新增成功"))
                {
                    MessageBox.Show("新增成功", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                try
                {
                    Directory.CreateDirectory(floerPath);
                }
                catch { }
            }
            if (txtreporttype.Text == "HF检测报告")
            {
                string floerPath = "\\\\QMSSVR\\hf\\" + txtsupper.Text + "\\" + txtproductcode.Text + "\\"+txtproductmodel.Text; ;
                string HFflag = OperSupperHF("新增", txtsupper.Text, txtproductcode.Text, txtproductname.Text, txtproductmodel.Text, txtsigndate.Text == "" ? "" : txtsigndate.DateTime.ToString("yyyy-MM-dd"), txtHFreportcode.Text, txtHFreportdate.DateTime.ToString("yyyy-MM-dd"), txtHFExpiryDate.DateTime.ToString("yyyy-MM-dd"),
                  txtresult_F.Text, txtresult_Cl.Text, txtresult_Br.Text, txtresult_I.Text, txtRemark.Text,Login.username);
                if (HFflag.Contains("新增成功"))
                {
                    MessageBox.Show("新增成功", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                try
                {
                    Directory.CreateDirectory(floerPath);
                }
                catch
                {
                }
            }
               Btnselect_Click(null, null);
        }

        private void Btnupdate_Click(object sender, EventArgs e)
        {
            if (gridView.RowCount < 1)
                return;
            if (gridView.FocusedRowHandle < 0)
                return;
            if (txtsupper.Text == "" || txtproductcode.Text == "" || txtproductmodel.Text == "")
                return;

            if (txtreporttype.Text == "RoHS测试报告")
            {
                string  rohsreportcode = gridView.GetFocusedRowCellValue("报告编码").ToString();
                string ROHSflag = OperSupperROHS("更新", txtsupper.Text, txtproductcode.Text, txtproductname.Text, txtproductmodel.Text, txtsigndate.Text == "" ? "" : txtsigndate.DateTime.ToString("yyyy-MM-dd"), txtrohsreportcode.Text, txtrohsreportdate.DateTime.ToString("yyyy-MM-dd"), txtrohsExpiryDate.DateTime.ToString("yyyy-MM-dd"),
                txtresult_Cd.Text, txtresult_Pb.Text, txtresult_Hg.Text, txtresult_Cr6.Text, txtresult_PBBs.Text, txtresult_PBDEs.Text, txtresult_DBP.Text, txtresult_BBP.Text, txtresult_DEHP.Text, txtresult_DIBP.Text, txtRemark.Text, Login.username);
                if (ROHSflag.Contains("更新成功"))
                {
                    MessageBox.Show("更新成功", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else if (txtreporttype.Text == "HF检测报告")  
            {
                string HFreportcode = gridView.GetFocusedRowCellValue("报告编码").ToString();
                string HFflag = OperSupperHF("更新", txtsupper.Text, txtproductcode.Text, txtproductname.Text, txtproductmodel.Text, txtsigndate.Text == "" ? "" : txtsigndate.DateTime.ToString("yyyy-MM-dd"), txtHFreportcode.Text, txtHFreportdate.DateTime.ToString("yyyy-MM-dd"), txtHFExpiryDate.DateTime.ToString("yyyy-MM-dd"),
                txtresult_F.Text, txtresult_Cl.Text, txtresult_Br.Text, txtresult_I.Text, txtRemark.Text, Login.username);
                if (HFflag.Contains("更新成功"))
                {
                    MessageBox.Show("更新成功", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            Btnselect_Click(null,null);
        }

        private void Btndelete_Click(object sender, EventArgs e)
        {

            if (gridView.RowCount < 1)
                return;
            if (gridView.FocusedRowHandle < 0)
                return;
            if (txtsupper.Text == "" || txtproductcode.Text == "" || txtproductmodel.Text == "")
                return;
            if (txtreporttype.Text == "RoHS测试报告")
            {
                string rohsreportcode = gridView.GetFocusedRowCellValue("报告编码").ToString();
                string ROHSflag = OperSupperROHS("删除", txtsupper.Text, txtproductcode.Text, txtproductname.Text, txtproductmodel.Text, txtsigndate.Text == "" ? "" : txtsigndate.DateTime.ToString("yyyy-MM-dd"), txtrohsreportcode.Text, txtrohsreportdate.DateTime.ToString("yyyy-MM-dd"), txtrohsExpiryDate.DateTime.ToString("yyyy-MM-dd"),
                txtresult_Cd.Text, txtresult_Pb.Text, txtresult_Hg.Text, txtresult_Cr6.Text, txtresult_PBBs.Text, txtresult_PBDEs.Text, txtresult_DBP.Text, txtresult_BBP.Text, txtresult_DEHP.Text, txtresult_DIBP.Text, txtRemark.Text, Login.username);
                if (ROHSflag.Contains("删除成功"))
                {
                    MessageBox.Show("删除成功", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else if (txtreporttype.Text == "HF检测报告")
            {
                string HFreportcode = gridView.GetFocusedRowCellValue("报告编码").ToString();
                string HFflag = OperSupperHF("删除", txtsupper.Text, txtproductcode.Text, txtproductname.Text, txtproductmodel.Text, txtsigndate.Text == "" ? "" : txtsigndate.DateTime.ToString("yyyy-MM-dd"), txtHFreportcode.Text, txtHFreportdate.DateTime.ToString("yyyy-MM-dd"), txtHFExpiryDate.DateTime.ToString("yyyy-MM-dd"),
                txtresult_F.Text, txtresult_Cl.Text, txtresult_Br.Text, txtresult_I.Text, txtRemark.Text, Login.username);
                if (HFflag.Contains("删除成功"))
                {
                    MessageBox.Show("删除成功", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            Btnselect_Click(null, null);
        }

        private void Btnreset_Click(object sender, EventArgs e)
        {
            txtsupper.Text = "";
            txtproductcode.Text = "";
            txtproductname.Text = "";
            txtproductmodel.Text = "";
            txtsigndate.Text = "";
            txtrohsreportcode.Text = "";
            txtresult_Cd.Text = "";
            txtresult_Pb.Text = "";
            txtresult_Hg.Text = "";
            txtresult_Cr6.Text = "";
            txtresult_PBBs.Text = "";
            txtresult_PBDEs.Text = "";
            txtresult_DBP.Text = "";
            txtresult_BBP.Text = "";
            txtresult_DEHP.Text = "";
            txtresult_DIBP.Text = "";
            txtHFreportcode.Text = "";
            txtresult_F.Text = "";
            txtresult_Cl.Text = "";
            txtresult_Br.Text = "";
            txtresult_I.Text = "";
            txtRemark.Text = "";
            gridView.Columns.Clear();
            gridControl.DataSource = null;
        }
        private string ShowSaveFileDialog(string title, string filter)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            string name = "供应商_"+ txtreporttype.Text;
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
                    XtraMessageBox.Show(this, "Cannot find an application on your system suitable for openning the file with exported data.", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        private void Btnreport_Click(object sender, EventArgs e)
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

        private void btnAudit_Click(object sender, EventArgs e)
        {
            if (gridView.RowCount < 1)
                return;
            if (gridView.FocusedRowHandle < 0)
                return;
            if (txtresult.SelectedIndex == -1)
                return;
            string sql = "";
            bool flag = false;
            string states = "审核OK";
            if (txtresult.SelectedIndex == 0)
                states = "审核OK";
            else if (txtresult.SelectedIndex == 1)
                states = "审核NG";

            string supper = gridView.GetFocusedRowCellValue("供应商").ToString();
            string productcode = gridView.GetFocusedRowCellValue("物料编码").ToString();
            string productmodel = gridView.GetFocusedRowCellValue("型号").ToString();

            if (gridView.Columns.Count == 23)
            {                
                sql = "   update IQC_SupperROHS set states = '"+states+"' ,AuditingByUser='"+Login.username+"',AuditingDate = GETDATE() where supper = '"+supper+"' and productcode = '"+productcode+"' and productmodel ='"+productmodel+"'   ";        
            }
            else if (gridView.Columns.Count == 17)
            {
                sql = "   update IQC_SupperHF set states = '"+states+ "' ,AuditingByUser='"+Login.username+ "',AuditingDate = GETDATE() where supper = '"+supper+"' and productcode = '"+productcode+"' and productmodel ='"+productmodel+"'  ";
            }
            flag = DbAccess.ExecuteSql(sql);
            if (flag == true)
            {
                MessageBox.Show("审核成功","提醒",MessageBoxButtons.OK,MessageBoxIcon.Information);
            }
        }

        private void btnAuditQuery_Click(object sender, EventArgs e)
        {
            if (ListReceiver.SelectedIndex == -1)
                return;
            string Receiver = ListReceiver.Text.Substring(0, ListReceiver.Text.IndexOf(':'));

            string sql = " ";
            if (txtreporttype.Text == "RoHS测试报告")
            {
                sql = "  select supper 供应商,productcode 物料编码,productname 物料名称,productmodel 型号,signdate 签署日期,rohsreportcode 报告编码,rohsreportdate 报告日期,rohsExpiryDate 到期日期,   ";
                sql += "  result_Cd Cd,result_Pb Pb,result_Hg Hg,result_Cr6 Cr6,result_PBBs PBBs,result_PBDEs PBDEs,result_DBP DBP,result_BBP BBP,result_DEHP DEHP,result_DIBP DIBP,remark 备注信息,updatetime 更新时间,updateman 记录人,states 状态,AuditingByUser 审核人  ";
                sql += "  from IQC_SupperROHS  where states like '审核%' and SendState='OK'  and Receiver = '"+Receiver+ "' and supper like '%"+txtsupper.Text+"%'  and productcode like '%"+txtproductcode.Text+ "%'  order by updatetime desc   ";
            }
            else if (txtreporttype.Text == "HF检测报告")
            {
                sql = "   select supper 供应商,productcode 物料编码,productname 物料名称,productmodel 型号,signdate 签署日期,HFreportcode 报告编码,HFreportdate 报告日期,HFExpiryDate 到期日期,   ";
                sql += "  result_F F,result_Cl Cl,result_Br Br,result_I I,updatetime 更新时间,updateman 记录人,states 状态,AuditingByUser 审核人    ";
                sql += "  from  IQC_SupperHF where  states like '审核%' and SendState='OK'  and Receiver = '"+Receiver+ "' and supper like '%"+txtsupper.Text+ "%'  and productcode like '%"+txtproductcode.Text+ "%'  order by updatetime desc    ";
            }
            DataSet ds = DbAccess.SelectBySql(sql);
            DataTable dt = ds.Tables[0];
            gridView.Columns.Clear();
            gridControl.DataSource = null;
            if (dt == null || dt.Rows.Count < 1)
            {
                MessageBox.Show("没有数据", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            gridControl.DataSource = dt;
        }

        private void btnUnAuditing_Click(object sender, EventArgs e)
        {

            if (ListReceiver.SelectedIndex == -1)
                return;
            string Receiver = ListReceiver.Text.Substring(0, ListReceiver.Text.IndexOf(':'));

            string sql = " ";
            if (txtreporttype.Text == "RoHS测试报告")
            {
                sql = "  select supper 供应商,productcode 物料编码,productname 物料名称,productmodel 型号,signdate 签署日期,rohsreportcode 报告编码,rohsreportdate 报告日期,rohsExpiryDate 到期日期,   ";
                sql += "  result_Cd Cd,result_Pb Pb,result_Hg Hg,result_Cr6 Cr6,result_PBBs PBBs,result_PBDEs PBDEs,result_DBP DBP,result_BBP BBP,result_DEHP DEHP,result_DIBP DIBP,remark 备注信息,updatetime 更新时间,updateman 记录人,states 状态,AuditingByUser 审核人  ";
                sql += "  from IQC_SupperROHS  where states = '未审' and SendState='OK' and  Receiver = '" + Receiver + "' and  supper like '%" + txtsupper.Text + "%'  and productcode like '%" + txtproductcode.Text + "%'  order by updatetime desc   ";
            }
            else if (txtreporttype.Text == "HF检测报告")
            {
                sql = "   select supper 供应商,productcode 物料编码,productname 物料名称,productmodel 型号,signdate 签署日期,HFreportcode 报告编码,HFreportdate 报告日期,HFExpiryDate 到期日期,   ";
                sql += "  result_F F,result_Cl Cl,result_Br Br,result_I I,updatetime 更新时间,updateman 记录人,states 状态,AuditingByUser 审核人    ";
                sql += "  from  IQC_SupperHF where  states = '未审' and SendState='OK'  and  Receiver = '" + Receiver + "' and  supper like '%" + txtsupper.Text + "%'  and productcode like '%" + txtproductcode.Text + "%'  order by updatetime desc    ";
            }
            DataSet ds = DbAccess.SelectBySql(sql);
            DataTable dt = ds.Tables[0];
            gridView.Columns.Clear();
            gridControl.DataSource = null;
            if (dt == null || dt.Rows.Count < 1)
            {
                MessageBox.Show("没有数据", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            gridControl.DataSource = dt;

        }

        private void btnselectall_Click(object sender, EventArgs e)
        {
            string sql = " ",where = " where 1=1 ";
            string supper = txtsupper.Text;
            string productcode = txtproductcode.Text;

            if (txtreporttype.Text == "RoHS测试报告")
            {
                if (!string.IsNullOrEmpty(supper))
                {
                    where += " and r.supper = '" + supper + "' ";
                }
                if (!string.IsNullOrEmpty(productcode))
                {
                    where += " and r.productcode = '" + productcode + "' ";
                }
                sql = "   select r.supper 供应商,r.productcode 物料编码,r.productname 物料名称,r.productmodel 型号,r.signdate 签署日期,rohsreportcode ROHS报告编码,rohsreportdate ROHS报告日期,rohsExpiryDate ROHS报告到期日期,  ";
                sql += "  result_Cd Cd,result_Pb Pb,result_Hg Hg,result_Cr6 Cr6,result_PBBs PBBs,result_PBDEs PBDEs,result_DBP DBP,result_BBP BBP,result_DEHP DEHP,result_DIBP DIBP,  ";
                sql += "   HFreportcode HF报告编码,HFreportdate HF报告日期,HFExpiryDate HF报告到期日期,result_F F,result_Cl Cl,result_Br Br,result_I I, r.remark+isnull(h.remark,'') 备注信息,r.updatetime 更新时间 ";
                sql += "  from IQC_SupperROHS r  left join IQC_SupperHF h on r.supper = h.supper and r.productcode= h.productcode  ";
                sql += where + "  order by r.updatetime desc   ";
            }
            else if (txtreporttype.Text == "HF检测报告")
            {
                if (!string.IsNullOrEmpty(supper))
                {
                    where += " and h.supper = '" + supper + "' ";
                }
                if (!string.IsNullOrEmpty(productcode))
                {
                    where += " and h.productcode = '" + productcode + "' ";
                }
                sql = "   select h.supper 供应商,h.productcode 物料编码,h.productname 物料名称,h.productmodel 型号,h.signdate 签署日期,rohsreportcode ROHS报告编码,rohsreportdate ROHS报告日期,rohsExpiryDate ROHS报告到期日期,  ";
                sql += "  result_Cd Cd,result_Pb Pb,result_Hg Hg,result_Cr6 Cr6,result_PBBs PBBs,result_PBDEs PBDEs,result_DBP DBP,result_BBP BBP,result_DEHP DEHP,result_DIBP DIBP, ";
                sql += "  HFreportcode HF报告编码,HFreportdate HF报告日期,HFExpiryDate HF报告到期日期,result_F F,result_Cl Cl,result_Br Br,result_I I, h.remark+isnull(r.remark,'') 备注信息,h.updatetime 更新时间 ";
                sql += "  from IQC_SupperROHS r  right join IQC_SupperHF h on r.supper = h.supper and r.productcode= h.productcode   ";
                sql += where + "  order by h.updatetime desc   ";
            }
            DataSet ds = DbAccess.SelectBySql(sql);
            DataTable dt = ds.Tables[0];
            gridView.Columns.Clear();
            gridControl.DataSource = null;
            if (dt == null || dt.Rows.Count < 1)
            {
                MessageBox.Show("没有数据", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            gridControl.DataSource = dt;
        }

        void export(DataTable dt)
        {
            int sheetCount = 1;
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = true;
            object missing = System.Reflection.Missing.Value;
            string templetFile = Environment.CurrentDirectory + @"\ReportFolder\供应商环保资料清单.xlsx";
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
            int a = (int)'A';
            for (int i= 0;i<dt.Rows.Count;i++)
            {
                for ( int j = 0;j<dt.Columns.Count;j++ )
                {
                    if (j == 25)
                    {
                        sheet.Cells.get_Range("AC" + (i + 5).ToString()).Value = dt.Rows[i]["备注信息"] == null ? "" : dt.Rows[i]["备注信息"].ToString();
                    }
                    else if (j == 26)
                    {
                        continue;
                    }
                    else
                    {
                        string ss = dt.Rows[i][j] == null ? "" : dt.Rows[i][j].ToString();
                        sheet.Cells.get_Range((char)(a + j) + (i + 5).ToString()).Value = dt.Rows[i][j] == null ? "" : dt.Rows[i][j].ToString();
                    }
                }
            }

        }


        private void Btnreportlist_Click(object sender, EventArgs e)
        {
            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
                return;
            if (gridView.RowCount < 1)
                return;
            if (gridView.FocusedRowHandle < 0)
                return;
            if (gridView.Columns.Count == 27)
            {
                try
                {
                    export(dt);
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(ex.ToString(), "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                     MessageBox.Show("生成完整报表失败", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }
        DataTable rohsreport(string supper,string productcode)
        {
            string sql = "   select supper 供应商,productcode 物料编码,productname 物料名称,productmodel 型号,signdate 签署日期,rohsreportcode 报告编码,rohsreportdate 报告日期,rohsExpiryDate 到期日期 from IQC_SupperROHS   where supper = '"+supper+ "' and productcode = '"+productcode+ "'   order by updatetime desc ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            return dt;
        }
        DataTable hfreport(string supper, string productcode)
        {
            string sql = "    select supper 供应商,productcode 物料编码,productname 物料名称,productmodel 型号,signdate 签署日期,HFreportcode 报告编码,HFreportdate 报告日期,HFExpiryDate 到期日期 from IQC_SupperHF   where supper = '" + supper + "' and productcode = '" + productcode + "'   order by updatetime desc ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            return dt;
        }
        private void Btnsend_Click(object sender, EventArgs e)
        {
            if (ListReceiver.SelectedIndex == -1)
                return;
            if (gridView.RowCount < 1)
                return;
            if (gridView.FocusedRowHandle < 0)
            {
                MessageBox.Show("请选择相应的物料", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            string Receiver = ListReceiver.Text.Substring(0, ListReceiver.Text.IndexOf(':'));
            string mail = ListReceiver.SelectedValue.ToString();
            string supper = gridView.GetFocusedRowCellValue("供应商").ToString();
            string productcode = gridView.GetFocusedRowCellValue("物料编码").ToString();
            string productmodel = gridView.GetFocusedRowCellValue("型号").ToString();
            string sql = " ";
            DataTable dt = null;
            string Directory = "";
            string title = "";
            if (gridView.Columns.Count == 23)
            {
                sql = "  update IQC_SupperROHS set SendState = 'OK' ,SendUser = '"+Login.username+"'  ,SendDate = GETDATE(),Receiver = '"+Receiver+ "'  where supper = '"+supper+ "' and productcode = '"+productcode+ "' ";
                dt = rohsreport(supper, productcode);
                Directory = "共享文件目录为：" + "\\\\QMSSVR\\rohs\\"+supper+"\\"+productcode+"\\"+productmodel;
                title = "ROHS供应商测试报告";
            }
            else if (gridView.Columns.Count == 17)
            {
                sql = "  update IQC_SupperHF set SendState = 'OK' ,SendUser = '"+ Login.username+ "'  ,SendDate = GETDATE(),Receiver = '"+Receiver+ "'  where supper = '"+ supper+ "' and productcode = '"+productcode+ "' ";
                dt = hfreport(supper, productcode);
                Directory = "共享文件目录为：" + "\\\\QMSSVR\\hf\\"+supper+"\\"+productcode+"\\"+productmodel;
                title = "HF测试报告";
            }
            bool flag = DbAccess.ExecuteSql(sql);                   
            string content = ProdTest.HtmlBoy(dt,title, "请注意查收，供应商："+ supper+" ，物料编码："+ productcode+" 需要你审核，" + Directory, "供应商："+ supper+"  ，物料编码："+ productcode);
            if (flag== true)
            {
                try
                {
                    ProdTest.SendHTMLboyMail(mail, "QMS系统：" + DateTime.Now.ToString("yyyyMMdd")+"供应商"+supper+ "物料"+productcode+ "测试报告提醒", content);
                }
                catch
                {
                }
            }
   
        }

        private void gridView_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {
            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
                return;
            if (gridView.RowCount < 1)
                return;
            if (gridView.FocusedRowHandle < 0)
                return;

            txtsupper.Text = gridView.GetFocusedRowCellValue("供应商").ToString();
            txtproductcode.Text = gridView.GetFocusedRowCellValue("物料编码").ToString();
            txtproductname.Text = gridView.GetFocusedRowCellValue("物料名称").ToString();
            txtproductmodel.Text = gridView.GetFocusedRowCellValue("型号").ToString();
            txtRemark.Text = gridView.GetFocusedRowCellValue("备注信息").ToString();

            if (gridView.Columns.Count == 23)
            {
                try
                {
                    txtsigndate.Text = gridView.GetFocusedRowCellValue("签署日期").ToString();
                    txtrohsreportcode.Text  = gridView.GetFocusedRowCellValue("报告编码").ToString();
                    txtrohsreportdate.Text = gridView.GetFocusedRowCellValue("报告日期").ToString();
                    txtrohsExpiryDate.Text = gridView.GetFocusedRowCellValue("到期日期").ToString();
                    txtresult_Cd.Text = gridView.GetFocusedRowCellValue("Cd").ToString();
                    txtresult_Pb.Text = gridView.GetFocusedRowCellValue("Pb").ToString();
                    txtresult_Hg.Text = gridView.GetFocusedRowCellValue("Hg").ToString();
                    txtresult_Cr6.Text = gridView.GetFocusedRowCellValue("Cr6").ToString();
                    txtresult_PBBs.Text = gridView.GetFocusedRowCellValue("PBBs").ToString();
                    txtresult_PBDEs.Text = gridView.GetFocusedRowCellValue("PBDEs").ToString();
                    txtresult_DBP.Text = gridView.GetFocusedRowCellValue("DBP").ToString();
                    txtresult_BBP.Text = gridView.GetFocusedRowCellValue("BBP").ToString();
                    txtresult_DEHP.Text = gridView.GetFocusedRowCellValue("DEHP").ToString();
                    txtresult_DIBP.Text = gridView.GetFocusedRowCellValue("DIBP").ToString();
                }
                catch
                {
                }
            }
            else if (gridView.Columns.Count == 17)
            {
                try
                {
                    txtHFreportcode.Text = gridView.GetFocusedRowCellValue("报告编码").ToString();
                    txtHFreportdate.Text = gridView.GetFocusedRowCellValue("报告日期").ToString();
                    txtHFExpiryDate.Text = gridView.GetFocusedRowCellValue("到期日期").ToString();
                    txtresult_F.Text = gridView.GetFocusedRowCellValue("F").ToString();
                    txtresult_Cl.Text = gridView.GetFocusedRowCellValue("Cl").ToString();
                    txtresult_Br.Text = gridView.GetFocusedRowCellValue("Br").ToString();
                    txtresult_I.Text = gridView.GetFocusedRowCellValue("I").ToString();
                }
                catch
                {

                }
            }
        }
        private void gridView_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {
            DataTable de = gridControl.DataSource as DataTable;
            if (de == null || de.Rows.Count < 1)
                return;
            if (gridView.FocusedRowHandle < 0)
                return;
            string supper = gridView.GetFocusedRowCellValue("供应商").ToString();
            string productcode = gridView.GetFocusedRowCellValue("物料编码").ToString();
            string productmodel = gridView.GetFocusedRowCellValue("型号").ToString();
            try
            {
                if (gridView.Columns.Count == 23)
                {
                    try
                    {
                        if (e.Column.FieldName == "物料编码")
                        {
                            string floerPath = "\\\\QMSSVR\\rohs\\"+supper+"\\"+ productcode+"\\"+productmodel;
                            Cursor.Current = Cursors.WaitCursor;
                            Process.Start(floerPath);
                            Cursor.Current = Cursors.Default;
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                }
                if (gridView.Columns.Count == 17)
                {
                    try
                    {
                        if (e.Column.FieldName == "物料编码")
                        {
                            string floerPath = "\\\\QMSSVR\\hf\\"+supper+"\\"+productcode+"\\"+productmodel;
                            Cursor.Current = Cursors.WaitCursor;
                            Process.Start(floerPath);
                            Cursor.Current = Cursors.Default;
                        }
                    }
                    catch
                    {
                    }
                }
      
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private void gridView_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
        {
            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
            {
                return;
            }
            if (gridView.Columns.Count == 23)
            {
                try
                {
                    if (gridView.GetDataRow(e.RowHandle)["状态"].ToString() == "未审")
                    {
                        e.Appearance.BackColor = Color.Yellow;
                    }
                }
                catch
                {
                }
            }
            if (gridView.Columns.Count == 17)
            {
                try
                {
                    if (gridView.GetDataRow(e.RowHandle)["状态"].ToString() == "未审")
                    {
                        e.Appearance.BackColor = Color.Yellow;
                    }
                }
                catch
                {
                }
            }


        }

        private void ROHScheckallresult_CheckedChanged(object sender, EventArgs e)
        {
            if (ROHScheckallresult.Checked == true)
            {
                txtresult_Cd.Text = "ND";
                txtresult_Pb.Text = "ND";
                txtresult_Hg.Text = "ND";
                txtresult_Cr6.Text = "ND";
                txtresult_PBBs.Text = "ND";
                txtresult_PBDEs.Text = "ND";
                txtresult_DBP.Text = "ND";
                txtresult_BBP.Text = "ND";
                txtresult_DEHP.Text = "ND";
                txtresult_DIBP.Text = "ND";
            }
            else if (ROHScheckallresult.Checked == false)
            {
                txtresult_Cd.Text = "";
                txtresult_Pb.Text = "";
                txtresult_Hg.Text = "";
                txtresult_Cr6.Text = "";
                txtresult_PBBs.Text = "";
                txtresult_PBDEs.Text = "";
                txtresult_DBP.Text = "";
                txtresult_BBP.Text = "";
                txtresult_DEHP.Text = "";
                txtresult_DIBP.Text = "";
            }
        }

        private void HFcheckallresult_CheckedChanged(object sender, EventArgs e)
        {
            if (HFcheckallresult.Checked == true)
            {
                txtresult_F.Text = "ND";
                txtresult_Cl.Text = "ND";
                txtresult_Br.Text = "ND";
                txtresult_I.Text = "ND";
            }
            else if (HFcheckallresult.Checked == false)
            {
                txtresult_F.Text = "";
                txtresult_Cl.Text = "";
                txtresult_Br.Text = "";
                txtresult_I.Text = "";
            }
        }


    }
}