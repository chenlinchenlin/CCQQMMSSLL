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
using System.Data;
using System.Data.SqlClient;
using DX_QMS.Common;
using System.IO;
using System.Diagnostics;
using DevExpress.XtraEditors;

namespace DX_QMS
{
    public partial class OQCCheckAgain : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private string serverFilePath = "192.168.0.204\\FilePath$";
        private string tranid = "", item_id = "", org_id = "";
        private string srecept = "", productcode = "";
        private string checktype = "NGToOK";
        public OQCCheckAgain()
        {
            InitializeComponent();
            bindOrg();
        }


        private void bindorg_id()
        {
            string sql = "select ORG_ID,ORG_NAME from ORG_INFO where ORG_TYPE='ORG' order by SORT ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            txttestorg_id.Properties.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
                txttestorg_id.Properties.Items.Add(row["ORG_NAME"]);
            }
            txtorg_id.SelectedIndex = 0;

        }

        private void bind(string types, ComboBoxEdit com)
        {
            string ssql = "select Definevalue from OQC_TypeDefine where Definetype='" + types + " ' order by sort ";
            DataTable dt = DbAccess.SelectBySql(ssql).Tables[0];
            com.Properties.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
                com.Properties.Items.Add(row["Definevalue"]);
            }
        }

        private void OQCCheckAgain_Load(object sender, EventArgs e)
        {
            bindorg_id();
            bind("客户", txtcustomer);

            //btnOK.Enabled = false;
            //btnbrowse.Enabled = false;
            //btnconfirm.Enabled = false;
        }

        private void txtreason_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13)
                btnOK_Click(sender, null);
        }

        private void txtlotno_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13)
            {
                if (checktype == "NGToOK")
                {
                    string sql = "select sum(sendqty) NGqty from OQC_TestListNew c "
                        + " where testresult='NG' and org_id='" + txtorg_id.SelectedValue.ToString() + "' and workno='" + txtlotno.Text + "'";
                    DataSet ds = DbAccess.SelectBySql(sql);
                    if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                    {
                        txtenableqty.Text = ds.Tables[0].Rows[0]["NGqty"].ToString();
                        txtqty.Focus();
                        txtlotno.Enabled = false;
                    }
                    else
                    {
                        lblinfo.ForeColor = Color.Red;
                        lblinfo.Text = txtlotno.Text + "不存在或没有不良数";
                        txtlotno.Focus();
                        txtlotno.Text = "";
                        txtlotno.Enabled = true;
                    }
                }
                else
                {
                    string sql = "select sum(sendqty) OKqty from OQC_TestListNew c "
                        + " where testresult='OK' and org_id='" + txtorg_id.SelectedValue.ToString() + "' and workno='" + txtlotno.Text + "'";
                    DataSet ds = DbAccess.SelectBySql(sql);
                    if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                    {
                        txtenableqty.Text = ds.Tables[0].Rows[0]["OKqty"].ToString();
                        txtqty.Focus();
                        txtlotno.Enabled = false;
                    }
                    else
                    {
                        lblinfo.ForeColor = Color.Red;
                        lblinfo.Text = txtlotno.Text + "不存在或没有良品数";
                        txtlotno.Focus();
                        txtlotno.Text = "";
                        txtlotno.Enabled = true;
                    }
                }

            }
        }

        private void txtqty_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13)
                txtreason.Focus();
        }
        private void clearform()
        {
            txtchecktype.Enabled = true;
            txtlotno.Enabled = true;
            txtlotno.Text = "";
            txtlotno.Focus();
            txtenableqty.Text = "";
            txtqty.Text = "";
            txtreason.Text = "";
        }



        public string AddOQCTestRecord(string opertype, string item, string customer, string cuscode, string model, string deliveryID, string hytcode
                            , string workno, string serialnumber, string sendqty, string org_id, string sampleqty, string sampleplan, string MA, string MI
                            , string checktype, string allowqty, string lineid, string latyper, string QC, string masters, string QE, string CartonNo, string productionphase
                            , string ECNnumber, string factsampleqty, string NGQty, string NGpoint, string rsno, string productstate, string teststandard, string testremark
                            , string testresult, string testman, string checkman, string Auditman)
        {
            SqlParameter[] para = new SqlParameter[37];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@item", item);
            para[2] = new SqlParameter("@customer", customer);
            para[3] = new SqlParameter("@cuscode", cuscode);
            para[4] = new SqlParameter("@model", model);
            para[5] = new SqlParameter("@deliveryID", deliveryID);
            para[6] = new SqlParameter("@hytcode", hytcode);
            para[7] = new SqlParameter("@workno", workno);
            para[8] = new SqlParameter("@serialnumber", serialnumber);
            para[9] = new SqlParameter("@sendqty", sendqty);
            para[10] = new SqlParameter("@org_id", org_id);
            para[11] = new SqlParameter("@sampleqty", sampleqty);
            para[12] = new SqlParameter("@sampleplan", sampleplan);
            para[13] = new SqlParameter("@MA", MA);
            para[14] = new SqlParameter("@MI", MI);
            para[15] = new SqlParameter("@checktype", checktype);
            para[16] = new SqlParameter("@allowqty", allowqty);
            para[17] = new SqlParameter("@lineid", lineid);
            para[18] = new SqlParameter("@latyper", latyper);
            para[19] = new SqlParameter("@QC", QC);
            para[20] = new SqlParameter("@masters", masters);
            para[21] = new SqlParameter("@QE", QE);
            para[22] = new SqlParameter("@CartonNo", CartonNo);
            para[23] = new SqlParameter("@productionphase", productionphase);
            para[24] = new SqlParameter("@ECNnumber", ECNnumber);
            para[25] = new SqlParameter("@factsampleqty", factsampleqty);
            para[26] = new SqlParameter("@NGQty", NGQty);
            para[27] = new SqlParameter("@NGpoint", NGpoint);
            para[28] = new SqlParameter("@rsno", rsno);
            para[29] = new SqlParameter("@productstate", productstate);
            para[30] = new SqlParameter("@teststandard", teststandard);
            para[31] = new SqlParameter("@testremark", testremark);
            para[32] = new SqlParameter("@testresult", testresult);
            para[33] = new SqlParameter("@testman", testman);
            para[34] = new SqlParameter("@checkman", checkman);
            para[35] = new SqlParameter("@Auditman", Auditman);
            para[36] = new SqlParameter("@msg", SqlDbType.VarChar, 100);
            para[36].Direction = ParameterDirection.Output;
            DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "OQC_AddTestListNew", para);
            return para[36].Value.ToString();
        }
        private void btnOK_Click(object sender, EventArgs e)
        {
            int i = 0;
            if (!int.TryParse(txtqty.Text.Trim(), out i)) return;

            if (int.Parse(txtqty.Text.Trim()) <= int.Parse(txtenableqty.Text))
            {
               IQC ic = new IQC();

                //string barcodemsg = ic.AddOQCTestRecord("checkagain", "", "", "", "", txtorg_id.SelectedValue.ToString(), this.txtlotno.Text, "", "", txtqty.Text
                //                    , "", "", "", "", txtreason.Text, Login.username, "", "", "", checktype, "", "", "", "");

                string barcodemsg = AddOQCTestRecord("checkagain", "", "", "", "", "", ""
                            , this.txtlotno.Text, "", txtqty.Text, txtorg_id.SelectedValue.ToString(), "", "", "", ""
                            , "", "", "", "", "", "", "", "", ""
                            , "", "", "", "", "", checktype, "", txtreason.Text
                            , "", Login.username, "", "");

                if (barcodemsg.IndexOf("成功") >= 0)
                {
                    lblinfo.ForeColor = Color.Blue;
                    lblinfo.Text = barcodemsg;
                    clearform();
                }
                else
                {
                    lblinfo.Text = barcodemsg;
                    lblinfo.ForeColor = Color.Red;
                }
            }
            else
            {
                lblinfo.ForeColor = Color.Red;
                lblinfo.Text = "超过了最大可改判数量:" + this.txtenableqty.Text;
                return;
            }
        }

        private void btnreset_Click(object sender, EventArgs e)
        {
            clearform();
        }
        public static DataSet SAP_GetORG(string type)
        {
            SqlParameter[] para = new SqlParameter[1];
            para[0] = new SqlParameter("@type", type);
            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "SAP_GetOrg", para);
        }
        private void bindOrg()
        {
            this.txtorg_id.DataSource = SAP_GetORG("ORG").Tables[0];
            this.txtorg_id.ValueMember = "ORG_ID";
            this.txtorg_id.DisplayMember = "ORG_NAME";
        }

        private void txtchecktype_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txtchecktype.Text == "由NG改OK")
                checktype = "NGToOK";
            else if (txtchecktype.Text == "由OK改NG")
                checktype = "OKToNG";
            this.txtchecktype.Enabled = false;
            this.txtlotno.Focus();
        }



        private DataTable GetStateByProduct(string pcode)
        {
            string sql = "select States from  OQC_TestMaterialState where Materialcode='" + pcode + "'";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            return dt;
        }

        private void txtproductcode_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13 && txtproductcode.Text != "")
            {
    
                DataTable dt = GetStateByProduct(txtproductcode.Text);
                if (dt.Rows.Count > 0)
                {
                    txtcurrentcheck.Text = dt.Rows[0]["States"].ToString();
                    if (txtcurrentcheck.Text == "暂停检验")
                    {
                        txtmodifycheck.Text = "加严检验";
                        this.btnconfirm.Enabled = true;
                    }
                    else if (txtcurrentcheck.Text == "放宽检验")
                    {
                        txtmodifycheck.Text = "正常检验";
                        this.btnconfirm.Enabled = true;
                    }
                    else
                    {
                        lblinfo2.Text = txtproductcode.Text + "当前检验标准为:" + txtcurrentcheck.Text + ",不能更改检验标准";
                        lblinfo2.ForeColor = Color.Red;
                        txtcurrentcheck.Text = "";
                        txtproductcode.Text = "";
                        txtproductcode.Focus();
                    }
                }
                else
                {
                    lblinfo2.Text = txtproductcode.Text + "编码不存在!";
                    lblinfo2.ForeColor = Color.Red;
                    txtcurrentcheck.Text = "";
                    txtmodifycheck.Text = "";
                    txtproductcode.Text = "";
                    txtproductcode.Focus();
                }

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

        protected void del_prefile(string filepath, string filename)
        {
            if (Directory.Exists(filepath))
            {
                if (Connect(serverFilePath))
                {
                    string[] Mulfile = Directory.GetFiles(filepath);
                    foreach (string ss in Mulfile)
                    {
                        if (System.IO.File.Exists(ss))
                            System.IO.File.Delete(ss);
                    }
                }
            }
        }

        private void btnbrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofdImport = new OpenFileDialog();

            ofdImport.Filter = "文件(*.jpg;bmp;png;jpeg;pdf)|*.jpg;*.bmp;*.png;*.jpeg;*.pdf";
            ofdImport.Multiselect = false;
            DialogResult dr = ofdImport.ShowDialog();
            if (dr == DialogResult.Cancel) return;
            this.txtcertificate.Text = "";
            foreach (string str in ofdImport.FileNames)
            {
                this.txtcertificate.Text += str;
            }
        }
        private bool CopyFileToServer(string folderType, string filePath)
        {
            try
            {
                string[] fileName = @filePath.Split('\\');
                string floerPath = "\\\\" + this.serverFilePath + "\\" + folderType + "\\" + this.txtproductcode.Text.Trim();
                string fileServerPath = floerPath + "\\" + fileName[fileName.Length - 1];
                if (Directory.Exists(floerPath) == false) //如果不存在就创建file文件夹
                {
                    Directory.CreateDirectory(floerPath);
                }
                File.Copy(filePath, fileServerPath, true);
                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }

        private bool UploadFile(string filepath)
        {
            bool b = false;
            if (Connect(serverFilePath))
            {
                bool re = true;
                string msg = "";
                string te = txtcertificate.Text.TrimEnd(',');
                if (te != "")
                {
                    string[] copy = te.Split(',');
                    foreach (string ss in copy)
                    {
                        if (!CopyFileToServer("OQC改判证明", ss))
                        {
                            msg += ss + ";";
                            re = false;
                        }
                    }
                }
                if (re)
                {
                    MessageBox.Show("操作成功");
                    b = true;
                }
                else MessageBox.Show(msg + " 文件上传失败");
            }
            else
            {
                MessageBox.Show("无法连接到服务器的共享目录");
            }
            return b;
        }

        private void btnconfirm_Click(object sender, EventArgs e)
        {
            if (txtcurrentcheck.Text == "")
            {
                MessageBox.Show("请输入PN号","提醒",MessageBoxButtons.OK,MessageBoxIcon.Information);
                return;
            }
            if (txtRemark.Text.Trim() =="")
            {
                MessageBox.Show("请输入备注信息", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (txtcertificate.Text.Trim() == "")
            {
                MessageBox.Show("请添加改判证明", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            DialogResult rt = MessageBox.Show("是否确定要更改检验标准?", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (DialogResult.Yes == rt)
            {
                string fileDBServerPath = "";
                if (txtcertificate.Text != "")
                {
                    string[] fileqty = txtcertificate.Text.TrimEnd(',').Split(',');
                    foreach (string s in fileqty)
                    {
                        string[] fileName = @s.Split('\\');
                        fileDBServerPath += "\\\\" + this.serverFilePath + "\\OQC改判证明" + "\\" + this.txtproductcode.Text.Trim() +"\\" + fileName[fileName.Length - 1] + ",";
                    }
                }
                string floerPath = "";
                floerPath = "\\\\" + this.serverFilePath + "\\OQC改判证明" + "\\" + this.txtproductcode.Text.Trim();
                if (floerPath != "")
                    del_prefile(floerPath, "");

                bool a = UploadFile(fileDBServerPath);

                if (a)
                {
                    string sql = "update OQC_TestMaterialState set SerialnumberOK='',SerialnumberNG='',NGCount=0,OKCount=0,States='" + txtmodifycheck.Text + "',OperUser='" + Login.userId + "',OperDate=getdate() ,Remark='" + txtRemark.Text + "' where Materialcode='" + txtproductcode.Text + "'";
                    bool b = DbAccess.ExecuteSql(sql);
                    if (b)
                    {
                        lblinfo2.Text = txtproductcode.Text + "检验标准更新成功OK!" + txtmodifycheck.Text;
                        lblinfo2.ForeColor = Color.Blue;
                        txtproductcode.Text = "";
                        txtproductcode.Focus();
                        txtcurrentcheck.Text = "";
                        txtmodifycheck.Text = "";
                        txtRemark.Text = "";
                    }
                }

            }
        }

        private void Sbtnreset2_Click(object sender, EventArgs e)
        {
            txtproductcode.Text = "";
            txtcurrentcheck.Text = "";
            txtmodifycheck.Text = "";
            lblinfo2.Text = "";
            txtRemark.Text = "";
            txtcertificate.Text = "";
        }
        private void sBtnselect_Click(object sender, EventArgs e)
        {
            string where = " where 1=1 ";
            string testorg_id = txttestorg_id.Text.Trim();
            string customer = txtcustomer.Text.Trim();
            string workno = txtworkno.Text.Trim();
            if (testorg_id == "" || customer == "" || workno == "")
            {
                MessageBox.Show("请输入组织，客户，工单号","提醒",MessageBoxButtons.OK ,MessageBoxIcon.Information);
                return;
            }
            if (!string.IsNullOrEmpty(testorg_id))
            {
                where += " and org_id = '" + testorg_id + "' ";
            }
            if (!string.IsNullOrEmpty(customer))
            {
                where += " and customer = '" + customer + "' ";
            }
            if (!string.IsNullOrEmpty(workno))
            {
                where += " and workno = '" + workno + "' ";
            }

            string sql = @" select checkdate 检查时间,item 顺序,testresult 检查结果,customer 客户,workno 工单号,sendqty 送检批量,org_id 组织,sampleqty 应抽数量,CartonNo 箱号,productionphase 生产阶段,productstate 产品状态,
                            NGQty NG数量,testman 检验人,testremark 检验备注 from OQC_TestListNew ";
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
        private void sBtnupdate_Click(object sender, EventArgs e)
        {
            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 0)
            {
                MessageBox.Show("没有数据！", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            int i = gridView.FocusedRowHandle;
            if (i < 0)
            {
                MessageBox.Show("请选中要更改的检验项目", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            string testorg_id = gridView.GetFocusedRowCellValue("组织").ToString();
            string customer = gridView.GetFocusedRowCellValue("客户").ToString();
            string workno = gridView.GetFocusedRowCellValue("工单号").ToString();
            string item = gridView.GetFocusedRowCellValue("顺序").ToString();
            string testresult = gridView.GetFocusedRowCellValue("检查结果").ToString();
            string remark  = gridView.GetFocusedRowCellValue("检验备注").ToString();

            if (testresult == "OK")
            {
                testresult = "NG";
            }
            else if (testresult == "NG")
            {
                testresult = "OK";
            }

            string sql = "";
            sql = @" update OQC_TestListNew set testresult = '"+ testresult + "',testremark ='"+ remark+"       "+Login.username+"改判为"+testresult+ "'  where org_id='" + testorg_id + "' and customer='"+ customer + "' and  workno='"+ workno + "' and item='"+ item + "' ";

            bool va = DbAccess.ExecuteSql(sql);

            if (va == true)
            {
                MessageBox.Show("更改成功","提醒",MessageBoxButtons.OK,MessageBoxIcon.Information);
                sBtnselect_Click(sender, e);
            }
            else
            {
                MessageBox.Show("更改失败", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void sBtnreset3_Click(object sender, EventArgs e)
        {
            txttestorg_id.Text = "";
            txtcustomer.Text = "";
            txtworkno.Text = "";
            gridControl.DataSource = null;
        }

        private void gridView_Click(object sender, EventArgs e)
        {
            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 0)
            {
                MessageBox.Show("没有数据！","提醒",MessageBoxButtons.OK,MessageBoxIcon.Information );
            }
            txttestorg_id.Text = gridView.GetFocusedRowCellValue("组织").ToString();
            txtcustomer.Text = gridView.GetFocusedRowCellValue("客户").ToString();
            txtworkno.Text = gridView.GetFocusedRowCellValue("工单号").ToString();       
        }

        private void gridView_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
        {
            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 0)
            {
                return;
            }
            if (gridView.GetDataRow(e.RowHandle)["检查结果"].ToString() == "NG")
            {
                e.Appearance.BackColor = Color.Red;
            }
        }

    }
}