using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using System.Data.SqlClient;
using System.IO;
using System.Windows.Forms;
using DX_QMS.Common;
using DevExpress.XtraBars;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraEditors;
using System.Globalization;

namespace DX_QMS
{
    public partial class overtimecheck : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private string serverFilePath = "192.168.0.204\\FilePath$";
        private Microsoft.Office.Interop.Excel.ApplicationClass appClsExcel = null;
        DataTable dtsubitem = null;
        IQC ic = new IQC();

        string shortnamesup ="", shortnamebrand ="";

        public overtimecheck()
        {
            InitializeComponent();
            BindReceiver("代理证审核");
            setRule();
            gridView.PrintExportProgress += new ProgressChangedEventHandler(gridView_PrintExportProgress);
        }
        private void setRule()
        {
            //Dictionary<string, bool> dic = GroupPermission.SelectRulesForForm(this.Name, Login.groupId);
            //this.btnsave.Enabled = dic["hasInsert"];
            //this.btndel.Enabled = dic["hasDelete"];
            //this.btnOK.Enabled = dic["hasUpdate"];
            //this.btnAudit.Enabled = dic["hasAuditing"];

            string post = "";
            if (Login.manager != "")
            {
                post = Login.manager;
            }
            else
            {
                post = Login.post;
            }
            Dictionary<string, bool> dic = GroupPermission.QMS_SelectRulesForForm(post, "代理证维护");
            this.btnsave.Enabled = dic["hasInsert"];
            this.btndel.Enabled = dic["hasDelete"];
            this.btnOK.Enabled = dic["hasUpdate"];
            this.btnAudit.Enabled = dic["hasAudit"];
        }
        private void BindReceiver(string sType)
        {
            string sql = "select userName,userMail from MailGroup where mailType='" + sType + "'";
            DataSet ds = DbAccess.SelectBySql(sql);
            txtreceiver.DataSource = ds.Tables[0];
            txtreceiver.DisplayMember = "userName";
            txtreceiver.ValueMember = "userMail";
            txtreceiver.SelectedIndex = -1;
        }
        public DateTime beforeTime, afterTime;

        public int IQC_AddProxyCard(string opertype, string productcode, string sup, string brand, string expirydate, string source, string attachment, string userid,string ShortNameSup ,string ShortNameBrand)
        {
            SqlParameter[] para = new SqlParameter[10];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@productcode", productcode);
            para[2] = new SqlParameter("@sup", sup);
            para[3] = new SqlParameter("@brand", brand);
            para[4] = new SqlParameter("@expirydate", expirydate);
            para[5] = new SqlParameter("@source", source);
            para[6] = new SqlParameter("@attachment", attachment);
            para[7] = new SqlParameter("@userid", userid);
            para[8] = new SqlParameter("@ShortNameSup", ShortNameSup);
            para[9] = new SqlParameter("@ShortNameBrand", ShortNameBrand);
            return DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "IQC_ProxyCardOperNew", para);
        }
        public DataSet IQC_ProxyCard(string opertype, string productcode, string sup, string brand, string expirydate, string source, string attachment, string userid)
        {
            SqlParameter[] para = new SqlParameter[8];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@productcode", productcode);
            para[2] = new SqlParameter("@sup", sup);
            para[3] = new SqlParameter("@brand", brand);
            para[4] = new SqlParameter("@expirydate", expirydate);
            para[5] = new SqlParameter("@source", source);
            para[6] = new SqlParameter("@attachment", attachment);
            para[7] = new SqlParameter("@userid", userid);
            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "IQC_ProxyCardOper", para);
        }
        private DataTable BindProxyCardRecord()
        {

            //DataTable dt = IQC_ProxyCard("查询", txtproductcode.Text, this.txtsup.Text, this.txtbrand.Text, this.txtExpiryDate.Value.ToString(), this.txtsource.Text, this.txtpiclist.Text, Login.userId).Tables[0];
            //return dt;

            string sql = @" select  materialsource 进货来源,productcode 产品编码, supplier 供应商, brand 品牌, expirydate 有效日期, attachment 附件,AuditingByUser 审核人,AuditingDate 审核时间,SendDate 发送时间,SendUser 发送人,
			Receiver 邮件接收人,username 操作员,eventtime 操作日期 from IQC_ProxyCardConfig p left join QMS_UserInfo u on p.eventuser=u.userId where productcode='"+ txtproductcode.Text.Trim() + "'";
            DataSet ds = DbAccess.SelectBySql(sql);
            DataTable dt = ds.Tables[0];
            return dt;
           
        }


        private void txtproductcode_Leave(object sender, EventArgs e)
        {
            this.lblcheckstate.Text = "";
            databind.DataSource = null;
            if (txtsource.Text == "") return;
            if (txtproductcode.Text == "") return;
            txtproductcode.Leave -= txtproductcode_Leave;
            //string Orasql = "select organization_id organization,segment1 materialcode,description materialname,en_des materialenname,INVENTORY_ITEM_ID,PRIMARY_UOM_CODE from cux_item_material_v where segment1='" + txtproductcode.Text.Trim() + "'";
            //DataSet ds = barclass.DbAccess.SelectByOracle(Orasql);
            string sql = "select materialname from MaterialSpec where materialcode='" + txtproductcode.Text.Trim() + "'";
            DataSet ds = DbAccess.SelectBySql(sql);
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                lblinfo.Text = ds.Tables[0].Rows[0]["materialname"].ToString();
                lblinfo.ForeColor = Color.Blue;
                txtsup.Focus();
                txtsup.Text = "";
                //绑定已维护过的数据
                if (txtsource.Text == "制造商")
                {
                    DataTable dt = BindProxyCardRecord();
                    databind.DataSource = dt;
                }

            }
            else
            {
                string ssql = "select materialcode from delivery where lotno='" + txtproductcode.Text + "'";
                DataSet dslotno = DbAccess.SelectBySql(ssql);
                if (dslotno != null && ds.Tables.Count > 0 && dslotno.Tables[0].Rows.Count > 0)
                {
                    string materialcode = dslotno.Tables[0].Rows[0]["materialcode"].ToString();
                    //string Orasqlbylotno = "select organization_id organization,segment1 materialcode,description materialname,en_des materialenname,INVENTORY_ITEM_ID,PRIMARY_UOM_CODE from cux_item_material_v where segment1='" + materialcode + "'";
                    //ds = barclass.DbAccess.SelectByOracle(Orasqlbylotno);
                    string Orasqlbylotno = "select materialname from MaterialSpec where materialcode='" + materialcode + "'";
                    ds = DbAccess.SelectBySql(Orasqlbylotno);
                    if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                    {
                        txtproductcode.Text = materialcode;
                        lblinfo.Text = ds.Tables[0].Rows[0]["materialname"].ToString();
                        lblinfo.ForeColor = Color.Blue;
                        txtsup.Focus();
                        txtsup.Text = "";
                        //绑定已维护过的数据
                        if (txtsource.Text == "制造商")
                        {
                            DataTable dt = BindProxyCardRecord();
                            databind.DataSource = dt;
                        }
                    }
                }
                else
                {
                    lblinfo.Text = txtproductcode.Text + "该料号不存在";
                    lblinfo.ForeColor = Color.Red;
                    txtproductcode.Text = "";
                    txtproductcode.Focus();
                }
            }
            txtproductcode.Leave += txtproductcode_Leave;
        }

        private void txtsup_Leave(object sender, EventArgs e)
        {
            txtsup.Leave -= txtsup_Leave;
            txtbrand.Focus();
            txtbrand.Text = "";
            txtsup.Leave += txtsup_Leave;
        }
        private string IfCheck(string sSource, string ssup, string sbrand, string productcode)
        {
            string Flag = "不需审核";
            string sql = "select min(case when materialsource='制造商' then '不需审核' else ISNULL(states,'未审') end) states from  IQC_ProxyCardConfig s where Productcode='" + productcode + "' and s.materialsource='" + sSource + "' and supplier='" + ssup + "' and brand='" + sbrand + "'";
            DataSet ds = DbAccess.SelectBySql(sql);
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                if (ds.Tables[0].Rows[0]["states"].ToString() == "未审")
                    Flag = "未审";
                else if (ds.Tables[0].Rows[0]["states"].ToString() == "不需审核")
                    Flag = "不需审核";
                else if (ds.Tables[0].Rows[0]["states"].ToString() == "NG")
                    Flag = "NG";
                else
                    Flag = "已审核";
            }
            return Flag;
        }
        private void txtbrand_Leave(object sender, EventArgs e)
        {
           ///// txtbrand.Leave -= txtbrand_Leave;
            DataTable dsprog = BindProxyCardRecord();
            if (dsprog != null && dsprog.Rows.Count > 0)
            {
                databind.DataSource = dsprog;

                if (txtbrand.Text == "")
                {
                    return;
                }
                string F = IfCheck(txtsource.Text, txtsup.Text, txtbrand.Text, txtproductcode.Text);
                if (F == "已审核")
                {
                    this.lblcheckstate.Text = "已审核";
                    this.lblcheckstate.ForeColor = Color.Blue;
                    this.btnOK.Enabled = false;
                }
                else if (F == "未审")
                {
                    this.lblcheckstate.Text = "未审核";
                    this.lblcheckstate.ForeColor = Color.Red;
                    this.btnOK.Enabled = true;
                }
                else if (F == "NG")
                {
                    this.lblcheckstate.Text = "NG";
                    this.lblcheckstate.ForeColor = Color.Red;
                    this.btnOK.Enabled = true;
                }
                else
                {
                    this.lblcheckstate.Text = "不需审核";
                    this.lblcheckstate.ForeColor = Color.Green;
                    btnOK.Enabled = false;
                    btnAudit.Enabled = false;
                }
            }
            else
                txtExpiryDate.Focus();
          //////  txtbrand.Leave += txtbrand_Leave;
        }

        private void btnpiclist_Click(object sender, EventArgs e)
        {
            //OpenFileDialog ofdImport = new OpenFileDialog();
            //ofdImport.Filter = "文件(*.jpg;bmp;png;jpeg;pdf)|*.jpg;*.bmp;*.png;*.jpeg;*.pdf";
            //ofdImport.Multiselect = false;
            //DialogResult dr = ofdImport.ShowDialog();
            //if (dr == DialogResult.Cancel) return;
            //this.txtpiclist.Text = "";
            //foreach (string str in ofdImport.FileNames)
            //{
            //    this.txtpiclist.Text += str;
            //}

            if (txtsup.Text.Trim() == "" && txtbrand.Text.Trim()== "")
            {
                MessageBox.Show("请输入供应商和品牌信息","提醒",MessageBoxButtons.OK,MessageBoxIcon.Information);
                return;
            }
            databind.DataSource = null;

            string where = " where 1=1 ";
            string supper = txtsup.Text;
            string brand = txtbrand.Text;
            if (!string.IsNullOrEmpty(supper))
            {
                where += " and  CHARINDEX(supplier,'"+ txtsup.Text+ "') > 0   ";    ///// CHARINDEX(supplier,'chedf东莞宏致')>0
            }
            if (!string.IsNullOrEmpty(brand))
            {
               ///// where += " and brand like '%" + brand + "%' ";

                where += " and  CHARINDEX(brand,'"+ txtbrand.Text + "') > 0  ";

            }

            string sql = @"  select supplier 供应商,brand 品牌,materialsource 进货来源,expirydate 有效日期,filename 附件,remark 备注,updateman 更新人,CONVERT(varchar(100),updatetime, 20) 更新时间 from IQC_ProxyCardLibrary  ";
            sql += where + " order by updatetime desc ";

            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];

            gridView.Columns.Clear();
          
            databind.DataSource = dt;

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

        private bool CopyFileToServer(string folderType, string filePath)
        {
            try
            {
                string[] fileName = @filePath.Split('\\');
                string floerPath = "\\\\" + this.serverFilePath + "\\" + folderType + "\\" + this.txtproductcode.Text.Trim() + "\\" + txtsup.Text + "\\" + txtbrand.Text;
                string fileServerPath = floerPath + "\\" + fileName[fileName.Length - 1];
                if (Directory.Exists(floerPath) == false)//如果不存在就创建file文件夹
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
                string te = txtpiclist.Text.TrimEnd(',');
                if (te != "")
                {
                    string[] copy = te.Split(',');
                    foreach (string ss in copy)
                    {
                        if (!CopyFileToServer("代理证", ss))
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

        private void btnsave_Click(object sender, EventArgs e)
        {
            if (txtsource.Text == "")
                return;
            if (txtproductcode.Text == "")
                return;
            //20150418新增加审核控制
            if (lblcheckstate.Text == "已审核" && btnAudit.Enabled == false)
            {
                MessageBox.Show(txtproductcode.Text + ",该物料已经审核,你没权再修改", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            if (txtsource.Text != "制造商")
            {
                if (this.txtproductcode.Text == "" || txtsup.Text.Trim() == "" || this.txtbrand.Text.Trim() == "" || this.txtpiclist.Text == "")
                {
                    lblinfo.Text = "供应商,品牌,代理证,必填内容不能为空!";
                    lblinfo.ForeColor = Color.Red;
                    return;
                }
            }


            //DateTime expiry = Convert.ToDateTime(txtExpiryDate.Text);
            //string ExpiryDate = expiry.ToString("yyyy-MM-dd");

            //string fileDBServerPath = "";
            //if (txtpiclist.Text != "")
            //{
            //    string[] fileqty = txtpiclist.Text.TrimEnd(',').Split(',');
            //    foreach (string s in fileqty)
            //    {
            //        string[] fileName = @s.Split('\\');
            //        fileDBServerPath += "\\\\" + this.serverFilePath + "\\代理证" + "\\" + this.txtproductcode.Text.Trim() + "\\" + txtsup.Text.Trim() + "\\" + txtbrand.Text.Trim() + "\\" + fileName[fileName.Length - 1] + ",";
            //    }
            //}         
            //string floerPath = "";
            //floerPath = "\\\\" + this.serverFilePath + "\\代理证" + "\\" + this.txtproductcode.Text.Trim() + "\\" + txtsup.Text.Trim() + "\\" + txtbrand.Text.Trim();
            //if (floerPath != "")
            //    del_prefile(floerPath, txtsup.Text);

            //bool b = UploadFile(fileDBServerPath);
            //if (b)
            //{
            int i = IQC_AddProxyCard("新增", txtproductcode.Text.Trim(), this.txtsup.Text.Trim(), this.txtbrand.Text.Trim(), this.txtExpiryDate.Text, this.txtsource.Text, this.txtpiclist.Text, Login.userId, shortnamesup, shortnamebrand);
                if (i > 0)
                {
                    lblinfo.Text = "来源:" + txtsource.Text + ",编码:" + txtproductcode.Text + ",OK";
                    lblinfo.ForeColor = Color.Blue;

                    txtsup.Text = "";
                    txtbrand.Text = "";
                    txtpiclist.Text = "";
                    shortnamesup = "";
                    shortnamebrand = "";

                }
           // }
            databind.DataSource = BindProxyCardRecord();
        }

        private void btndel_Click(object sender, EventArgs e)
        {
            DataTable de = databind.DataSource as DataTable;
            if (de == null || de.Rows.Count < 1)
                return;
            //if (gridView.FocusedRowHandle < 0)
            //    return;
            if (gridView.GetSelectedRows().Length < 1)
                return;
            //20150418新增加审核控制

            if (gridView.Columns.Count == 8)
                return;

            if (lblcheckstate.Text == "已审核" && btnAudit.Enabled == false)
            {
                MessageBox.Show(txtproductcode.Text + ",该物料已经审核,你没权再删除", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            for (int k = gridView.GetSelectedRows().Length; k > 0; k--)
            {
                DataRow dr = gridView.GetDataRow(gridView.GetSelectedRows()[k - 1]);
                string floerPath = "";
                floerPath = "\\\\" + this.serverFilePath + "\\代理证" + "\\" + dr["产品编码"].ToString() + "\\" + dr["供应商"].ToString() + "\\" + dr["品牌"].ToString();
                if (floerPath != "")
                    del_prefile(floerPath, txtsup.Text);

                int i = IQC_AddProxyCard("删除", dr["产品编码"].ToString(), dr["供应商"].ToString(), dr["品牌"].ToString(), "", dr["进货来源"].ToString(), "", "","","");
            }
            databind.DataSource = BindProxyCardRecord();
        }

        private void btnsearch_Click(object sender, EventArgs e)
        {

            gridView.Columns.Clear();
            databind.DataSource = null;
            DataTable dt = BindProxyCardRecord();
            if (dt == null || dt.Rows.Count < 1)
            {
                MessageBox.Show("没有数据","提醒",MessageBoxButtons.OK,MessageBoxIcon.Information );
                return;
            }
            databind.DataSource = dt;
            gridView.RefreshData();
            // txtproductcode.Focus();
        }

        private void btnreset_Click(object sender, EventArgs e)
        {
            txtproductcode.Text = "";
            txtsup.Text = "";
            txtbrand.Text = "";
            txtpiclist.Text = "";
            shortnamesup = "";
            shortnamebrand = "";
            this.btnsave.Enabled = true;
            this.btndel.Enabled = true;
            this.btnOK.Enabled = true;
            this.btnAudit.Enabled = true;
            txtremarks.Text = "";
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            string users = "";
            for (int i = 0; i < txtreceiver.SelectedItems.Count; i++)
            {
                users = users + "," + txtreceiver.GetItemText(txtreceiver.SelectedItems[i]);
            }
            users = users.TrimStart(',').TrimEnd(',');

            if (users == "")
            {
                MessageBox.Show("请选择收件人", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            string sql = "update IQC_ProxyCardConfig set SendState='OK',states='未审',SendUser='" + Login.username + "',Receiver='" + users + "',SendDate=getdate() where Productcode='" + txtproductcode.Text + "' and supplier='" + txtsup.Text.Trim() + "' and brand='" + txtbrand.Text.Trim() + "'";
            if (DbAccess.ExecuteSql(sql))
            {
                string subject = "【重要信息】" + txtsource.Text + ",物料编码:" + txtproductcode.Text + "供应商:" + txtsup.Text + ",品牌:" + txtbrand.Text + "代理证,需要您审核处理";
                string body = "QMS系统提醒您，进货来源：" + txtsource.Text + ",物料编码:" + txtproductcode.Text + txtsup.Text + ",品牌:" + txtbrand.Text + "代理证，需要您审核，谢谢!";
                //ProdTest.SendMailToSQE(Login.deptId, subject, body, "代理证审核", users);
                ProdTest.SendMailToSQE("", subject, body, "代理证审核", users);
            }
           
        }
        private void Message()
        {
            string msg = "进货来源:" + txtsource.Text + ",编码：" + txtproductcode.Text + ",供应商:" + txtsup.Text + "品牌:" + txtbrand.Text + ",没有上传代理证,不能审核!";
            MessageBox.Show(msg, "提示", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            lblinfo.Text = msg;
            lblinfo.ForeColor = Color.Red;
        }
        private void btnAudit_Click(object sender, EventArgs e)
        {

            DataTable dt = databind.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
                return;
            if (gridView.RowCount < 1)
                return;
            if (gridView.FocusedRowHandle < 0)
                return;


            string state = "";
            if (!cbok.Checked && !cbNG.Checked)
            {
                MessageBox.Show("请选择一个审核结果", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (cbNG.Checked && txtremarks.Text == "")
            {
                MessageBox.Show("请输入NG备注", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (txtsource.Text != "制造商")
            {
                string source = txtsource.Text;
                string pno = txtproductcode.Text.Trim();
                string sup = txtsup.Text.Trim();
                string brand = txtbrand.Text;
                string selectproxysql = @"  select ShortNameSup,ShortNameBrand from IQC_ProxyCardConfig where productcode = '"+ pno + "' and supplier = '"+ sup + "' and brand = '"+ brand + "' ";
                DataTable selectproxydt = DbAccess.SelectBySql(selectproxysql).Tables[0];
                string ShortNameSup = selectproxydt.Rows[0]["ShortNameSup"].ToString();
                string ShortNameBrand = selectproxydt.Rows[0]["ShortNameBrand"].ToString();

                string floerPath = "\\\\" + this.serverFilePath + "\\代理证归档库" + "\\" + ShortNameSup + "\\" + ShortNameBrand + "\\" + source;

                //// string floerPath = "\\\\" + this.serverFilePath + "\\代理证" + "\\" + pno + "\\" + sup + "\\" + brand;

                if (Connect(serverFilePath))
                {
                    if (!Directory.Exists(floerPath))  
                    {
                        Message();
                        return;
                    }
                    else
                    {
                        string[] pt = Directory.GetFiles(floerPath);
                        if (pt.Length == 0)
                        {
                            Message();
                            return;
                        }
                    }
                }

            }

            if (cbok.Checked)
            {
                state = "已审核";
            }
            else if (cbNG.Checked)
            {
                state = "NG";
            }

            string sql = "update IQC_ProxyCardConfig set states= '" + state + "',RemarksSQE='" + txtremarks.Text + "',AuditingByUser='" + Login.username + "',AuditingDate=getdate() where Productcode='" + txtproductcode.Text + "' and supplier='" + txtsup.Text.Trim() + "' and brand='" + txtbrand.Text.Trim() + "'";
            if (DbAccess.ExecuteSql(sql))
                MessageBox.Show(txtproductcode.Text + ",审核成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

            txtremarks.Text = "";
        }

        private void btnAuditQuery_Click(object sender, EventArgs e)
        {
            string sql = "select  materialsource 进货来源,productcode 产品编码, supplier 供应商, brand 品牌, expirydate 有效日期, attachment 附件,AuditingByUser 审核人,AuditingDate 审核时间,SendDate 发送时间,SendUser 发送人,";
            sql += " Receiver 邮件接收人,username 操作员,eventtime 操作日期 from IQC_ProxyCardConfig s left join QMS_UserInfo u on s.eventuser=u.userId ";
            sql += " where isnull(states,'未审')='已审核' and SendState='OK'  and isnull(Receiver," + "'" + txtreceiver.Text + "')= case " + "'" + txtreceiver.Text + "' when '' then isnull(Receiver," + "'" + txtreceiver.Text + "') else '" + txtreceiver.Text + "' end and Productcode like " + "'" + txtproductcode.Text + "%'";
            DataSet ds = DbAccess.SelectBySql(sql);
            DataTable dt = ds.Tables[0];

            gridView.Columns.Clear();
            databind.DataSource = null;

            if (dt == null || dt.Rows.Count < 1)
            {
                MessageBox.Show("没有数据","提醒",MessageBoxButtons.OK,MessageBoxIcon.Information);
                return;
            }
            databind.DataSource = dt;
        }

        private string ShowSaveFileDialog(string title, string filter)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            string name = "IQC代理证维护信息";
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

        private void btnPrint_Click(object sender, EventArgs e)
        {
            //if (databind.Rows.Count <= 0) return;
            //Solder so = new Solder();
            //so.DataToExcel(databind);
            DataTable dt = databind.DataSource as DataTable;
            if (dt == null)
                return;
            if (dt.Rows.Count <= 0) return;
            string fileName = ShowSaveFileDialog("Microsoft Excel 2007 Document", "Microsoft Excel|*.xlsx");
            if (fileName == string.Empty) return;
            ExportToEx(fileName, "xlsx", gridView);
            OpenFile(fileName);

        }

        private void btnUnAuditing_Click(object sender, EventArgs e)
        {
            string sql = "select  materialsource 进货来源,productcode 产品编码, supplier 供应商, brand 品牌, expirydate 有效日期, attachment 附件,AuditingByUser 审核人,AuditingDate 审核时间,SendDate 发送时间,SendUser 发送人,";
            sql += " Receiver 邮件接收人,username 操作员,eventtime 操作日期 from IQC_ProxyCardConfig s left join QMS_UserInfo u on s.eventuser=u.userId ";
            sql += " where isnull(states,'未审')='未审' and SendState='OK'  and isnull(Receiver," + "'" + txtreceiver.Text + "')= case " + "'" + txtreceiver.Text + "' when '' then isnull(Receiver," + "'" + txtreceiver.Text + "') else '" + txtreceiver.Text + "' end and Productcode like " + "'" + txtproductcode.Text + "%'";
            DataSet ds = DbAccess.SelectBySql(sql);
            DataTable dt = ds.Tables[0];

            gridView.Columns.Clear();
            databind.DataSource = null;

            if (dt == null || dt.Rows.Count < 1)
            {
                MessageBox.Show("没有数据","提醒",MessageBoxButtons.OK,MessageBoxIcon.Information);
                return;
            }
            databind.DataSource = dt;

        }

        private void cbok_CheckedChanged(object sender, EventArgs e)
        {
            if (cbok.Checked)
            {
                cbok.Text = "OK";
                txtremarks.Enabled = false;
                txtremarks.Text = "";
                cbNG.Checked = false;
            }
        }

        private void cbNG_CheckedChanged(object sender, EventArgs e)
        {
            if (cbNG.Checked)
            {
                cbNG.Text = "NG";
                this.txtremarks.Enabled = true;
                cbok.Checked = false;
            }
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
        private void gridView_Click(object sender, EventArgs e)
        {
            ////  DataGridView dgv = (DataGridView)sender;
            //如果是"instock"列，被点击   gridView.GetFocusedRowCellValue("品牌").ToString();
            //////try
            //////{   /////e.Column.FieldName == "附件"
            //////    if (e.Column.FieldName == "附件" && txtsource.Text != "制造商")
            //////    {
            //////        string pno = gridView.GetFocusedRowCellValue("产品编码").ToString().Trim();
            //////        string sup = gridView.GetFocusedRowCellValue("供应商").ToString().Trim();
            //////        string brand = gridView.GetFocusedRowCellValue("品牌").ToString().Trim();
            //////        string floerPath = "\\\\" + this.serverFilePath + "\\代理证" + "\\" + pno + "\\" + sup + "\\" + brand;

            //////        if (Connect(serverFilePath))
            //////        {
            //////            if (Directory.Exists(floerPath))
            //////            {
            //////                string[] pt = Directory.GetFiles(floerPath);
            //////                if (pt.Length == 0) return;
            //////                if (pt.Length > 0)
            //////                {
            //////                    for (int i = 0; i < pt.Length; i++)
            //////                    {
            //////                        if (pt[i].ToString().Contains(".db")) continue;
            //////                        //string filename = Path.GetFileNameWithoutExtension(pt[i].ToString());
            //////                        string filename = Path.GetFileName(pt[i].ToString());

            //////                        FileStream fs = new FileStream(pt[i].ToString(), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            //////                        //FileStream fs = new FileStream(pt[i].ToString(), FileMode.Open);
            //////                        BinaryReader reader = new BinaryReader(fs);
            //////                        string fileclass = "";
            //////                        try
            //////                        {
            //////                            for (int j = 0; j < 2; j++)
            //////                            {
            //////                                fileclass += reader.ReadByte().ToString();
            //////                            }

            //////                        }
            //////                        catch (Exception)
            //////                        {

            //////                            throw;
            //////                        }

            //////                        //if (fileclass == "3780")
            //////                        {
            //////                            OpenFile(floerPath, filename);
            //////                            fs.Close();
            //////                            fs.Dispose();
            //////                            reader.Close();
            //////                            return;
            //////                        }
            //////                    }
            //////                }
            //////            }
            //////            else
            //////            {
            //////                lblinfo.Text = "没有上传代理证";
            //////                lblinfo.ForeColor = Color.Red;
            //////            }
            //////        }
            //////        return;
            //////    }

            //////    try
            //////    {
            //////        ////gridView.GetFocusedRowCellValue("有效日期").ToString()

            //////        txtproductcode.Text = gridView.GetFocusedRowCellValue("产品编码").ToString();
            //////        this.txtsource.Text = gridView.GetFocusedRowCellValue("进货来源").ToString();
            //////        this.txtsup.Text = gridView.GetFocusedRowCellValue("供应商").ToString();
            //////        this.txtbrand.Text = gridView.GetFocusedRowCellValue("品牌").ToString();
            //////        this.txtExpiryDate.Text = gridView.GetFocusedRowCellValue("有效日期").ToString();
            //////        string F = IfCheck(txtsource.Text, txtsup.Text, txtbrand.Text, txtproductcode.Text);
            //////        if (F == "已审核")
            //////        {
            //////            this.lblcheckstate.Text = "已审核";
            //////            this.lblcheckstate.ForeColor = Color.Blue;
            //////            this.btnOK.Enabled = false;
            //////        }
            //////        else if (F == "未审")
            //////        {
            //////            this.lblcheckstate.Text = "未审核";
            //////            this.lblcheckstate.ForeColor = Color.Red;
            //////            this.btnOK.Enabled = true;
            //////        }
            //////        else if (F == "NG")
            //////        {
            //////            this.lblcheckstate.Text = "NG";
            //////            this.lblcheckstate.ForeColor = Color.Red;
            //////            this.btnOK.Enabled = true;
            //////        }
            //////        else
            //////        {
            //////            this.lblcheckstate.Text = "不需审核";
            //////            this.lblcheckstate.ForeColor = Color.Green;
            //////            btnOK.Enabled = false;
            //////            btnAudit.Enabled = false;
            //////        }
            //////    }
            //////    catch { }
            //////}
            //////catch { }
        }

        private void gridView_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
        {

            DataTable dt = databind.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
            {
                return;
            }

            if (gridView.Columns.Count == 13)
            {
                try
                {
                    if (gridView.GetDataRow(e.RowHandle)["states"].ToString() == "NG")
                    {
                        e.Appearance.BackColor = Color.Yellow;
                    }
                    if (gridView.GetDataRow(e.RowHandle)["states"].ToString() == "已审核")
                    {
                        e.Appearance.BackColor = Color.Gold;
                    }
                }
                catch
                {
                }
            }
        }

        private void gridView_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {
            DataTable de = databind.DataSource as DataTable;
            if (de == null || de.Rows.Count < 1)
                return;
            if (gridView.FocusedRowHandle < 0)
                return;
            

            try
            {

                if (gridView.Columns.Count == 13)
                {
                    try
                    {

                        if (e.Column.FieldName == "附件" && txtsource.Text != "制造商")
                        {
                            string pno = gridView.GetFocusedRowCellValue("产品编码").ToString().Trim();
                            string sup = gridView.GetFocusedRowCellValue("供应商").ToString().Trim();
                            string brand = gridView.GetFocusedRowCellValue("品牌").ToString().Trim();


                            //string source = txtsource.Text;
                            //string pno = txtproductcode.Text.Trim();
                            //string sup = txtsup.Text.Trim();
                            //string brand = txtbrand.Text;
                            //string selectproxysql = @"  select ShortNameSup,ShortNameBrand from IQC_ProxyCardConfig where productcode = '" + pno + "' and supplier = '" + sup + "' and brand = '" + brand + "' ";
                            //DataTable selectproxydt = DbAccess.SelectBySql(selectproxysql).Tables[0];
                            //string ShortNameSup = selectproxydt.Rows[0]["ShortNameSup"].ToString();
                            //string ShortNameBrand = selectproxydt.Rows[0]["ShortNameBrand"].ToString();
                            //string floerPath = "\\\\" + this.serverFilePath + "\\代理证归档库" + "\\" + ShortNameSup + "\\" + ShortNameBrand + "\\" + source;


                            string floerPath = "\\\\" + this.serverFilePath + "\\代理证" + "\\" + pno + "\\" + sup + "\\" + brand;

                            if (Connect(serverFilePath))
                            {
                                if (Directory.Exists(floerPath))
                                {
                                    string[] pt = Directory.GetFiles(floerPath);
                                    if (pt.Length == 0) return;
                                    if (pt.Length > 0)
                                    {
                                        for (int i = 0; i < pt.Length; i++)
                                        {
                                            if (pt[i].ToString().Contains(".db")) continue;
                                            //string filename = Path.GetFileNameWithoutExtension(pt[i].ToString());
                                            string filename = Path.GetFileName(pt[i].ToString());

                                            FileStream fs = new FileStream(pt[i].ToString(), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                                            //FileStream fs = new FileStream(pt[i].ToString(), FileMode.Open);
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

                                            //if (fileclass == "3780")
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
                                else
                                {

                                    string selectproxysql = @"  select ShortNameSup,ShortNameBrand,materialsource  from IQC_ProxyCardConfig where productcode = '" + pno + "' and supplier = '" + sup + "' and brand = '" + brand + "' ";
                                    DataTable selectproxydt = DbAccess.SelectBySql(selectproxysql).Tables[0];
                                    string ShortNameSup = selectproxydt.Rows[0]["ShortNameSup"].ToString();
                                    string ShortNameBrand = selectproxydt.Rows[0]["ShortNameBrand"].ToString();
                                    string source = selectproxydt.Rows[0]["materialsource"].ToString();
                                    string openfloerPath = "\\\\" + this.serverFilePath + "\\代理证归档库" + "\\" + ShortNameSup + "\\" + ShortNameBrand + "\\" + source;

                                    if (Directory.Exists(openfloerPath))
                                    {
                                        string[] pt = Directory.GetFiles(openfloerPath);
                                        if (pt.Length == 0) return;
                                        if (pt.Length > 0)
                                        {
                                            for (int i = 0; i < pt.Length; i++)
                                            {
                                                if (pt[i].ToString().Contains(".db")) continue;
                                                //string filename = Path.GetFileNameWithoutExtension(pt[i].ToString());
                                                string filename = Path.GetFileName(pt[i].ToString());

                                                FileStream fs = new FileStream(pt[i].ToString(), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                                                //FileStream fs = new FileStream(pt[i].ToString(), FileMode.Open);
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

                                                //if (fileclass == "3780")
                                                {
                                                    OpenFile(openfloerPath, filename);
                                                    fs.Close();
                                                    fs.Dispose();
                                                    reader.Close();
                                                    return;
                                                }
                                            }
                                        }

                                    }
                                    else
                                    {
                                        lblinfo.Text = "没有上传代理证";
                                        lblinfo.ForeColor = Color.Red;
                                    }
                                }
                            }
                            return;
                        }
                    }
                    catch
                    {

                    }
                }

                if (gridView.Columns.Count == 8)
                {
                    try
                    {
                        string supper = gridView.GetFocusedRowCellValue("供应商").ToString();
                        string brand = gridView.GetFocusedRowCellValue("品牌").ToString();
                        string source = gridView.GetFocusedRowCellValue("进货来源").ToString();
                        string filename = gridView.GetFocusedRowCellValue("附件").ToString();

                        if (e.Column.FieldName == "附件")
                        {
                            string openfilepath = "\\\\192.168.0.204\\FilePath$\\代理证归档库" + "\\" + supper + "\\" + brand + "\\" + source;
                            string openfilename = Path.GetFileName(openfilepath + "\\" + filename);
                            FileStream fs = new FileStream(openfilepath + "\\" + filename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
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
                                OpenFile(openfilepath, openfilename);
                                fs.Close();
                                fs.Dispose();
                                reader.Close();
                                return;
                            }
                        }

                    }
                    catch
                    {

                    }
                }
                //try
                //{
                //    ////gridView.GetFocusedRowCellValue("有效日期").ToString()

                //    txtproductcode.Text = gridView.GetFocusedRowCellValue("产品编码").ToString();
                //    this.txtsource.Text = gridView.GetFocusedRowCellValue("进货来源").ToString();
                //    this.txtsup.Text = gridView.GetFocusedRowCellValue("供应商").ToString();
                //    this.txtbrand.Text = gridView.GetFocusedRowCellValue("品牌").ToString();
                //    this.txtExpiryDate.Text = gridView.GetFocusedRowCellValue("有效日期").ToString();
                //    string F = IfCheck(txtsource.Text, txtsup.Text, txtbrand.Text, txtproductcode.Text);
                //    if (F == "已审核")
                //    {
                //        this.lblcheckstate.Text = "已审核";
                //        this.lblcheckstate.ForeColor = Color.Blue;
                //        this.btnOK.Enabled = false;
                //    }
                //    else if (F == "未审")
                //    {
                //        this.lblcheckstate.Text = "未审核";
                //        this.lblcheckstate.ForeColor = Color.Red;
                //        this.btnOK.Enabled = true;
                //    }
                //    else if (F == "NG")
                //    {
                //        this.lblcheckstate.Text = "NG";
                //        this.lblcheckstate.ForeColor = Color.Red;
                //        this.btnOK.Enabled = true;
                //    }
                //    else
                //    {
                //        this.lblcheckstate.Text = "不需审核";
                //        this.lblcheckstate.ForeColor = Color.Green;
                //        btnOK.Enabled = false;
                //        btnAudit.Enabled = false;
                //    }
                //}
                //catch { }
            }
            catch { }
        }

        private void txtproductcode_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13 && txtproductcode.Text != "")
            {
                this.lblcheckstate.Text = "";
                databind.DataSource = null;
               ///// if (txtsource.Text == "") return;
                if (txtproductcode.Text == "") return;
             
                //string Orasql = "select organization_id organization,segment1 materialcode,description materialname,en_des materialenname,INVENTORY_ITEM_ID,PRIMARY_UOM_CODE from cux_item_material_v where segment1='" + txtproductcode.Text.Trim() + "'";
                //DataSet ds = barclass.DbAccess.SelectByOracle(Orasql);
                string sql = "select materialname from MaterialSpec where materialcode='" + txtproductcode.Text.Trim() + "'";
                DataSet ds = DbAccess.SelectBySql(sql);
                if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                {
                    lblinfo.Text = ds.Tables[0].Rows[0]["materialname"].ToString();
                    lblinfo.ForeColor = Color.Blue;
                    txtsup.Focus();
                    txtsup.Text = "";
                    //绑定已维护过的数据
                    if (txtsource.Text == "制造商")
                    {
                        DataTable dt = BindProxyCardRecord();
                        databind.DataSource = dt;
                    }

                }
                else
                {
                    string ssql = "select materialcode from delivery where lotno='" + txtproductcode.Text + "'";
                    DataSet dslotno = DbAccess.SelectBySql(ssql);
                    if (dslotno != null && ds.Tables.Count > 0 && dslotno.Tables[0].Rows.Count > 0)
                    {
                        string materialcode = dslotno.Tables[0].Rows[0]["materialcode"].ToString();
                        //string Orasqlbylotno = "select organization_id organization,segment1 materialcode,description materialname,en_des materialenname,INVENTORY_ITEM_ID,PRIMARY_UOM_CODE from cux_item_material_v where segment1='" + materialcode + "'";
                        //ds = barclass.DbAccess.SelectByOracle(Orasqlbylotno);
                        string Orasqlbylotno = "select materialname from MaterialSpec where materialcode='" + materialcode + "'";
                        ds = DbAccess.SelectBySql(Orasqlbylotno);
                        if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                        {
                            txtproductcode.Text = materialcode;
                            lblinfo.Text = ds.Tables[0].Rows[0]["materialname"].ToString();
                            lblinfo.ForeColor = Color.Blue;
                            txtsup.Focus();
                            txtsup.Text = "";
                            //绑定已维护过的数据
                            if (txtsource.Text == "制造商")
                            {
                                DataTable dt = BindProxyCardRecord();
                                databind.DataSource = dt;
                            }
                        }
                    }
                    else
                    {
                        lblinfo.Text = txtproductcode.Text + "该料号不存在";
                        lblinfo.ForeColor = Color.Red;
                        txtproductcode.Text = "";
                        txtproductcode.Focus();
                    }
                }
              
            }
        }

        private void txtbrand_KeyUp(object sender, KeyEventArgs e)
        {

            if (e.KeyValue == 13 && txtbrand.Text != "")
            {
                txtbrand_Leave(sender,e);
            }
         }

        private void gridView_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {
            DataTable dt = databind.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
                return;
            if (gridView.RowCount < 1)
                return;
            if (gridView.FocusedRowHandle < 0)
                return;
            if (gridView.Columns.Count == 8)
            {
                try
                {
                    txtsource.Text = gridView.GetFocusedRowCellValue("进货来源").ToString().Trim();
                    txtExpiryDate.Text = gridView.GetFocusedRowCellValue("有效日期").ToString().Trim();
                    txtpiclist.Text = gridView.GetFocusedRowCellValue("附件").ToString().Trim();

                    ////txtsup.Text = gridView.GetFocusedRowCellValue("供应商").ToString();
                    ////txtbrand.Text = gridView.GetFocusedRowCellValue("品牌").ToString();

                    shortnamesup = gridView.GetFocusedRowCellValue("供应商").ToString();
                    shortnamebrand = gridView.GetFocusedRowCellValue("品牌").ToString();
                }
                catch
                {
                }
            }

            if (gridView.Columns.Count == 13)
            {
                try
                {
                    ////gridView.GetFocusedRowCellValue("有效日期").ToString()

                    txtproductcode.Text = gridView.GetFocusedRowCellValue("产品编码").ToString();
                    this.txtsource.Text = gridView.GetFocusedRowCellValue("进货来源").ToString();
                    this.txtsup.Text = gridView.GetFocusedRowCellValue("供应商").ToString();
                    this.txtbrand.Text = gridView.GetFocusedRowCellValue("品牌").ToString();
                    this.txtExpiryDate.Text = gridView.GetFocusedRowCellValue("有效日期").ToString();
                    string F = IfCheck(txtsource.Text, txtsup.Text, txtbrand.Text, txtproductcode.Text);
                    if (F == "已审核")
                    {
                        this.lblcheckstate.Text = "已审核";
                        this.lblcheckstate.ForeColor = Color.Blue;
                        this.btnOK.Enabled = false;
                    }
                    else if (F == "未审")
                    {
                        this.lblcheckstate.Text = "未审核";
                        this.lblcheckstate.ForeColor = Color.Red;
                        this.btnOK.Enabled = true;
                    }
                    else if (F == "NG")
                    {
                        this.lblcheckstate.Text = "NG";
                        this.lblcheckstate.ForeColor = Color.Red;
                        this.btnOK.Enabled = true;
                    }
                    else
                    {
                        this.lblcheckstate.Text = "不需审核";
                        this.lblcheckstate.ForeColor = Color.Green;
                        btnOK.Enabled = false;
                        btnAudit.Enabled = false;
                    }
                }
                catch
                { }

            }


        }

        private void overtimecheck_FormClosing(object sender, FormClosingEventArgs e)
        {
            Process[] myProcesses;
            DateTime startTime;
            myProcesses = Process.GetProcessesByName("Excel");

            // 得不到Excel进程ID，暂时只能判断进程启动时间 
            foreach (Process myProcess in myProcesses)
            {
                startTime = myProcess.StartTime;

                if (startTime > beforeTime && startTime < afterTime)
                {
                    myProcess.Kill();
                }
            }
        }
    }
}