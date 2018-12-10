using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.IO;
using System.Diagnostics;
using DX_QMS.Common;

namespace DX_QMS
{
    public partial class IQCRevExcept_2 : DevExpress.XtraEditors.XtraForm
    {
        private string serverFilePath = "";
        public IQCRevExcept_2()
        {
            InitializeComponent();
            setRule();
            string path = System.Configuration.ConfigurationSettings.AppSettings["ServerFilePathRev"].ToString();
            this.serverFilePath = @path;
        }

        public IQCRevExcept_2(string productcode)
        {
            InitializeComponent();
            setRule();
            string path = System.Configuration.ConfigurationSettings.AppSettings["ServerFilePathRev"].ToString();
            this.serverFilePath = @path;
            txtlotno.Text = productcode;
            txtproductcode.Text = productcode;
            btnsearch_Click(null, null);
        }

        private void setRule()
        {
            string post = "";
            if (Login.manager != "")
            {
                post = Login.manager;
            }
            else
            {
                post = Login.post;
            }
            Dictionary<string, bool> dic = GroupPermission.QMS_SelectRulesForForm(post, "异常");
            this.btnsave.Enabled = bool.Parse(dic["hasInsert"].ToString());
            this.btndel.Enabled = bool.Parse(dic["hasDelete"].ToString());
        }

        private void txtlotno_KeyUp(object sender, KeyEventArgs e)
        {
            if (txtlotno.Text == "") return;
            if (e.KeyValue == 13)
                txtlotno_Leave(sender, null);
        }
        private DataSet bindData(string productcode)
        {
            string ssql = "select operdate 日期,productcode 产品编码,pname 描述,supplier 供应商,Reason 不良现象,remark 备注,operuser 维护人 from IQC_RevExcept where productcode like '" + productcode + "%'";
            return DbAccess.SelectBySql(ssql);
        }
        private void txtlotno_Leave(object sender, EventArgs e)
        {

            if (txtlotno.Text == "") return;
            txtlotno.Leave -= txtlotno_Leave;
            string Orasql = "select organization_id organization,segment1 materialcode,description materialname,en_des materialenname,INVENTORY_ITEM_ID,PRIMARY_UOM_CODE from cux_item_material_v where segment1='" + txtlotno.Text.Trim() + "'";
            DataSet ds = DbAccess.SelectByOracle(Orasql);
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                txtlotno.Enabled = false;
                txtproductcode.Text = ds.Tables[0].Rows[0]["materialcode"].ToString();
                txtproductdesc.Text = ds.Tables[0].Rows[0]["materialname"].ToString();
                lblinfo.Text = txtproductcode.Text + "正确,请添加不良项";
                lblinfo.ForeColor = Color.Blue;
                txtsupplier.Text = "";
                txtsupplier.Focus();
                DataTable des  = bindData(txtproductcode.Text).Tables[0];
                databind.DataSource = des;
                lblinfo2.Text = txtproductcode.Text + "总共有不良项目数:" + des.Rows.Count;
                lblinfo2.ForeColor = Color.Blue;
                txtlotno.Leave += txtlotno_Leave;
                return;
            }
            string ssql = "select materialcode from delivery where lotno='" + txtlotno.Text + "'";
            DataSet dslotno = DbAccess.SelectBySql(ssql);
            if (dslotno != null && ds.Tables.Count > 0 && dslotno.Tables[0].Rows.Count > 0)
            {
                string materialcode = dslotno.Tables[0].Rows[0][0].ToString();
                string Orasql2 = "select organization_id organization,segment1 materialcode,description materialname,en_des materialenname,INVENTORY_ITEM_ID,PRIMARY_UOM_CODE from cux_item_material_v where segment1='" + materialcode + "'";
                DataSet ds2 = DbAccess.SelectByOracle(Orasql2);
                if (ds2 != null && ds2.Tables.Count > 0 && ds2.Tables[0].Rows.Count > 0)
                {
                    txtlotno.Enabled = false;
                    txtproductcode.Text = ds2.Tables[0].Rows[0]["materialcode"].ToString();
                    txtproductdesc.Text = ds2.Tables[0].Rows[0]["materialname"].ToString();
                    lblinfo.Text = txtproductcode.Text + "正确,请添加不良项";
                    lblinfo.ForeColor = Color.Blue;
                    txtsupplier.Text = "";
                    txtsupplier.Focus();
                    DataTable des = bindData(txtproductcode.Text).Tables[0];
                    databind.DataSource = des;
                    lblinfo2.Text = txtproductcode.Text + "总共有不良项目数:" + des.Rows.Count;
                    lblinfo2.ForeColor = Color.Blue;
                    txtlotno.Leave += txtlotno_Leave;
                    return;
                }
            }
            else
            {
                lblinfo.Text = txtlotno.Text + "不正确,请重新输入";
                lblinfo.ForeColor = Color.Red;
                txtsupplier.Text = "";
                txtlotno.SelectAll();
                txtlotno.Enabled = true;
                txtlotno.Focus();
            }

            txtlotno.Leave += txtlotno_Leave;

        }

        private void txtsupplier_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13)
                txtReason.Focus();
        }

        private void txtReason_KeyUp(object sender, KeyEventArgs e)
        {
            if (txtReason.Text.Trim() == "") return;
            if (e.KeyValue == 13)
                btnsave_Click(sender, null);
        }

        private void btnpiclist_Click(object sender, EventArgs e)
        {

            OpenFileDialog ofdImport = new OpenFileDialog();

            ofdImport.Filter = "文件(*.jpg;bmp;png;jpeg;pdf)|*.jpg;*.bmp;*.png;*.jpeg;*.pdf";
            ofdImport.Multiselect = false;
            DialogResult dr = ofdImport.ShowDialog();
            if (dr == DialogResult.Cancel) return;
            this.txtpiclist.Text = "";
            foreach (string str in ofdImport.FileNames)
            {
                this.txtpiclist.Text += str + ",";
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

        private bool CopyFileToServer(string folderType, string filePath)
        {
            try
            {
                string[] fileName = @filePath.Split('\\');
                string floerPath = "\\\\" + this.serverFilePath + "\\" + folderType + "\\" + this.txtproductcode.Text.Trim() + "\\" + txtReason.Text;
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

        private void UploadFile(string filepath)
        {

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
                        if (!CopyFileToServer("来料异常", ss))
                        {
                            msg += ss + ";";
                            re = false;
                        }
                    }
                }
                if (re)
                {
                    MessageBox.Show("操作成功");

                }
                else MessageBox.Show(msg + " 文件上传失败");
            }
            else
            {
                MessageBox.Show("无法连接到服务器的共享目录");
            }
        }
        private void btnsave_Click(object sender, EventArgs e)
        {
            if (txtproductcode.Text.Trim() == "") return;

            string fileDBServerPath = "";
            if (txtpiclist.Text != "")
            {
                string[] fileqty = txtpiclist.Text.TrimEnd(',').Split(',');
                foreach (string s in fileqty)
                {
                    string[] fileName = @s.Split('\\');
                    fileDBServerPath += "\\\\" + this.serverFilePath + "\\来料异常" + "\\" + this.txtproductcode.Text.Trim() + "\\" + txtReason.Text + "\\" + fileName[fileName.Length - 1] + ",";
                }
            }
            string floerPath = "";
            floerPath = "\\\\" + this.serverFilePath + "\\来料异常" + "\\" + this.txtproductcode.Text.Trim() + "\\" + txtReason.Text;


            if (this.txtproductcode.Text == "" || txtReason.Text == "") return;
            string sql = "if not exists(select 1 from IQC_RevExcept where productcode='" + txtproductcode.Text.Trim() + "' and Reason='" + this.txtReason.Text + "'" + ")";
            sql += " insert into IQC_RevExcept(productcode,pname,supplier,Reason,operuser,operdate) values('" + txtproductcode.Text + "','";
            sql += txtproductdesc.Text + "','" + txtsupplier.Text + "','" + txtReason.Text + "','" + Login.username + "',getdate()" + ",'" + txtremark.Text + "')";
            if (DbAccess.ExecuteSql(sql))
            {
                if (floerPath != "")
                    del_prefile(floerPath, txtReason.Text);
                UploadFile(fileDBServerPath);
                DataTable des = bindData(txtproductcode.Text).Tables[0];
                databind.DataSource = des;
                lblinfo2.Text = txtproductcode.Text + "总共有不良项目数:" + des.Rows.Count;
                lblinfo2.ForeColor = Color.Blue;
                txtReason.Text = "";
                txtsupplier.Text = "";
                txtpiclist.Text = "";
                txtsupplier.Focus();
            }
            else
            {
                MessageBox.Show("新增失败," + txtproductcode.Text + "," + txtsupplier.Text + "该编码已存在!");
                txtReason.Text = "";
                txtsupplier.Text = "";
                txtsupplier.Focus();
            }
        }

        private void btndel_Click(object sender, EventArgs e)
        {
            DataTable dt = databind.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
                return;
            if (gridView.FocusedRowHandle < 0)
                return;
            if (gridView.GetSelectedRows().Length < 0)
                return;
            string sql = "";
            for (int i = 0; i < gridView.GetSelectedRows().Length; i++)
            {
                sql = "delete IQC_RevExcept where productcode='" + gridView.GetDataRow(gridView.GetSelectedRows()[i])["产品编码"].ToString() + "' and Reason='" + gridView.GetDataRow(gridView.GetSelectedRows()[i])["不良现象"].ToString() + "'";
                Common.DbAccess.ExecuteSql(sql);

                string floerPath = "";
                floerPath = "\\\\" + this.serverFilePath + "\\来料异常" + "\\" + gridView.GetDataRow(gridView.GetSelectedRows()[i])["产品编码"].ToString() + "\\" + gridView.GetDataRow(gridView.GetSelectedRows()[i])["不良现象"].ToString();
                string filename = gridView.GetDataRow(gridView.GetSelectedRows()[i])["产品编码"].ToString();
                del_prefile(floerPath, filename);
            }

            DataTable ded = bindData(txtproductcode.Text).Tables[0];
            databind.DataSource = ded;
            lblinfo2.Text = txtproductcode.Text + "总共有不良项目数:" + ded.Rows.Count;
            lblinfo2.ForeColor = Color.Blue;
            txtsupplier.Focus();
        }

        private void btnreset_Click(object sender, EventArgs e)
        {
            txtsupplier.Text = "";
            txtReason.Text = "";
            txtlotno.Text = "";
            txtproductcode.Text = "";
            txtproductdesc.Text = "";
            txtpiclist.Text = "";
            txtremark.Text = "";
            txtlotno.Enabled = true;
            txtlotno.Focus();
        }

        private void btnsearch_Click(object sender, EventArgs e)
        {
            DataSet ds = bindData(txtlotno.Text);
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                databind.DataSource = ds.Tables[0];
                lblinfo2.Text = txtproductcode.Text + "总共有不良项目数:" + ds.Tables[0].Rows.Count;
                lblinfo2.ForeColor = Color.Blue;
            }
            else
                MessageBox.Show("没有符合条件的记录");
        }
        private void OpenFile(string filepath, string pdffile)
        {
            string filename = "";
            filename = pdffile;
            //定义一个ProcessStartInfo实例
            System.Diagnostics.ProcessStartInfo info = new System.Diagnostics.ProcessStartInfo();
            //设置启动进程的初始目录

            //filename = fileName[fileName.Length - 1];
            info.FileName = @filename;
            info.WorkingDirectory = filepath;
            //设置启动进程的参数
            info.Arguments = "";
            //启动由包含进程启动信息的进程资源
            try
            {
                System.Diagnostics.Process.Start(info);

            }

            catch (System.ComponentModel.Win32Exception we)
            {

                MessageBox.Show(this, we.Message);
                return;
            }

        }
        private void gridView_DoubleClick(object sender, EventArgs e)
        {
            DataTable dt = databind.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 0)
                return;
            if (gridView.FocusedRowHandle < 0)
                return;
            string pno = gridView.GetFocusedRowCellValue("产品编码").ToString();
            string reason = gridView.GetFocusedRowCellValue("不良现象").ToString();
            string floerPath = "\\\\" + this.serverFilePath + "\\来料异常" + "\\" + pno + "\\" + reason;
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
                            //string filename = Path.GetFileNameWithoutExtension(pt[i].ToString());
                            string filename = Path.GetFileName(pt[i].ToString());
                            OpenFile(floerPath, filename);
                        }
                    }
                }
            }
        }

        private void IQCRevExcept_2_Load(object sender, EventArgs e)
        {
            txtReason.Text = "控制15个字以内";
        }
    }
}