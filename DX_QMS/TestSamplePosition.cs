using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using DevExpress.XtraBars;
using DX_QMS.Common;
using System.Diagnostics;
using System.Configuration;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraEditors;

namespace DX_QMS
{
    public partial class TestSamplePosition : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private string serverFilePath = "";
        private DataTable dtsuppall;
        private IQC ic = new IQC();

        public TestSamplePosition()
        {
            InitializeComponent();
            setRule();
            string path = System.Configuration.ConfigurationSettings.AppSettings["ServerFilePath"].ToString();
            //string path = System.Configuration.ConfigurationManager.AppSettings["ServerFilePath"].ToString();
            this.serverFilePath = @path;
            this.txtsampletype.SelectedIndex = 0;
            txttempqty.Enabled = false;
            txttempqty.Text = "0";
            bindSup();
            DataTable dtblock = bindSuppAndBlock("楼层");
            if (dtblock.Rows.Count > 0)
            {           
                txtblock.Properties.Items.Clear();
                foreach (DataRow row in dtblock.Rows)
                {
                    txtblock.Properties.Items.Add(row["Definevalue"]);
                }
                txtblock.SelectedIndex = 0;
            }
            gridView.PrintExportProgress += new ProgressChangedEventHandler(gridView_PrintExportProgress);
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
            Dictionary<string, bool> dic = GroupPermission.QMS_SelectRulesForForm(post, "样品管理");
            btnsave.Enabled = bool.Parse(dic["hasInsert"].ToString());
            btndel.Enabled = bool.Parse(dic["hasDelete"].ToString());
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Enter)
            {

                if (ActiveControl is TextBox || ActiveControl is System.Windows.Forms.ComboBox || ActiveControl is DateTimePicker)
                {
                    if (ActiveControl.Name == "txtsuppsearch")
                    {
                        return false;
                    }
                    else
                    {
                        System.Windows.Forms.SendKeys.Send("{tab}");
                        return true;
                    }
                }
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void TestSamplePosition_Load(object sender, EventArgs e)
        {
            txtsampleClass.SelectedIndex = 0;
        }

        private void bindSup()
        {
            DataTable dtsup = bindSuppAndBlock("供应商");
            if (dtsup.Rows.Count > 0)
            { 
                txtsupp.Properties.Items.Clear();
                foreach (DataRow row in dtsup.Rows)
                {
                    txtsupp.Properties.Items.Add(row["Definevalue"]);
                }
                txtsupp.SelectedIndex = 0;
                dtsuppall = dtsup;
            }
        }
        private DataTable bindSuppAndBlock(string supplier)
        {
            string ssql = "select Definetype,Definevalue from OQC_TypeDefine where Definetype='" + supplier + "' order by sort ";
            DataTable dt = Common.DbAccess.SelectBySql(ssql).Tables[0];
            return dt;
        }

        private void txtproductcode_Leave(object sender, EventArgs e)
        {
            if (txtproductcode.Text == "") return;
            txtproductcode.Leave -= txtproductcode_Leave;
            string Orasql = "select organization_id organization,segment1 materialcode,description materialname,en_des materialenname,INVENTORY_ITEM_ID,PRIMARY_UOM_CODE from cux_item_material_v where segment1='" + txtproductcode.Text.Trim() + "'";
            DataSet ds = Common.DbAccess.SelectByOracle(Orasql);
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                lblinfo.Text = ds.Tables[0].Rows[0]["materialname"].ToString();
                lblinfo.ForeColor = Color.Blue;
                txtblock.Focus();
            }
            else
            {
                string ssql = "select materialcode from delivery where lotno='" + txtproductcode.Text + "'";
                DataSet dslotno = Common.DbAccess.SelectBySql(ssql);
                if (dslotno != null && ds.Tables.Count > 0 && dslotno.Tables[0].Rows.Count > 0)
                {
                    string materialcode = dslotno.Tables[0].Rows[0]["materialcode"].ToString();
                    string Orasqlbylotno = "select organization_id organization,segment1 materialcode,description materialname,en_des materialenname,INVENTORY_ITEM_ID,PRIMARY_UOM_CODE from cux_item_material_v where segment1='" + materialcode + "'";
                    ds = Common.DbAccess.SelectByOracle(Orasqlbylotno);
                    if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                    {
                        txtproductcode.Text = materialcode;
                        lblinfo.Text = ds.Tables[0].Rows[0]["materialname"].ToString();
                        lblinfo.ForeColor = Color.Blue;
                        txtblock.Focus();
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


        private void txtproductcode_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13 && txtproductcode.Text != "")
            {
                txtproductcode_Leave(sender, e);
            }
        }


        private void txtblock_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txtsampletype.Text == "正式样")
            {
                this.txttempqty.Enabled = false;
                this.txttempqty.Text = "0";
            }
            else
            {
                txttempqty.Enabled = true;
            }
        }

        private void txtsupp_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txtsampletype.Text == "正式样")
            {
                this.txttempqty.Enabled = false;
                this.txttempqty.Text = "0";
            }
            else
            {
                txttempqty.Enabled = true;
            }
        }

        private void txtsuppsearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13)
            {
                btnsuppsearch_Click(sender, null);
            }
        }

        private void btnsuppsearch_Click(object sender, EventArgs e)
        {
            bindSup();
            DataTable dtnew = this.dtsuppall.Clone();
            DataRow[] rw = dtsuppall.Select("Definevalue like '" + txtsuppsearch.Text.Trim() + "%'");
            for (int i = 0; i < rw.Length; i++)
            {
                dtnew.Rows.Add(rw[i].ItemArray);
            }
            txtsupp.Properties.Items.Clear();
            foreach (DataRow row in dtnew.Rows)
            {
                txtsupp.Properties.Items.Add(row["Definevalue"]);
            }
            txtsupp.SelectedIndex = 0;
        }

        private void txtsampletype_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txtsampletype.Text == "正式样")
            {
                this.txttempqty.Enabled = false;
                this.txttempqty.Text = "0";
            }
            else
            {
                txttempqty.Enabled = true;
            }
        }

        private void btnbrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofdImport = new OpenFileDialog();

            ofdImport.Filter = "文件(*.jpg;bmp;png;jpeg;pdf)|*.jpg;*.bmp;*.png;*.jpeg;*.pdf";
            ofdImport.Multiselect = false;
            DialogResult dr = ofdImport.ShowDialog();
            if (dr == DialogResult.Cancel) return;
            this.txtelecbook.Text = "";
            foreach (string str in ofdImport.FileNames)
            {
                this.txtelecbook.Text += str + ",";
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
                        string sfilename = Path.GetFileNameWithoutExtension(ss);
                        if (sfilename == filename)
                        {
                            if (System.IO.File.Exists(ss))
                                System.IO.File.Delete(ss);
                        }
                    }
                }
            }

        }

        //IQCProdPic ipp = new IQCProdPic();
        private bool CopyFileToServer(string folderType, string sfilePath)
        {
            try
            {
                string[] fileName = @sfilePath.Split('\\');
                string floerPath = "\\\\" + this.serverFilePath + "\\" + folderType + "\\" + this.txtsupp.Text + this.txtproductcode.Text.Trim();
                string fileServerPath = floerPath + "\\" + fileName[fileName.Length - 1];
                if (Directory.Exists(floerPath) == false)//如果不存在就创建file文件夹
                {
                    Directory.CreateDirectory(floerPath);
                }
                File.Copy(sfilePath, fileServerPath, true);
                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }
        private void UploadFile(string filename)
        {

            if (Connect(serverFilePath))
            {
                bool re = true;
                string msg = "";
                string te = filename.TrimEnd(',');
                if (te != "")
                {
                    string[] copy = te.Split(',');
                    foreach (string ss in copy)
                    {
                        if (!CopyFileToServer("电子档", ss))
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
            if (txtproductcode.Text == "") return;
            if (txtblock.Text == "" || txtsupp.Text  == "") return;
            int m = 0;
            if (!int.TryParse(txttempqty.Text == "" ? "0" : txttempqty.Text, out m))
            {
                lblinfo.Text = "数量格式不正确!";
                lblinfo.ForeColor = Color.Red;
                txttempqty.Text = "";
                txttempqty.Focus();
                return;
            }
            if (txtsampletype.Text == "")
                return;


            if (txtsampleClass.Text == "实物样" && txtsampleCode.Text == "")
            {
                MessageBox.Show("实物样的虚拟编码不能为空","提醒",MessageBoxButtons.OK ,MessageBoxIcon.Information);
                return;
            }

            string floerPath = "";
            floerPath = "\\\\" + this.serverFilePath + "\\电子档" + "\\" + txtsupp.Text.Trim() + this.txtproductcode.Text.Trim();
            if (txtelecbook.Text != "")
            {
                if (floerPath != "")
                    del_prefile(floerPath, txtsupp.Text.Trim() + this.txtproductcode.Text.Trim());
                UploadFile(txtelecbook.Text);
            }

            string ifYesNo = "0";
            if (cbYes.Checked)
                ifYesNo = "1";
            if (cbNo.Checked)
                ifYesNo = "0";

            //int i = ic.AddNewTestSamplePosition("新增", txtproductcode.Text.Trim(), txtposition.Text, date.Value.ToString(), txtpaperposition.Text, txtsampletype.Text, txtsampledate.Value.ToString(),
            //                                    txttempqty.Text, txtdeveloper.Text, txtSQEengineer.Text, txtremarks.Text, txtsupp.Text.Trim(), txtblock.Text.Trim(), Login.username, ifYesNo);
         
            int i = ic.AddNewTestSamplePositionNew("新增", txtproductcode.Text.Trim(), txtposition.Text, date.Value.ToString(), txtpaperposition.Text, txtsampletype.Text, txtsampleClass.Text, txtsampleCode.Text, txtsampledate.Value.ToString(),
                                               txttempqty.Text, txtdeveloper.Text, txtSQEengineer.Text, txtremarks.Text, txtsupp.Text.Trim(), txtblock.Text.Trim(), Login.username, ifYesNo);


            if (i > 0)
            {
                //DataSet ds = ic.SelectTestSamplePosition("查询", txtproductcode.Text.Trim(), txtposition.Text, date.Value.ToString(), txtpaperposition.Text, txtsampletype.Text, txtsampledate.Value.ToString(),
                //                                txttempqty.Text, txtdeveloper.Text, txtSQEengineer.Text, txtremarks.Text, txtsupp.Text.Trim(), txtblock.Text.Trim(), Login.username, ifYesNo);

                DataSet ds = ic.SelectTestSamplePositionNew("查询", txtproductcode.Text.Trim(), txtposition.Text, date.Value.ToString(), txtpaperposition.Text, txtsampletype.Text,txtsampleClass.Text, txtsampleCode.Text, txtsampledate.Value.ToString(),
                                  txttempqty.Text, txtdeveloper.Text, txtSQEengineer.Text, txtremarks.Text, txtsupp.Text.Trim(), txtblock.Text.Trim(), Login.username, ifYesNo);


                databind.DataSource = ds.Tables[0];
                txtproductcode.Text = "";
                txtposition.Text = "";
                lblinfo.Text = "";
                txtpaperposition.Text = "";
                txtdeveloper.Text = "";
                txtSQEengineer.Text = "";
                txtremarks.Text = "";
                txtelecbook.Text = "";
                txtproductcode.Focus();
            }
        }

        private void btnsearch_Click(object sender, EventArgs e)
        {
            string ifYesNo = "0";
            if (cbYes.Checked)
                ifYesNo = "1";
            if (cbNo.Checked)
                ifYesNo = "0";
            //DataSet ds = ic.SelectTestSamplePosition("查询", txtproductcode.Text.Trim(), txtposition.Text, date.Value.ToString(), txtpaperposition.Text, txtsampletype.Text, txtsampledate.Value.ToString(),
            //    txttempqty.Text, txtdeveloper.Text, txtSQEengineer.Text, txtremarks.Text, txtsupp.Text.Trim() == "" ? "" : txtsupp.Text, txtblock.Text.Trim() == "" ? "" : txtblock.Text, Login.username, ifYesNo);

            DataSet ds = ic.SelectTestSamplePositionNew("查询", txtproductcode.Text.Trim(), txtposition.Text, date.Value.ToString(), txtpaperposition.Text, txtsampletype.Text, txtsampleClass.Text, txtsampleCode.Text, txtsampledate.Value.ToString(),
                           txttempqty.Text, txtdeveloper.Text, txtSQEengineer.Text, txtremarks.Text, txtsupp.Text.Trim(), txtblock.Text.Trim(), Login.username, ifYesNo);

            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                databind.DataSource = ds.Tables[0];
            }
            else
                MessageBox.Show("没有满足条件的记录");


        }

        private void btndel_Click(object sender, EventArgs e)
        {
            if (gridView.FocusedRowHandle < 0)
                return;
            if (gridView.GetSelectedRows().Length < 0)
                return;

            for (int i = gridView.GetSelectedRows().Length; i > 0; i--)
            {
                DataRow dr = gridView.GetDataRow(gridView.GetSelectedRows()[i - 1]);
                string supper = gridView.GetDataRow(gridView.GetSelectedRows()[i - 1])["供应商"].ToString() + gridView.GetDataRow(gridView.GetSelectedRows()[i - 1])["产品编码"].ToString();
                int m = ic.AddNewTestSamplePosition("删除", dr["产品编码"].ToString(), "", "", "", "", "", "", "", "", "", dr["供应商"].ToString(), "", "","");
                if (m > 0)
                {
                    //this.databind.Rows.RemoveAt(databind.SelectedRows[i - 1].Index);
                    gridView.DeleteRow(gridView.GetSelectedRows()[i - 1]);
                }
                string floerPath = "";
                floerPath = "\\\\" + this.serverFilePath + "\\电子档" + "\\" + supper;
                string filename = supper;
                del_prefile(floerPath, filename);
            }
        }

        private void btnreset_Click(object sender, EventArgs e)
        {
            databind.DataSource = null;
            txtproductcode.Text = "";
            txtposition.Text = "";
            txtpaperposition.Text = "";
            txtsampletype.Text = "";
            txtsampledate.Value = DateTime.Now;
            txtdeveloper.Text = "";
            txtSQEengineer.Text = "";
            txtremarks.Text = "";
            txttempqty.Text = "";
            txtsuppsearch.Text = "";
            txtelecbook.Text = "";
            txtblock.Text = "";
            txtsupp.Text = "";
            txtsampleCode.Text = "";
            lblinfo.Text = "";

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
            string name = "样品存放位置";
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
            // Solder so = new Solder();
            DataTable dt = databind.DataSource as DataTable;
            if (dt == null)
                return;
            if (dt.Rows.Count <= 0)
                return;

           // DataToExcel(databind);


            string fileName = ShowSaveFileDialog("Microsoft Excel 2007 Document", "Microsoft Excel|*.xlsx");
            if (fileName == string.Empty) return;
            ExportToEx(fileName, "xlsx", gridView);
            OpenFile(fileName);

        }

        private void btnsearchpho_Click(object sender, EventArgs e)
        {
    //        DataSet ds = ic.SelectTestSamplePosition("查询", txtproductcode.Text.Trim(), txtposition.Text, date.Value.ToString(), txtpaperposition.Text, txtsampletype.Text, txtsampledate.Value.ToString(),
    //txttempqty.Text, txtdeveloper.Text, txtSQEengineer.Text, txtremarks.Text, txtsupp.Text.Trim() == "" ? "" : txtsupp.Text, txtblock.Text.Trim() == "" ? "" : txtblock.Text, Login.username,"");

            DataSet ds = ic.SelectTestSamplePositionNew("查询", txtproductcode.Text.Trim(), txtposition.Text, date.Value.ToString(), txtpaperposition.Text, txtsampletype.Text, txtsampleClass.Text, txtsampleCode.Text, txtsampledate.Value.ToString(),
                           txttempqty.Text, txtdeveloper.Text, txtSQEengineer.Text, txtremarks.Text, txtsupp.Text.Trim(), txtblock.Text.Trim(), Login.username, "");

            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                ds.Tables[0].Columns.Add("pho");
                if (Connect(serverFilePath))
                {
                    string pho = "";
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        string floerPath = "\\\\" + this.serverFilePath + "\\电子档" + "\\" + ds.Tables[0].Rows[i]["供应商"].ToString() + ds.Tables[0].Rows[i]["产品编码"].ToString();
                        if (Directory.Exists(floerPath))
                        {
                            string[] pt = Directory.GetFiles(floerPath);
                            if (pt.Length > 0)
                                pho = "有";
                            ds.Tables[0].Rows[i]["pho"] = pho;
                        }
                    }
                }
                databind.DataSource = ds.Tables[0];
               
            }
            else
                MessageBox.Show("没有满足条件的记录");
        }

        private void databind_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //txtproductcode.Text = databind.CurrentRow.Cells["产品编码"].Value.ToString();
            //txtposition.Text = databind.CurrentRow.Cells["存放位置"].Value.ToString();
            //txtpaperposition.Text = databind.CurrentRow.Cells["图纸位置"].Value.ToString();
            //txtsampletype.Text = databind.CurrentRow.Cells["样品类别"].Value.ToString();
            //txtsampledate.Value = DateTime.Parse(databind.CurrentRow.Cells["签样日期"].Value.ToString());
            //txtdeveloper.Text = databind.CurrentRow.Cells["研发工程师"].Value.ToString();
            //txtSQEengineer.Text = databind.CurrentRow.Cells["SQE工程师"].Value.ToString();
            //txtremarks.Text = databind.CurrentRow.Cells["备注"].Value.ToString();
            //txttempqty.Text = databind.CurrentRow.Cells["数量"].Value.ToString();

            txtproductcode.Text = gridView.GetFocusedRowCellValue("产品编码").ToString();
            txtposition.Text = gridView.GetFocusedRowCellValue("存放位置").ToString();
            txtpaperposition.Text = gridView.GetFocusedRowCellValue("图纸位置").ToString();
            txtsampletype.Text = gridView.GetFocusedRowCellValue("样品类别").ToString();
            txtsampledate.Value = DateTime.Parse(gridView.GetFocusedRowCellValue("签样日期").ToString());
            txtdeveloper.Text = gridView.GetFocusedRowCellValue("研发工程师").ToString();
            txtSQEengineer.Text = gridView.GetFocusedRowCellValue("SQE工程师").ToString();
            txtremarks.Text = gridView.GetFocusedRowCellValue("备注").ToString();
            txttempqty.Text = gridView.GetFocusedRowCellValue("数量").ToString();



            try
            {
                //txtsupp.SelectedValue = databind.CurrentRow.Cells["供应商"].Value.ToString();
                //txtblock.SelectedValue = databind.CurrentRow.Cells["楼层"].Value.ToString();
                txtsampleClass.Text = gridView.GetFocusedRowCellValue("样品种类").ToString();
                //txtsampleCode.Text = gridView.GetFocusedRowCellValue("虚拟代码").ToString();
                txtsupp.Text = gridView.GetFocusedRowCellValue("供应商").ToString();
                txtblock.Text = gridView.GetFocusedRowCellValue("楼层").ToString();
            }
            catch { }
        }

        private void OpenFile(string filepath, string pdffile)
        {
            string filename = "";
            filename = pdffile + ".pdf";
            //定义一个ProcessStartInfo实例
            System.Diagnostics.ProcessStartInfo info = new System.Diagnostics.ProcessStartInfo();
            //设置启动进程的初始目录
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

        private void databind_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //if (btnsave.Enabled == false) return;
            //if (databind.CurrentRow.Cells["样品类别"].Value.ToString() == "临时样")
            //{
            //    TestSampleList TList = new TestSampleList(databind.CurrentRow.Cells["样品类别"].Value.ToString(), databind.CurrentRow.Cells["产品编码"].Value.ToString(), databind.CurrentRow.Cells["供应商"].Value.ToString());
            //    DialogResult dr = TList.ShowDialog();
            //}
            //string floerPath = "\\\\" + this.serverFilePath + "\\电子档" + "\\" + databind.CurrentRow.Cells["供应商"].Value.ToString() + txtproductcode.Text;
            //if (Connect(serverFilePath))
            //{
            //    if (Directory.Exists(floerPath))
            //    {
            //        string[] pt = Directory.GetFiles(floerPath);
            //        if (pt.Length == 0) return;
            //        if (pt.Length > 0)
            //        {
            //            for (int i = 0; i < pt.Length; i++)
            //            {
            //                string filename = Path.GetFileNameWithoutExtension(pt[i].ToString());

            //                FileStream fs = new FileStream(pt[i].ToString(), FileMode.Open);
            //                BinaryReader reader = new BinaryReader(fs);
            //                string fileclass = "";
            //                try
            //                {
            //                    for (int j = 0; j < 2; j++)
            //                    {
            //                        fileclass += reader.ReadByte().ToString();
            //                    }

            //                }
            //                catch (Exception)
            //                {

            //                    throw;
            //                }

            //                OpenFile(floerPath, filename);
            //                fs.Close();
            //                fs.Dispose();
            //                reader.Close();
            //            }
            //        }
            //    }
            //}
        }

        private void databind_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            //DataGridViewRow dgr = databind.Rows[e.RowIndex];
            //try
            //{
            //    if (dgr.Cells["状态"].Value.ToString() == "已过期")
            //    {
            //        dgr.DefaultCellStyle.BackColor = Color.Red;
            //    }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}
        }

        private void gridView_Click(object sender, EventArgs e)
        {
            DataTable de = databind.DataSource as DataTable;
            if (de == null || de.Rows.Count < 1)
                return;
            if (gridView.RowCount < 1)
                return;
            
            txtproductcode.Text = gridView.GetFocusedRowCellValue("产品编码").ToString();
            txtposition.Text = gridView.GetFocusedRowCellValue("存放位置").ToString();
            txtpaperposition.Text = gridView.GetFocusedRowCellValue("图纸位置").ToString();
            txtsampletype.Text = gridView.GetFocusedRowCellValue("样品类别").ToString();
            txtsampledate.Value = DateTime.Parse(gridView.GetFocusedRowCellValue("签样日期").ToString());
            txtdeveloper.Text = gridView.GetFocusedRowCellValue("研发工程师").ToString();
            txtSQEengineer.Text = gridView.GetFocusedRowCellValue("SQE工程师").ToString();
            txtremarks.Text = gridView.GetFocusedRowCellValue("备注").ToString();
            txttempqty.Text = gridView.GetFocusedRowCellValue("数量").ToString();
            try
            {
                txtsampleClass.Text = gridView.GetFocusedRowCellValue("样品种类").ToString();
                //txtsampleCode.Text = gridView.GetFocusedRowCellValue("虚拟代码").ToString();
                txtsupp.Text = gridView.GetFocusedRowCellValue("供应商").ToString();
                txtblock.Text = gridView.GetFocusedRowCellValue("楼层").ToString();
            }
            catch { }
        }

        private void gridView_DoubleClick(object sender, EventArgs e)
        {
            DataTable de = databind.DataSource as DataTable;
            if (de == null || de.Rows.Count < 1)
                return;
            if (gridView.RowCount < 1)
                return;

            if (btnsave.Enabled == false) return;
            if (gridView.GetFocusedRowCellValue("样品类别").ToString() == "临时样")//databind.CurrentRow.Cells["样品类别"].Value.ToString() == "临时样"
            {
                TestSampleList TList = new TestSampleList(gridView.GetFocusedRowCellValue("样品类别").ToString(), gridView.GetFocusedRowCellValue("产品编码").ToString(), gridView.GetFocusedRowCellValue("供应商").ToString());
                DialogResult dr = TList.ShowDialog();
            }
            string floerPath = "\\\\" + this.serverFilePath + "\\电子档" + "\\" + gridView.GetFocusedRowCellValue("供应商").ToString() + txtproductcode.Text;
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
                            string filename = Path.GetFileNameWithoutExtension(pt[i].ToString());

                            FileStream fs = new FileStream(pt[i].ToString(), FileMode.Open);
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

                            OpenFile(floerPath, filename);
                            fs.Close();
                            fs.Dispose();
                            reader.Close();
                        }
                    }
                }
            }
        }

        private void gridView_CustomDrawCell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
        {
            //////try
            //////{
            //////    if (e.Column.FieldName == "状态")
            //////    {
            //////        GridCellInfo GridCell = e.Cell as GridCellInfo;
            //////        if (GridCell.CellValue.ToString() == "已过期")
            //////        {
            //////            e.Appearance.BackColor = Color.Red;
            //////        }

            //////    }
            //////}
            //////catch (Exception ex)
            //////{
            //////    MessageBox.Show(ex.Message);
            //////}

        }

        private void gridView_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
          {
 
            DataTable dt = databind.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
            {
                return;
            }

            if (gridView.RowCount < 1)
                return;

            if (gridView.GetDataRow(e.RowHandle)["状态"].ToString() == "已过期")
            {
                e.Appearance.BackColor = Color.Red;
            }

        }
    }
}