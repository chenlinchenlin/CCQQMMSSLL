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
using System.Diagnostics;
using System.IO;
using DX_QMS.Common;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraEditors;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using System.Net;

namespace DX_QMS
{
    public partial class TestROHS : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private string serverFilePath = "";
        private IQC ic = new IQC();
        public TestROHS()
        {
            InitializeComponent();
            string path = System.Configuration.ConfigurationManager.AppSettings["ServerFilePathRohs"].ToString();
            /////this.serverFilePath = @path;
            this.serverFilePath = @"10.100.0.150\rohs$";
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
            Dictionary<string, bool> dic = GroupPermission.QMS_SelectRulesForForm(post, "RoHS");
            btnsave.Enabled = bool.Parse(dic["hasInsert"].ToString());
            btnupdate.Enabled = bool.Parse(dic["hasUpdate"].ToString());
            btndel.Enabled = bool.Parse(dic["hasDelete"].ToString());
        }
        private void TestROHS_Load(object sender, EventArgs e)
        {
            txtsenttestdate.EditValue = DateTime.Now.ToLocalTime().ToString();
            date.EditValue = DateTime.Now.ToLocalTime().ToString();
           // setRule();
            gridView.PrintExportProgress += new ProgressChangedEventHandler(gridView_PrintExportProgress);
        }
        public DataTable dtpub;
        private void bindReliabilityType(string productcode)
        {
            //DataTable dt = ic.SelectTestTypeRecord("查询", "", "可靠性测试项", "").Tables[0];
            DataTable dt = ic.SelectTestROHS("测试项查询", productcode, "", "", "", "", Login.username, 0, "", "", "", "").Tables[0];
            if (dt.Rows.Count > 0)
            {
                dtpub = dt;              
                cbTestType.Properties.Items.Clear();
                foreach (DataRow row in dt.Rows)
                {
                    cbTestType.Properties.Items.Add(row["TestType"]);
                }
                cbTestType.SelectedIndex = 0;

                DataRow[] rw = dtpub.Select("TestType='" + cbTestType.Text + "'");
                if (rw.Length > 0)
                    txtdesc.Text = rw[0]["TestDesc"].ToString();
            }
            else
                cbTestType.Properties.Items.Clear();

        }
        private void txtproductcode_Leave(object sender, EventArgs e)
        {
            if (txtproductcode.Text.Trim() == "")
                return;
            txtproductcode.Leave -= txtproductcode_Leave;
            string Orasql = "select organization_id organization,segment1 materialcode,description materialname,en_des materialenname,INVENTORY_ITEM_ID,PRIMARY_UOM_CODE from cux_item_material_v where rownum=1 and segment1='" + txtproductcode.Text.Trim() + "'";
            DataSet ds = Common.DbAccess.SelectByOracle(Orasql);
            databind.DataSource = null;
            txtdesc.Text = "";
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                bindReliabilityType(txtproductcode.Text);
                if (cbTestType.Properties.Items.Count <= 0)
                {
                    MessageBox.Show("没有维护料号:" + txtproductcode.Text + ",可靠性/RoHS测试项", "系统提醒", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtproductcode.Focus();
                    txtproductcode.Text = "";
                    txtproductcode.Leave += txtproductcode_Leave;
                    return;
                }
                lblinfo.Text = ds.Tables[0].Rows[0]["materialname"].ToString();
                lblinfo.ForeColor = Color.Blue;

                txtTestResult.Focus();
                txtTestResult.SelectAll();
                btnsearch_Click(sender, e);
            }
            else
            {
                string ssql = "select materialcode from delivery where lotno='" + txtproductcode.Text + "'";
                DataSet dslotno = Common.DbAccess.SelectBySql(ssql);
                if (dslotno != null && ds.Tables.Count > 0 && dslotno.Tables[0].Rows.Count > 0)
                {
                    string materialcode = dslotno.Tables[0].Rows[0]["materialcode"].ToString();
                    string Orasqlbylotno = "select organization_id organization,segment1 materialcode,description materialname,en_des materialenname,INVENTORY_ITEM_ID,PRIMARY_UOM_CODE from cux_item_material_v where rownum=1 and segment1='" + materialcode + "'";
                    ds = Common.DbAccess.SelectByOracle(Orasqlbylotno);
                    if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                    {
                        bindReliabilityType(txtproductcode.Text);
                        if (cbTestType.Properties.Items.Count <= 0)
                        {
                            MessageBox.Show("没有维护料号:" + txtproductcode.Text + ",可靠性测试项", "系统提醒", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            txtproductcode.Focus();
                            txtproductcode.Text = "";
                            txtproductcode.Leave += txtproductcode_Leave;
                            return;
                        }
                        txtproductcode.Text = materialcode;
                        lblinfo.Text = ds.Tables[0].Rows[0]["materialname"].ToString();
                        lblinfo.ForeColor = Color.Blue;

                        txtTestResult.Focus();
                        txtTestResult.SelectAll();
                        btnsearch_Click(sender, e);
                    }
                }
                else
                {

                    lblinfo.Text = txtproductcode.Text + "该料号不存在";
                    lblinfo.ForeColor = Color.Red;
                    cbTestType.Properties.Items.Clear();
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
                txtproductcode_Leave( sender, e);
            }
        }



        private void cbTestType_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                DataRow[] rw = dtpub.Select("TestType='" + cbTestType.Text+ "'");
                if (rw.Length > 0)
                    txtdesc.Text = rw[0]["TestDesc"].ToString();
            }
            catch { }
        }

        private void btntestreport_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofdImport = new OpenFileDialog();
            ofdImport.Filter = "文件(*.jpg;bmp;png;jpeg;pdf)|*.jpg;*.bmp;*.png;*.jpeg;*.pdf";
            ofdImport.Multiselect = false;
            DialogResult dr = ofdImport.ShowDialog();
            if (dr == DialogResult.Cancel)
                return;
            this.txttestreport.Text = "";
            foreach (string str in ofdImport.FileNames)
            {
                this.txttestreport.Text += str + ",";
            }
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
            string name = "ROHS信息";
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
        private void OpenExcelFile(string fileName)
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
            //if (databind.Rows.Count < 1) return;
            //DataToExcel(databind);

            DataTable dt = databind.DataSource as DataTable;
            if (dt == null)
                return;
            if (dt.Rows.Count <= 0) return;

            //  DataToExcel(databind);

            string fileName = ShowSaveFileDialog("Microsoft Excel 2007 Document", "Microsoft Excel|*.xlsx");
            if (fileName == string.Empty) return;
            ExportToEx(fileName, "xlsx", gridView);
            OpenExcelFile(fileName);
        }

        private void btnsave_Click(object sender, EventArgs e)
        {
            if (txtproductcode.Text.Trim() == "")
                return;
            string msg = ic.AddNewTestROHS("新增", txtproductcode.Text, txtguigei.Text, txtsenttestdate.DateTime.ToString(), cbTestType.Text, txtsentremarks.Text, Login.username, 0, date.DateTime.ToString(), "", txtTestResult.Text, "");
            if (msg.IndexOf("OK") >= 0)
            {
                DataSet ds = ic.SelectTestROHS("查询", txtproductcode.Text, txtguigei.Text, txtsenttestdate.DateTime.ToString(), cbTestType.Text, txtTestResult.Text, Login.username, 0, date.DateTime.ToString(), "", txtTestResult.Text, "");
                databind.DataSource = ds.Tables[0];
                txtTestResult.Text = "";
                lblinfo.Text = msg;
                txtguigei.Text = "";
                txtproductcode.Focus();
                txtproductcode.SelectAll();
            }
            else
            {
                MessageBox.Show(msg, "提示");
            }
        }

        private void btnsearch_Click(object sender, EventArgs e)
        {
            databind.DataSource = null;
            DataSet ds = ic.SelectTestROHS("查询", txtproductcode.Text, txtguigei.Text, this.txtsenttestdate.DateTime.ToString(), cbTestType.Text.Trim() == "" ? "" : cbTestType.Text.Trim(), txtTestResult.Text, Login.username, 0, this.date.DateTime.ToString(), "", txtTestResult.Text, "");
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                databind.DataSource = ds.Tables[0];
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

        protected void del_prefile(string filepath)
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
                string floerPath = "\\\\" + this.serverFilePath + "\\" + folderType + "\\" + this.txtproductcode.Text.Trim();
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
            
            if (Connect(serverFilePath))
            {
             
                bool re = true;
                string msg = "";
                string te = txttestreport.Text.TrimEnd(',');
                if (te != "")
                {
                    string[] copy = te.Split(',');
                    foreach (string ss in copy)
                    {
                        if (!CopyFileToServer("测试报告", ss))
                        {
                            msg += ss + ";";
                            re = false;
                        }
                    }
                }
                if (re)
                {
                    MessageBox.Show("操作成功");
                    return true;
                }
                else
                {
                    MessageBox.Show(msg + " 文件上传失败");
                    return false;
                }
            }
            else
            {
                MessageBox.Show("无法连接到服务器的共享目录");
                return false;
            }

        }
        public void UpLoadFile(string fileNamePath, string uriString, bool IsAutoRename)
        {
            string fileName = fileNamePath.Substring(fileNamePath.LastIndexOf("\\") + 1);
            string NewFileName = fileName;
            if (IsAutoRename)
            {
                NewFileName = DateTime.Now.ToString("yyMMddhhmmss") + DateTime.Now.Millisecond.ToString() + fileNamePath.Substring(fileNamePath.LastIndexOf("."));
            }

            string fileNameExt = fileName.Substring(fileName.LastIndexOf(".") + 1);
            if (uriString.EndsWith("/") == false) uriString = uriString + "/";
            if (!Directory.Exists(uriString))//如果不存在就创建file文件夹
            {
                Directory.CreateDirectory(uriString);
            }
            uriString = uriString + NewFileName;
            /**/
            /// 创建WebClient实例  
            System.Net.WebClient myWebClient = new WebClient();

            myWebClient.Credentials = new NetworkCredential("10.1.31.218", "0.");

            // myWebClient.Credentials = CredentialCache.DefaultCredentials;
            // 要上传的文件  
            FileStream fs = new FileStream(fileNamePath, FileMode.Open, FileAccess.Read);
            BinaryReader r = new BinaryReader(fs);
            byte[] postArray = r.ReadBytes((int)fs.Length);

            Stream postStream = myWebClient.OpenWrite(uriString, "PUT");
            try
            {
                //使用UploadFile方法可以用下面的格式  
                if (postStream.CanWrite)
                {
                    postStream.Write(postArray, 0, postArray.Length);
                    postStream.Close();
                    fs.Dispose();
                    //  log.AddLog("上传日志文件成功！", "Log");
                    //  basicInfo.writeLogger("上传日志文件成功！" );
                }
                else
                {
                    postStream.Close();
                    fs.Dispose();
                }
            }
            catch (Exception err)
            {
                postStream.Close();
                fs.Dispose();
                throw err;
            }
            finally
            {
                postStream.Close();
                fs.Dispose();
            }
        }
        public int item = 0;
        private void btnupdate_Click(object sender, EventArgs e)
        {
            if (txtproductcode.Text == "") return;
            string sresult = "";
            if (cbOK.Checked)
                sresult = "OK";
            else if (cbNG.Checked)
                sresult = "NG";
            if (cbOK.Checked && cbNG.Checked)
            {
                MessageBox.Show("只能选择一种测试结果!");
                return;
            }
            if (!cbOK.Checked && !cbNG.Checked)
            {
                MessageBox.Show("只能选择一种测试结果!");
                return;
            }

            ////  \\10.100.0.150\rohs$\测试报告\5208002200641
            ////  50020100002359[18-05-18-09-59-25]   50020100002359[18-05-18-10-04-04]
            ///// 

            if (gridView.RowCount < 1)
                return;
            if (gridView.FocusedRowHandle < 0)
                return;

            try
            {
                string itemstr = gridView.GetFocusedRowCellValue("序号").ToString();
                item = int.Parse(itemstr);
            }
            catch
            {

            }

            string stringtime = DateTime.Now.ToString("yy-MM-dd-HH-mm-ss");
            string filename = this.txtproductcode.Text.Trim() + "["+stringtime+ "]";

            string fileDBServerPath = "";
                string floerfilePath = "";
                floerfilePath = "\\\\" + this.serverFilePath + "\\测试报告" + "\\" + this.txtproductcode.Text.Trim();
                if (txttestreport.Text != "" && (cbTestType.Text.Contains("可焊性") || cbTestType.Text.Contains("切片测试") || cbTestType.Text.Contains("可靠性测试")))
                {
                    string[] fileqty = txttestreport.Text.TrimEnd(',').Split(',');
                    foreach (string s in fileqty)
                    {
                        string[] fileName = @s.Split('\\');
                        fileDBServerPath += "\\\\" + this.serverFilePath + "\\测试报告" + "\\" + this.txtproductcode.Text.Trim() + "\\" + fileName[fileName.Length - 1] + ",";
                    }

                if (floerfilePath != "")
                  ////// del_prefile(floerfilePath); 

                if (UploadFile(fileDBServerPath) == true)
                  {
                    string newfilename = "\\\\" + this.serverFilePath + "\\测试报告" + "\\" + this.txtproductcode.Text.Trim() + "\\" + filename+".PDF";
                    File.Move(fileDBServerPath.TrimEnd(','), newfilename);
                    
                    txttestreport.Text = "";
                  }
               }                 
            //20151106新增自动上传测试报告
            FileInfo newestFile = null;
            if (cbTestType.Text.Contains("有害物质") || cbTestType.Text.Contains("可焊性") || cbTestType.Text.Contains("切片测试"))
            {
                string floerPath = "";
                floerPath = "\\\\" + this.serverFilePath + "\\测试报告" + "\\" + this.txtproductcode.Text.Trim();
                if (Connect(serverFilePath))
                {
                    if (Directory.Exists(floerPath) == false)
                    {
                        MessageBox.Show("还没有上传测试报告!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        return;
                    }
                    string[] file = Directory.GetFiles(floerPath);
                    if (file.Length >= 0)
                    {
                        DirectoryInfo info = new DirectoryInfo(floerPath);
                        FileInfo[] files = info.GetFiles();
                        DateTime time = new DateTime(1900, 1, 1);
                    
                        foreach (FileInfo f in files)
                        {
                            if (f.LastWriteTime > time && f.Extension.ToUpper() == ".PDF")
                            {
                                time = f.LastWriteTime;
                                newestFile = f;
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("没有测试报告文件", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                else
                {
                    MessageBox.Show("无法连接到服务器的共享目录");
                }
            }
            //20151106新增自动上传测试报告

            string msg = ic.AddNewTestROHS("更新", txtproductcode.Text, txtguigei.Text, txtsenttestdate.DateTime.ToString(), cbTestType.Text, txtsentremarks.Text, Login.username, item, date.DateTime.ToString(), sresult, txtTestResult.Text, newestFile == null ? "" : newestFile.ToString());
            if (msg.IndexOf("OK") >= 0)
            {
                DataSet ds = ic.SelectTestROHS("查询", txtproductcode.Text, txtguigei.Text, txtsenttestdate.DateTime.ToString(), cbTestType.Text, txtTestResult.Text, Login.username, 0, date.DateTime.ToString(), sresult, txtTestResult.Text, "");

                databind.DataSource = ds.Tables[0];
                txtproductcode.Text = "";
                txtTestResult.Text = "";
                lblinfo.Text = "";
                txtguigei.Text = "";
                txtproductcode.Focus();
                item = 0;
            }
            else
            {
                MessageBox.Show(msg, "提示");
            }
        }

        private void btndel_Click(object sender, EventArgs e)
        {
            ////if (databind.SelectedRows.Count < 1) return;

            if (gridView.FocusedRowHandle < 0)
                return;
            if (gridView.GetSelectedRows().Length < 1)
                return;

            for (int i = gridView.GetSelectedRows().Length; i > 0; i--)
            {
                DataRow dr = gridView.GetDataRow(gridView.GetSelectedRows()[i - 1]);
                string msg = ic.AddNewTestROHS("删除", dr["物料编码"].ToString(), "", "", "", "", Login.username, int.Parse(dr["序号"].ToString()), "", "", "", "");
                if (msg.IndexOf("OK") >= 0)
                {
                    //this.databind.Rows.RemoveAt(databind.SelectedRows[i - 1].Index);
                    gridView.DeleteRow(gridView.GetSelectedRows()[i - 1]);

                }
            }
        }

        private void btnsearchall_Click(object sender, EventArgs e)
        {
            databind.DataSource = null;
            DataSet ds = ic.SelectTestROHS("所有查询", txtproductcode.Text, txtguigei.Text, this.txtsenttestdate.DateTime.ToString(), cbTestType.Text.Trim() == "" ? "" : cbTestType.Text.Trim(), txtTestResult.Text, Login.username, 0, this.date.DateTime.ToString(), "", txtTestResult.Text, "");
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                databind.DataSource = ds.Tables[0];
        }

        private void databind_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //try
            //{
            //    item = int.Parse(databind.CurrentRow.Cells["序号"].Value.ToString());
            //    txtproductcode.Text = databind.CurrentRow.Cells["物料编码"].Value.ToString();
            //    bindReliabilityType(txtproductcode.Text);
            //    cbTestType.SelectedValue = databind.CurrentRow.Cells["测试项"].Value.ToString();
            //    txtdesc.Text = databind.CurrentRow.Cells["测试内容"].Value.ToString();
            //    txtsenttestdate.Value = (DateTime)databind.CurrentRow.Cells["送测日期"].Value;
            //    txtguigei.Text = databind.CurrentRow.Cells["样品规格书编号"].Value.ToString();
            //    txtsentremarks.Text = databind.CurrentRow.Cells["送测说明"].Value.ToString();
            //    txtTestResult.Text = databind.CurrentRow.Cells["测试备注"].Value.ToString();
            //    if (databind.CurrentRow.Cells["测试结果"].Value.ToString() == "OK")
            //        cbOK.Checked = true;
            //    else if (databind.CurrentRow.Cells["测试结果"].Value.ToString() == "NG")
            //        cbNG.Checked = true;
            //}
            //catch { }
        }
        private void OpenFile(string sfile)
        {
            string filename = "";
            string Att = sfile;
            if (Att != "")
            {
                //定义一个ProcessStartInfo实例
                System.Diagnostics.ProcessStartInfo info = new System.Diagnostics.ProcessStartInfo();
                //设置启动进程的初始目录
                string[] fileName = @Att.Split('\\');

                filename = fileName[fileName.Length - 1];
                info.FileName = @filename;
                info.WorkingDirectory = Att.Substring(0, Att.LastIndexOf('\\'));
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

        }
        private void databind_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            DataGridView dg = (DataGridView)sender;
            if (dg.Columns[e.ColumnIndex].Name == "测试报告" && dg.CurrentRow.Cells["测试报告"].Value.ToString() != "")
            {
                string Att = dg.CurrentRow.Cells["物料编码"].Value.ToString().Replace(" ", "");
                string floerPath = "\\\\" + this.serverFilePath + "\\测试报告" + "\\" + Att;
                if (Connect(serverFilePath))
                {
                    if (Directory.Exists(floerPath) == false)
                    {
                        MessageBox.Show("还没有上传测试报告!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        return;
                    }
                    //string[] str = Directory.GetFiles(floerPath);
                    string[] str = Directory.GetFiles(floerPath);
                    if (str.Length > 0)
                    {
                        foreach (string s in str)
                        {
                            string[] fileName = s.Split('\\');
                            if (fileName[fileName.Length - 1] == dg.CurrentRow.Cells["测试报告"].Value.ToString() && fileName[fileName.Length - 1].ToUpper().EndsWith(".PDF"))
                            {
                                OpenFile(s);
                                break;
                            }
                        }

                    }
                }
            }
        }

        private void databind_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            //if (databind.Rows.Count <= 0) return;
            //DataGridViewRow dgr = databind.Rows[e.RowIndex];
            //if (dgr.Cells["测试项"].Value.ToString() == "")
            //{
            //    dgr.DefaultCellStyle.BackColor = Color.Red;
            //}
            //else if (int.Parse(dgr.Cells["已送测天数"].Value.ToString()) >= 7 && dgr.Cells["测试日期"].Value.ToString() == "")
            //{
            //    dgr.DefaultCellStyle.BackColor = Color.YellowGreen;
            //}
            //else if (int.Parse(dgr.Cells["距离天数"].Value.ToString()) <= int.Parse(dgr.Cells["提前期"].Value.ToString()) /*&& dgr.Cells["测试日期"].Value.ToString() == ""*/)
            //{
            //    dgr.DefaultCellStyle.BackColor = Color.Yellow;

            //}
            //else if (dgr.Cells["测试结果"].Value.ToString() == "NG")
            //{
            //    dgr.DefaultCellStyle.BackColor = Color.Red;

            //}
        }

        private void gridView_Click(object sender, EventArgs e)
        {
            //try
            //{
            //    item = int.Parse(databind.CurrentRow.Cells["序号"].Value.ToString());
            //    txtproductcode.Text = databind.CurrentRow.Cells["物料编码"].Value.ToString();
            //    bindReliabilityType(txtproductcode.Text);
            //    cbTestType.SelectedValue = databind.CurrentRow.Cells["测试项"].Value.ToString();
            //    txtdesc.Text = databind.CurrentRow.Cells["测试内容"].Value.ToString();
            //    txtsenttestdate.Value = (DateTime)databind.CurrentRow.Cells["送测日期"].Value;
            //    txtguigei.Text = databind.CurrentRow.Cells["样品规格书编号"].Value.ToString();
            //    txtsentremarks.Text = databind.CurrentRow.Cells["送测说明"].Value.ToString();
            //    txtTestResult.Text = databind.CurrentRow.Cells["测试备注"].Value.ToString();
            //    if (databind.CurrentRow.Cells["测试结果"].Value.ToString() == "OK")
            //        cbOK.Checked = true;
            //    else if (databind.CurrentRow.Cells["测试结果"].Value.ToString() == "NG")
            //        cbNG.Checked = true;
            //}
            //catch { }

            try
            {
                item = int.Parse(gridView.GetFocusedRowCellValue("序号").ToString());
                txtproductcode.Text = gridView.GetFocusedRowCellValue("物料编码").ToString();
                bindReliabilityType(txtproductcode.Text);
                cbTestType.Text = gridView.GetFocusedRowCellValue("测试项").ToString();
                txtdesc.Text = gridView.GetFocusedRowCellValue("测试内容").ToString();
                txtsenttestdate.EditValue = (DateTime)gridView.GetFocusedRowCellValue("送测日期");
                txtguigei.Text = gridView.GetFocusedRowCellValue("样品规格书编号").ToString();
                txtsentremarks.Text = gridView.GetFocusedRowCellValue("送测说明").ToString();
                txtTestResult.Text = gridView.GetFocusedRowCellValue("测试备注").ToString();
                if (gridView.GetFocusedRowCellValue("测试结果").ToString() == "OK")
                    cbOK.Checked = true;
                else if (gridView.GetFocusedRowCellValue("测试结果").ToString() == "NG")
                    cbNG.Checked = true;
            }
            catch { }

        }

        private void gridView_CustomDrawCell(object sender, RowCellCustomDrawEventArgs e)
        {
           // if (databind.Rows.Count <= 0) return;
            //DataGridViewRow dgr = databind.Rows[e.RowIndex];
            //if (dgr.Cells["测试项"].Value.ToString() == "")
            //{
            //    dgr.DefaultCellStyle.BackColor = Color.Red;
            //}
            //else if (int.Parse(dgr.Cells["已送测天数"].Value.ToString()) >= 7 && dgr.Cells["测试日期"].Value.ToString() == "")
            //{
            //    dgr.DefaultCellStyle.BackColor = Color.YellowGreen;
            //}
            //else if (int.Parse(dgr.Cells["距离天数"].Value.ToString()) <= int.Parse(dgr.Cells["提前期"].Value.ToString()) /*&& dgr.Cells["测试日期"].Value.ToString() == ""*/)
            //{
            //    dgr.DefaultCellStyle.BackColor = Color.Yellow;

            //}
            //else if (dgr.Cells["测试结果"].Value.ToString() == "NG")
            //{
            //    dgr.DefaultCellStyle.BackColor = Color.Red;

            //}            
            ////////DataTable dt = databind.DataSource as DataTable;
            ////////if (dt == null || dt.Rows.Count < 0)
            ////////{
            ////////    return;
            ////////}

            ////////if (e.Column.FieldName == "测试项")
            ////////{
            ////////    GridCellInfo GridCell = e.Cell as GridCellInfo;
            ////////    if (GridCell.CellValue.ToString() == "")
            ////////    {
            ////////        e.Appearance.BackColor = Color.Red;
            ////////    }
            ////////}
            ////////if (int.Parse (gridView.GetDataRow(e.RowHandle)["已送测天数"].ToString()) >= 7 && gridView.GetDataRow(e.RowHandle)["测试日期"].ToString()=="")
            ////////{
            ////////  e.Appearance.BackColor = Color.YellowGreen;               
            ////////}
            ////////if (int.Parse(gridView.GetDataRow(e.RowHandle)["距离天数"].ToString()) <= int.Parse(gridView.GetDataRow(e.RowHandle)["提前期"].ToString()))
            ////////{
            ////////  e.Appearance.BackColor = Color.Yellow;                
            ////////}
            ////////if (e.Column.FieldName == "测试结果")
            ////////{
            ////////    GridCellInfo GridCell = e.Cell as GridCellInfo;
            ////////    if (GridCell.CellValue.ToString() == "NG")
            ////////    {
            ////////        e.Appearance.BackColor = Color.Red;
            ////////    }
            ////////}

        }

        private void gridView_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
        {

            DataTable dt = databind.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
            {
                return;
            }
            if (gridView.GetDataRow(e.RowHandle)["测试项"].ToString() == "")
            {
                e.Appearance.BackColor = Color.Red;
            }
             if (int.Parse(gridView.GetDataRow(e.RowHandle)["已送测天数"].ToString()) >= 7 && gridView.GetDataRow(e.RowHandle)["测试日期"].ToString() == "")
            {
                e.Appearance.BackColor = Color.YellowGreen;
            }
            if (int.Parse(gridView.GetDataRow(e.RowHandle)["距离天数"].ToString()) <= int.Parse(gridView.GetDataRow(e.RowHandle)["提前期"].ToString()))
            {
                e.Appearance.BackColor = Color.Yellow;
            }
            if (gridView.GetDataRow(e.RowHandle)["测试结果"].ToString() == "NG")
            {
                e.Appearance.BackColor = Color.Red;
            }

        }

        private void gridView_DoubleClick(object sender, EventArgs e)
        {
            ////  DataGridView dg = (DataGridView)sender;
            DataTable dt = databind.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
                return;
            if (gridView.FocusedRowHandle < 0)
                return;

            // gridView.GetFocusedRowCellValue("测试报告").ToString()

            if (gridView.GetFocusedRowCellValue("测试报告").ToString() != "")
            {
                string Att = gridView.GetFocusedRowCellValue("物料编码").ToString().Replace(" ", "");
                string floerPath = "\\\\" + this.serverFilePath + "\\测试报告" + "\\" + Att;
                if (Connect(serverFilePath))
                {
                    if (Directory.Exists(floerPath) == false)
                    {
                        MessageBox.Show("还没有上传测试报告!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        return;
                    }
                    //string[] str = Directory.GetFiles(floerPath);
                    string[] str = Directory.GetFiles(floerPath);
                    if (str.Length > 0)
                    {
                        foreach (string s in str)
                        {
                            string[] fileName = s.Split('\\');
                            if (fileName[fileName.Length - 1] == gridView.GetFocusedRowCellValue("测试报告").ToString() && fileName[fileName.Length - 1].ToUpper().EndsWith(".PDF"))
                            {
                                OpenFile(s);
                                break;
                            }
                        }
                    }
                }
            }

        }

        private void gridView_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {

            if (gridView.RowCount < 1)
                return;
            if (gridView.FocusedRowHandle < 0)
                return;
            if (e.Column.FieldName == "物料编码")
            {
                if (Login.userId == "160819116" || Login.userId == "12029040" || Login.userId == "130610551" || Login.userId == "admin")
                {
               
                    string Att = gridView.GetFocusedRowCellValue("物料编码").ToString().Replace(" ", "");
                    string floerPath = "\\\\" + this.serverFilePath + "\\测试报告" + "\\" + Att;
                    try
                    {
                        if (Connect(serverFilePath))
                        {
                            Cursor.Current = Cursors.WaitCursor;
                            Process.Start(floerPath);
                            Cursor.Current = Cursors.Default;
                        }
                        else
                        {
                            MessageBox.Show("无法连接到服务器的共享目录");
                        }
                    }
                    catch (Exception ex)
                    {
                        if (ex.ToString().Contains("系统找不到指定的文件"))
                        {
                            MessageBox.Show("系统找不到指定的文件夹", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }                  
                }
            }

         }




    }
}