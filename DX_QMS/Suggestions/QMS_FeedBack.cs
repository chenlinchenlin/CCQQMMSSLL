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
using System.IO;
using DX_QMS.Common;
using System.Diagnostics;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraEditors;

namespace DX_QMS.Suggestions
{
    public partial class QMS_FeedBack : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private string serverFilePath = "192.168.0.204\\FilePath$";
        public QMS_FeedBack()
        {
            InitializeComponent();
        }

        private void QMS_FeedBack_Load(object sender, EventArgs e)
        {
            datebegin.EditValue = DateTime.Now.ToLocalTime().ToString();
            dateend.EditValue = DateTime.Now.ToLocalTime().ToString();
            gridView.PrintExportProgress += new ProgressChangedEventHandler(gridView_PrintExportProgress);
            string sql = "select  substring(replace(replace(replace(replace(convert(varchar(30),getdate(),121),'-',''),' ',''),':',''),'.',''),1,12)";
            string item = DbAccess.SelectBySql(sql).Tables[0].Rows[0][0].ToString();
            txtproblemitme.Text = item;
            setRule();

        }

        private void setRule()
        {
            string post = "";
            if (Login.manager == "IT管理员")
            {
                Btnhandle.Enabled = true;
                Btnupdate.Enabled = true;
            }
            else
            {
                Btnhandle.Enabled = false;
                Btnupdate.Enabled = false;
            }
           // Dictionary<string, bool> dic = GroupPermission.QMS_SelectRulesForForm(post, "建议与反馈");
            //Btnhandle.Enabled = bool.Parse(dic["hasInsert"].ToString());
            //Btnupdate.Enabled = bool.Parse(dic["hasDelete"].ToString());
            //this.btntoexcel.Enabled = bool.Parse(dic["hasPrint"].ToString());
        }






        //private void Btnbrowseproblem_Click(object sender, EventArgs e)
        //{
        //    OpenFileDialog ofdImport = new OpenFileDialog();

        //    ofdImport.Filter = "文件(*.jpg;bmp;png;jpeg;pdf;docx;xlsx;xls)|*.jpg;*.bmp;*.png;*.jpeg;*.pdf;*.docx;*.xlsx;*.xlsx";
        //    ofdImport.Multiselect = false;
        //    DialogResult dr = ofdImport.ShowDialog();
        //    if (dr == DialogResult.Cancel) return;
        //    this.txtproblemfile.Text = "";
        //    foreach (string str in ofdImport.FileNames)
        //    {
        //        this.txtproblemfile.Text += str ;            
        //    }
        //}

        private void Btnbrowsedemand_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofdImport = new OpenFileDialog();

            ofdImport.Filter = "文件(*.jpg;bmp;png;jpeg;pdf;docx;xlsx;xls)|*.jpg;*.bmp;*.png;*.jpeg;*.pdf;*.docx;*.xlsx;*.xlsx";
            ofdImport.Multiselect = false;
            DialogResult dr = ofdImport.ShowDialog();
            if (dr == DialogResult.Cancel) return;
            this.txtdemandfile.Text = "";
            foreach (string str in ofdImport.FileNames)
            {
                this.txtdemandfile.Text += str ;
            }
        }


        public string AddNewQMSFeedBack(string opertype,  string  problemItems , string  moduleLocation , string  emergencyLevel, string  problemType,
                           string contact, string officeLocation,string  problemDescribe, string demandDescribe , string  feedbackMan ,  string  manageMan, string state,string remark)
        {
            SqlParameter[] para = new SqlParameter[14];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@problemItems", problemItems);
            para[2] = new SqlParameter("@moduleLocation", moduleLocation);
            para[3] = new SqlParameter("@emergencyLevel", emergencyLevel);
            para[4] = new SqlParameter("@problemType", problemType);
            para[5] = new SqlParameter("@contact", contact);
            para[6] = new SqlParameter("@officeLocation", officeLocation);
            para[7] = new SqlParameter("@problemDescribe", problemDescribe);
            para[8] = new SqlParameter("@demandDescribe", demandDescribe);
            para[9] = new SqlParameter("@feedbackMan", feedbackMan);
            para[10] = new SqlParameter("@manageMan", manageMan);
            para[11] = new SqlParameter("@state", state);
            para[12] = new SqlParameter("@remark", remark);
            para[13] = new SqlParameter("@msg", SqlDbType.VarChar, 50);
            para[13].Direction = ParameterDirection.Output;
            DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "QMS_OperFeedBack", para);
            return para[13].Value.ToString();
        }
        public DataSet SelectQMSFeedBack(string opertype, string problemItems, string moduleLocation, string emergencyLevel, string problemType,
                          string contact, string officeLocation, string problemDescribe, string demandDescribe, string feedbackMan, string manageMan, string state, string remark)
        {
            SqlParameter[] para = new SqlParameter[14];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@problemItems", problemItems);
            para[2] = new SqlParameter("@moduleLocation", moduleLocation);
            para[3] = new SqlParameter("@emergencyLevel", emergencyLevel);
            para[4] = new SqlParameter("@problemType", problemType);
            para[5] = new SqlParameter("@contact", contact);
            para[6] = new SqlParameter("@officeLocation", officeLocation);
            para[7] = new SqlParameter("@problemDescribe", problemDescribe);
            para[8] = new SqlParameter("@demandDescribe", demandDescribe);
            para[9] = new SqlParameter("@feedbackMan", feedbackMan);
            para[10] = new SqlParameter("@manageMan", manageMan);
            para[11] = new SqlParameter("@state", state);
            para[12] = new SqlParameter("@remark", remark);
            para[13] = new SqlParameter("@msg", SqlDbType.VarChar, 50);
            para[13].Direction = ParameterDirection.Output;
            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "QMS_OperFeedBack", para);
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
                string floerPath = "\\\\" + this.serverFilePath + "\\" + folderType + "\\" + txtproblemitme.Text.Trim();
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

        private bool UploadFile(string filepath,string file)
        {
            bool b = false;
            if (Connect(serverFilePath))
            {
                bool re = true;
                string msg = "";
                string te = file.TrimEnd(',');
                if (te != "")
                {
                    string[] copy = te.Split(',');
                    foreach (string ss in copy)
                    {
                        if (!CopyFileToServer("用户反馈与建议（QMS）", ss))
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

        private void Btnsave_Click(object sender, EventArgs e)
        {
            if (txtdemandfile.Text.Trim() == "")
            {
                MessageBox.Show("请上传需求描述文件","提醒",MessageBoxButtons.OK ,MessageBoxIcon.Information);
                return;
            }

           string problemItems,moduleLocation, emergencyLevel,problemType,contact,officeLocation, 
                  problemDescribe,demandDescribe,feedbackMan;

            problemItems = txtproblemitme.Text.Trim();
            moduleLocation = txtmodulelocation.Text;
            emergencyLevel = txtemergencylevel.Text;
            problemType = txtproblemtype.Text;
            contact = txtcontact.Text;
            officeLocation = txtofficelocation.Text;
            problemDescribe = txtproblemdescribe.Text;
            demandDescribe = txtdemanddescribe.Text;
            feedbackMan = Login.username;
            if (moduleLocation == "" || officeLocation == "")
            {
                MessageBox.Show("模块位置和办公位置不能为空", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string fileDBServerPathproblem = "";
            string fileDBServerPathdemand = "";
            if (txtdemandfile.Text != "")
            {
                string[] fileqty = txtdemandfile.Text.TrimEnd(',').Split(',');
                foreach (string s in fileqty)
                {
                    string[] fileName = @s.Split('\\');

                    string filename = Path.GetFileNameWithoutExtension(txtdemandfile.Text);

                    if (txtproblemitme.Text != filename)
                    {
                        MessageBox.Show("上传的文件名:" + fileName[fileName.Length - 1].Trim() + "与问题流水号:" + txtproblemitme.Text + "不一致");
                        return;
                    }

                    fileDBServerPathdemand += "\\\\" + this.serverFilePath + "\\用户反馈与建议（QMS）" + "\\" + txtproblemitme.Text.Trim() + "\\" + fileName[fileName.Length - 1] + ",";
                }
            }
            string floerPath = "";
            floerPath = "\\\\" + this.serverFilePath + "\\用户反馈与建议（QMS）" + "\\" + txtproblemitme.Text.Trim();
            if (floerPath != "")
                del_prefile(floerPath, "");

           // bool a = UploadFile(fileDBServerPathproblem, txtproblemfile.Text);
            bool b = UploadFile(fileDBServerPathdemand, txtdemandfile.Text);

            if (b)

            {

                string msg = AddNewQMSFeedBack("保存", problemItems, moduleLocation, emergencyLevel, problemType,
                               contact, officeLocation, problemDescribe, demandDescribe, feedbackMan, "", "", "");

                if (msg.Contains("成功"))
                {
                    MessageBox.Show(msg, "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                   // return;
                }
                else
                {
                    MessageBox.Show("保存失败", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

            }
        }

        private void Btnreset_Click(object sender, EventArgs e)
        {
            txtmodulelocation.Text = "";
            txtcontact.Text = "";
            txtofficelocation.Text = "";
            txtproblemitme.Text = "";
            txtemergencylevel.Text = "";
            txtproblemtype.Text = "";
            txtproblemdescribe.Text = "";
            //txtproblemfile.Text = "";
            txtdemanddescribe.Text = "";
            txtdemandfile.Text = "";
        }

        private void Btnsubmit_Click(object sender, EventArgs e)
        {
           string problemItems = txtproblemitme.Text.Trim();
            if ( problemItems =="" )
            {
                MessageBox.Show("问题流水号不能为空", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
           string msg = AddNewQMSFeedBack("提交", problemItems, "", "", "","", "", "", "", "", "", "", "");
            if (msg.Contains("成功"))
            {
                MessageBox.Show(msg, "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("提交失败", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }

        private void Btnselect_Click(object sender, EventArgs e)
        {
            string where = " where 1=1 ";
            string feedbackMan = txtfeedbackMan.Text.Trim();
            string emergencyLevel = comemergencylevel.Text.Trim();
            string problemType = comproblemtype.Text.Trim();
            string manageMan = txtmanageMan.Text.Trim();
            string state = txtstate.Text.Trim();
            string begintime = datebegin.Text;
            string endtime = dateend.Text;

            if (!string.IsNullOrEmpty(feedbackMan))
            {
                where += " and feedbackMan like '%" + feedbackMan + "%' ";
            }
            if (!string.IsNullOrEmpty(emergencyLevel))
            {
                where += " and emergencyLevel = '" + emergencyLevel + "' ";
            }
            if (!string.IsNullOrEmpty(problemType))
            {
                where += " and problemType = '" + problemType + "' ";
            }
            if (!string.IsNullOrEmpty(manageMan))
            {
                where += " and manageMan like '%" + manageMan + "%' ";
            }
            if (!string.IsNullOrEmpty(state))
            {
                where += " and state = '" + state + "' ";
            }
            if (!string.IsNullOrEmpty(begintime))
            {
                where += " and submitTime >= '" + begintime + " '";
            }
            if (!string.IsNullOrEmpty(endtime))
            {
                where += " and submitTime <= '" + endtime + " '";
            }

            string sql = @"  select problemItems 问题流水号,moduleLocation 模块位置,emergencyLevel 紧急程度,problemType 问题分类,contact 联系方式,officeLocation 办公位置,
		       problemDescribe 问题描述,demandDescribe 需求描述,feedbackMan 反馈人,manageMan 受理人,ifSubmit 是否提交,submitTime 提交时间,state 状态,handleTime 处理时间,remark 备注 from QMS_FeedBack  ";
            sql += where + " order by submitTime desc ";

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

        private void Btnhandle_Click(object sender, EventArgs e)
        {
            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
            {
                return;
            }
            if (gridView.FocusedRowHandle < 0)
            {              
                MessageBox.Show("请选择需要受理的项目", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            string problemitem = gridView.GetFocusedRowCellValue("问题流水号").ToString();
            string state = txtstate.Text.Trim();
            string manageMan = txtmanageMan.Text.Trim();
            if (state == "" || manageMan =="")
            {
                MessageBox.Show("状态和受理人不能为空", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            string msg = AddNewQMSFeedBack("受理", problemitem,"","", "","","","","","",manageMan, state, txtremark.Text);
            if (msg.Contains("成功"))
            {
                MessageBox.Show(msg, "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Btnselect_Click(null , null);
            }
            else
            {
                MessageBox.Show("受理失败", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }

        private void Btnupdate_Click(object sender, EventArgs e)
        {


            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
            {
                return;
            }
            if (gridView.FocusedRowHandle < 0)
            {
                MessageBox.Show("请选择需要更新的项目", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            string problemitem = gridView.GetFocusedRowCellValue("问题流水号").ToString();
            string state = txtstate.Text.Trim();
            string manageMan = txtmanageMan.Text.Trim();
            if (state == "" || manageMan == "")
            {
                MessageBox.Show("状态和受理人不能为空", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            string msg = AddNewQMSFeedBack("更新", problemitem, "", "", "",
                           "", "", "", "", "", manageMan, state, txtremark.Text);
            if (msg.Contains("成功"))
            {
                MessageBox.Show(msg, "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Btnselect_Click(null, null);
            }
            else
            {
                MessageBox.Show("更新失败", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }


        private string ShowSaveFileDialog(string title, string filter)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            string name = "反馈信息";
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
        private void btntoexcel_Click(object sender, EventArgs e)
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

        private void Btnclear_Click(object sender, EventArgs e)
        {
            txtfeedbackMan.Text = "";
            comemergencylevel.Text = "";
            comproblemtype.Text = "";
            txtmanageMan.Text = "";
            txtstate.Text = "";
            txtremark.Text = "";
            datebegin.Text = "";
            dateend.Text = "";
            gridControl.DataSource = null;

        }

        private void gridView_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {
            try
            {
                DataTable dt = gridControl.DataSource as DataTable;
                if (dt == null || dt.Rows.Count < 1)
                {
                    return;
                }
                txtfeedbackMan.Text = gridView.GetFocusedRowCellValue("反馈人").ToString();
                comemergencylevel.Text = gridView.GetFocusedRowCellValue("紧急程度").ToString();
                comproblemtype.Text = gridView.GetFocusedRowCellValue("问题分类").ToString();
                txtmanageMan.Text = gridView.GetFocusedRowCellValue("受理人").ToString();
                txtstate.Text = gridView.GetFocusedRowCellValue("状态").ToString();

            }
            catch
            {
            }


        }

        private void OpenFile(string filepath, string pdffile)
        {
            string filename = "";
            filename = pdffile + ".pdf";
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

            }

            catch (System.ComponentModel.Win32Exception we)
            {

                MessageBox.Show(this, we.Message);
                return;
            }

        }


        private void gridView_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {

        }

        private void gridView_DoubleClick(object sender, EventArgs e)
        {

            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
            {
                return;
            }

            if (gridView.FocusedRowHandle < 0)
                return;
            if (gridView.GetSelectedRows().Length != 1)
                return;


            string problemitem = gridView.GetFocusedRowCellValue("问题流水号").ToString();
            string floerPath = "\\\\" + this.serverFilePath + "\\用户反馈与建议（QMS）" + "\\" + problemitem;
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
                            string filename = Path.GetFileNameWithoutExtension(pt[i].ToString());
                           // string filename = Path.GetFileName(pt[i].ToString());


                            if (filename == problemitem)
                            {
                               // FileStream fs = new FileStream(pt[i].ToString(), FileMode.Open);
                                FileStream fs = new FileStream(pt[i].ToString(), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
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
                                    OpenFile(floerPath, filename);
                                    fs.Close();
                                    fs.Dispose();
                                    reader.Close();
                                    return;
                                }

                            }
                        }
                    }
                }
                else
                {
                    string pnoold = "";
                    string sqlold = "  select problemItems  from  QMS_FeedBack where problemItems = '"+ problemitem + "'";
                    DataTable dtold = DbAccess.SelectBySql(sqlold).Tables[0];
                    if (dtold.Rows.Count > 0)
                    {
                        pnoold = dtold.Rows[0]["problemItems"].ToString();
                        string floerPathold = "\\\\" + this.serverFilePath + "\\用户反馈与建议（QMS）" + "\\" + pnoold;
                        if (Directory.Exists(floerPathold))
                        {
                            string[] pt = Directory.GetFiles(floerPathold);
                            if (pt.Length == 0) return;
                            if (pt.Length > 0)
                            {
                                for (int i = 0; i < pt.Length; i++)
                                {
                                    if (pt[i].ToString().Contains(".db")) continue;
                                    string filename = Path.GetFileNameWithoutExtension(pt[i].ToString());
                                   // string filename = Path.GetFileName(pt[i].ToString());

                                    if (filename == pnoold)
                                    {
                                       // FileStream fs = new FileStream(pt[i].ToString(), FileMode.Open);
                                        FileStream fs = new FileStream(pt[i].ToString(), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
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
                                            OpenFile(floerPath, filename);
                                            fs.Close();
                                            fs.Dispose();
                                            reader.Close();
                                            return;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }




        }
    }
}