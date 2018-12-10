using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Windows.Forms;
using DevExpress.XtraBars;
using System.Diagnostics;
using System.IO;
using DX_QMS.Common;
using DevExpress.XtraTreeList.Nodes;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraEditors;

namespace DX_QMS.IQCFilePosition
{
    public partial class IQCUploadFile : DevExpress.XtraBars.Ribbon.RibbonForm
    {

        private string serverFilePath = "192.168.0.204\\FilePath$";
        bool loadDrives = false;
        public IQCUploadFile()
        {
            InitializeComponent();
            InitData();
            gridViewproxy.PrintExportProgress += new ProgressChangedEventHandler(gridView_PrintExportProgress);
        }
        private void IQCUploadFile_Load(object sender, EventArgs e)
        {
            if (Login.userId == "150915722" || Login.manager == "IT管理员")
            {
                sBtnsavepaper.Enabled = true;
                sBtndeletepaper.Enabled = true;
            }
            else
            {
                sBtnsavepaper.Enabled = false;
                sBtndeletepaper.Enabled = false;
            }
            if (Login.post == "SQE" || Login.manager == "IQC管理员" || Login.manager == "IT管理员")
            {
                sBtnsavaproxy.Enabled = true;
                sBtndeleteproxy.Enabled = true;
            }
            else
            {
                sBtnsavaproxy.Enabled = false;
                sBtndeleteproxy.Enabled = false;
            }
        }

        void InitData()
        {
            treeListpaper.DataSource = new object();
        }


        private void treeListpaper_VirtualTreeGetChildNodes(object sender, DevExpress.XtraTreeList.VirtualTreeGetChildNodesInfo e)
        {
            Cursor current = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;
            if (!loadDrives)
            {
                if (Connect(serverFilePath))
                {
                    string[] roots;
                    if (txtpapermaterialcode.Text.Trim() == "")
                    {
                        roots = new string[1] { "\\\\192.168.0.204\\FilePath$\\图纸归档库 " };    //////  \制程异常处理报告\18045079-W
                    }
                    else
                    {
                        roots = new string[1] { "\\\\192.168.0.204\\FilePath$\\图纸归档库\\" + txtpapermaterialcode.Text};
                    }
                                    
                    e.Children = roots;
                    loadDrives = true;
                    lblmessage.Text = "";
                }
                else
                {
                    //MessageBox.Show("无法连接到服务器的共享目录");
                    lblmessage.Text = "无法连接到服务器的共享目录";
                }

            }
            else
            {
                try
                {
                    string path = (string)e.Node;
                    if (Directory.Exists(path))
                    {
                        string[] dirs = Directory.GetDirectories(path);
                        string[] files = Directory.GetFiles(path);
                        string[] arr = new string[dirs.Length + files.Length];
                        dirs.CopyTo(arr, 0);
                        files.CopyTo(arr, dirs.Length);
                        e.Children = arr;
                    }
                    else e.Children = new object[] { };
                }
                catch
                {
                    e.Children = new object[] { };

                }
            }
                Cursor.Current = current;
        }
        private void treeListpaper_CustomDrawNodeCell(object sender, DevExpress.XtraTreeList.CustomDrawNodeCellEventArgs e)
        {
            if (e.Column == this.colSize)
            {
                if (e.Node.GetDisplayText("Type") == "File")
                {
                    e.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
                    e.Appearance.Font = new Font(e.Appearance.Font, FontStyle.Italic);
                    Int64 size = Convert.ToInt64(e.Node.GetValue("Size"));
                    if (size >= 1024)
                        e.CellText = string.Format("{0:### ### ###} KB", size / 1024);
                    else e.CellText = string.Format("{0} Bytes", size);
                }
                else e.CellText = String.Format("<{0}>", e.Node.GetDisplayText("Type"));
            }

            if (e.Column == this.colName)
            {
                if (e.Node.GetDisplayText("Type") == "File")
                {
                    e.Appearance.Font = new Font(e.Appearance.Font, FontStyle.Bold);
                }
            }

        }
        private void treeListpaper_VirtualTreeGetCellValue(object sender, DevExpress.XtraTreeList.VirtualTreeGetCellValueInfo e)
        {
            DirectoryInfo di = new DirectoryInfo((string)e.Node);
            if (e.Column == colName)
                e.CellData = di.Name;
            if (e.Column == colType)
            {
                if (IsDrive((string)e.Node))
                    e.CellData = "Drive";
                else if (!IsFile(di))
                    e.CellData = "Folder";
                else
                    e.CellData = "File";
            }
            if (e.Column == colSize)
            {
                if (IsFile(di))
                {
                    e.CellData = new FileInfo((string)e.Node).Length;
                }
                else e.CellData = null;
            }
        }


        private void txtExpiryDate_EditValueChanged(object sender, EventArgs e)
        {
            //if (((DateTime)txtExpiryDate.EditValue).ToString("yyyy-MM-dd") == "0001-01-01")
            //{
            //    txtExpiryDate.EditValue = "1900-01-01";
            //}

            if (txtExpiryDate.Text == "")
            {
                txtExpiryDate.EditValue = "1900-01-01";
            }
        }

        bool IsFile(DirectoryInfo info)
        {
            try
            {
                return (info.Attributes & FileAttributes.Directory) == 0;
            }
            catch
            {
                return false;
            }
        }
        bool IsDrive(string val)
        {
            string[] drives = Directory.GetLogicalDrives();
            foreach (string drive in drives)
            {
                if (drive.Equals(val)) return true;
            }
            return false;
        }
        //</treeList1>

        private void treeList_GetStateImage(object sender, DevExpress.XtraTreeList.GetStateImageEventArgs e)
        {
            if (e.Node.GetDisplayText("Type") == "Folder")
                e.NodeImageIndex = e.Node.Expanded ? 1 : 0;
            else if (e.Node.GetDisplayText("Type") == "File")
                e.NodeImageIndex = 2;
            else e.NodeImageIndex = 3;
        }


        private string FullNameByNode(TreeListNode node, int columnId)
        {
            string ret = node.GetValue(columnId).ToString();
            while (node.ParentNode != null)
            {
                node = node.ParentNode;
                ret = node.GetValue(columnId).ToString() + "\\" + ret;
            }
            return ret;
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

        private void treeListpaper_DoubleClick(object sender, EventArgs e)
        {
            string papermaterialcode = "", filefix = "";
            string filename = treeListpaper.FocusedNode.GetValue("Name").ToString();
            filefix = GetPostfixStr(filename);

            if (filefix == "")
                return;

            if (filefix.Contains (".db"))
            {
                return;
            }           
            int start = filename.IndexOf("[");
            papermaterialcode = filename.Substring(0, start);


            string openfilepath = "";
            openfilepath = "\\\\192.168.0.204\\FilePath$\\图纸归档库" + "\\"+ papermaterialcode;

            //string filepath = FullNameByNode(treeListpaper.FocusedNode, 0);
            //string[] fileName = @filepath.Split('\\');
            //if (filepath.Contains("图纸归档库"))
            //{
            //    openfilepath = "\\\\192.168.0.204\\FilePath$" + filepath;
            //}
            //else
            //{
            //    openfilepath = "\\\\192.168.0.204\\FilePath$\\图纸归档库" + filepath;
            //}
            //if (openfilepath == "")
            //    return;

            string openfilename = Path.GetFileName(openfilepath+"\\"+ filename);
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

        private void txtpapermaterialcode_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13 && txtpapermaterialcode.Text.Trim() != "")
            {
                sBtnselectpaper_Click(sender,e);
            }
        }

        private void sBtnbrowsepaper_Click(object sender, EventArgs e)
        {

            OpenFileDialog ofdImport = new OpenFileDialog();
            ofdImport.Filter = "文件(*.jpg;png;jpeg;pdf)|*.jpg;*.png;*.jpeg;*.pdf";
            ofdImport.Multiselect = true;
            DialogResult dr = ofdImport.ShowDialog();
            if (dr == DialogResult.Cancel)
                return;     
                  
            txtpaperdoc.Text = "";
            foreach (string str in ofdImport.FileNames)
            {
               txtpaperdoc.Text += str + ",";              
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



        public bool connectState(string path)
        {
            bool Flag = false;
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
               ///// string dosLine = "net use " + path + " " + passWord + " /user:" + userName;

                string dosLine = @"net use \\" + path + " hytera;2012" + " /user:" + serverFilePath.Split('\\')[0] + "\\Upload";

                proc.StandardInput.WriteLine(dosLine);
                proc.StandardInput.WriteLine("exit");
                while (!proc.HasExited)
                {
                    proc.WaitForExit(1000);
                }
                string errormsg = proc.StandardError.ReadToEnd();
                proc.StandardError.Close();
                if (string.IsNullOrEmpty(errormsg))
                {
                    Flag = true;
                }
                else
                {
                    throw new Exception(errormsg);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                proc.Close();
                proc.Dispose();
            }
            return Flag;
        }


        /// <summary>  
        /// 从远程服务器下载文件到本地  
        /// </summary>  
        /// <param name="src">下载到本地后的文件路径，包含文件的扩展名</param>  
        /// <param name="dst">远程服务器路径（共享文件夹路径）</param>  
        /// <param name="fileName">远程服务器（共享文件夹）中的文件名称，包含扩展名</param>  
        public bool TransportRemoteToLocal(string src, string dst, string fileName)  
        {
            try
            {
                if (!Directory.Exists(dst))
                {
                    Directory.CreateDirectory(dst);
                }
                dst = dst + "\\" + fileName;
                FileStream inFileStream = new FileStream(dst, FileMode.Open);    //远程服务器文件  此处假定远程服务器共享文件夹下确实包含本文件，否则程序报错  

                FileStream outFileStream = new FileStream(src, FileMode.OpenOrCreate);   //从远程服务器下载到本地的文件  

                byte[] buf = new byte[inFileStream.Length];

                int byteCount;

                while ((byteCount = inFileStream.Read(buf, 0, buf.Length)) > 0)
                {

                    outFileStream.Write(buf, 0, byteCount);

                }
                inFileStream.Flush();

                inFileStream.Close();

                outFileStream.Flush();

                outFileStream.Close();

                return true;
            }
            catch
            {
                return false;
            }
        }


        /// <summary>  
        /// 将本地文件上传到远程服务器共享目录  
        /// </summary>  
        /// <param name="src">本地文件的绝对路径，包含扩展名</param>  
        /// <param name="dst">远程服务器共享文件路径，不包含文件扩展名</param>  
        /// <param name="fileName">上传到远程服务器后的文件扩展名</param>  
        public bool TransportLocalToRemote(string src, string dst, string fileName) 
        {
            try
            {
                FileStream inFileStream = new FileStream(src, FileMode.Open);    //此处假定本地文件存在，不然程序会报错                           
                if (!Directory.Exists(dst))        //判断上传到的远程服务器路径是否存在  
                {
                    Directory.CreateDirectory(dst);
                }
                dst = dst +"\\"+fileName;   //上传到远程服务器共享文件夹后文件的绝对路径  

                FileStream outFileStream = new FileStream(dst, FileMode.OpenOrCreate);

                byte[] buf = new byte[inFileStream.Length];

                int byteCount;

                while ((byteCount = inFileStream.Read(buf, 0, buf.Length)) > 0)
                {
                    outFileStream.Write(buf, 0, byteCount);
                }

                inFileStream.Flush();
                inFileStream.Close();
                outFileStream.Flush();
                outFileStream.Close();
                return true;
            }
            catch
            {
                return false;
            }
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
                       ///// string sfilename = Path.GetFileNameWithoutExtension(ss);
                        string sfilename = Path.GetFileName(ss);
                        if (sfilename == filename)
                        {
                            if (System.IO.File.Exists(ss))
                            {
                                System.IO.File.Delete(ss);
                                MessageBox.Show("删除成功", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            
                        }
                    }
                }
                else
                {
                    MessageBox.Show("连接失败，删除失败","提醒",MessageBoxButtons.OK,MessageBoxIcon.Information);
                }
            }
        }   

        private bool UploadReportFile(string filepath, string directype, string workno, string item)
        {
            bool b = false;
            if (Connect(serverFilePath))
            {
                bool re = true;
                if (filepath != "")
                {
                    try
                    {
                        string[] fileName = @filepath.Split('\\');
                        string floerPath = "\\\\" + this.serverFilePath + "\\" + directype + "\\" + workno + "\\" + item;
                        string fileServerPath = floerPath + "\\" + fileName[fileName.Length - 1];
                        if (Directory.Exists(floerPath) == false)  
                        {
                            Directory.CreateDirectory(floerPath);
                        }
                        File.Copy(filepath, fileServerPath, true);
                        re = true;
                    }
                    catch (Exception e)
                    {
                        re = false;
                    }
                }
                if (re)
                {
                    ////// MessageBox.Show("文件上传成功");
                    b = true;
                }
                else MessageBox.Show(filepath + " 文件上传失败");
            }
            else
            {
                MessageBox.Show("无法连接到服务器的共享目录");
            }
            return b;
        }


        public string GetPostfixStr(string filename)
        {
            try
            {
                int start = filename.LastIndexOf(".");
                int length = filename.Length;
                string postfix = filename.Substring(start, length - start);
                return postfix;
            }
            catch
            {
                return "";
            }
        }

        private bool CopyFileToServer(string folderType, string filePath)
        {
            try
            {
                string[] fileName = @filePath.Split('\\');
                string floerPath = "\\\\" + this.serverFilePath + "\\" + folderType + "\\" + txtpapermaterialcode.Text.Trim();

                string stringtime = DateTime.Now.ToString("yy-MM-dd-HH-mm-ss-fff");
                string filename = txtpapermaterialcode.Text.Trim() + "[" + stringtime + "]";

                //string fileServerPath = floerPath + "\\" + fileName[fileName.Length - 1];

                string filefix = GetPostfixStr(fileName[fileName.Length - 1]);
                string fileServerPath = floerPath + "\\" + filename +filefix;

                if (Directory.Exists(floerPath) == false)
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
                string te = txtpaperdoc.Text.TrimEnd(',');
                if (te != "")
                {
                    string[] copy = te.Split(',');
                    foreach (string ss in copy)
                    {
                        if (!CopyFileToServer("图纸归档库", ss))
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

        private void sBtnsavepaper_Click(object sender, EventArgs e)
        {
            if (txtpapermaterialcode.Text.Trim() == "")
                return;
            if (txtpaperdoc.Text == "")
                return;


            string fileDBServerPath = "";
            if (txtpaperdoc.Text != "")
            {

                UploadFile("");
                //string[] fileqty = txtpaperdoc.Text.TrimEnd(',').Split(',');
                //foreach (string s in fileqty)
                //{
                //    string[] fileName = @s.Split('\\');
                //    fileDBServerPath += "\\\\" + this.serverFilePath + "\\图纸归档库" + "\\" + txtpapermaterialcode.Text.Trim() + "\\" + fileName[fileName.Length - 1] + ",";
                //}

            }
            ////string floerPath = "";
            ////floerPath = "\\\\" + this.serverFilePath + "\\图纸归档库" + "\\" + txtpapermaterialcode.Text.Trim();
            ////if (floerPath != "")
            //// del_prefile(floerPath, txtpapermaterialcode.Text);
             //// UploadFile("");
            //// UploadReportFile(string filepath, string directype, string workno, string item)



        }

        private void sBtnselectpaper_Click(object sender, EventArgs e)
        {

            treeListpaper.DataSource = null;
            InitData();
            treeListpaper.VirtualTreeGetChildNodes -= new DevExpress.XtraTreeList.VirtualTreeGetChildNodesEventHandler(this.treeListpaper_VirtualTreeGetChildNodes);
            loadDrives = false;
            treeListpaper.VirtualTreeGetChildNodes += new DevExpress.XtraTreeList.VirtualTreeGetChildNodesEventHandler(this.treeListpaper_VirtualTreeGetChildNodes);

            { 
                InitData();
                treeListpaper.VirtualTreeGetChildNodes -= new DevExpress.XtraTreeList.VirtualTreeGetChildNodesEventHandler(this.treeListpaper_VirtualTreeGetChildNodes);
                loadDrives = false;
                treeListpaper.VirtualTreeGetChildNodes += new DevExpress.XtraTreeList.VirtualTreeGetChildNodesEventHandler(this.treeListpaper_VirtualTreeGetChildNodes);              
            }        
        }

        private void sBtndeletepaper_Click(object sender, EventArgs e)
        {
            string papermaterialcode = "", filefix = "";
            string filename = treeListpaper.FocusedNode.GetValue("Name").ToString();
            filefix = GetPostfixStr(filename);
            if (filefix.Contains(".db"))
            {
                return;
            }
            int start = filename.IndexOf("[");
            papermaterialcode = filename.Substring(0, start);

            string deletefilepath = "";
            deletefilepath = "\\\\192.168.0.204\\FilePath$\\图纸归档库" + "\\" + papermaterialcode;

             del_prefile(deletefilepath,filename);
             sBtnselectpaper_Click(sender,e);

        }

        private void sBtnresetpaper_Click(object sender, EventArgs e)
        {
            txtpapermaterialcode.Text = "";
            txtpapersequence.Text = "";
            txtpaperdoc.Text = "";
            lblmessage.Text = "";
            treeListpaper.DataSource = null;

        }

        private void sBtnbrowseproxy_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofdImport = new OpenFileDialog();
            ofdImport.Filter = "文件(*.jpg;png;jpeg;pdf)|*.jpg;*.png;*.jpeg;*.pdf";
            ofdImport.Multiselect = false;
            DialogResult dr = ofdImport.ShowDialog();
            if (dr == DialogResult.Cancel)
                return;
            txtproxydoc.Text = ofdImport.FileName;
        }

    

        private void sBtnsavaproxy_Click(object sender, EventArgs e)
        {
            if (txtsupper.Text == "" || txtbrand.Text == "" || txtsource.Text == "")
            {
                MessageBox.Show("请输入完整的信息","提醒",MessageBoxButtons.OK,MessageBoxIcon.Information);
                return;
            }

            if (txtsource.Text != "制造商" && txtproxydoc.Text== "")
            {
                MessageBox.Show("请上传相关的文件", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (txtsource.Text != "制造商" && txtExpiryDate.Text == "")
            {
                MessageBox.Show("请输入有效日期", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            string filename = "";
            string expirydate = "";
            if (txtExpiryDate.Text != "")
            {
                filename = ((DateTime)txtExpiryDate.EditValue).ToString("yyyyMMdd");       
                expirydate = ((DateTime)txtExpiryDate.EditValue).ToString("yyyy-MM-dd");
            }
            else
            {
                filename = "";
                expirydate = "";
            }

            if (txtproxydoc.Text.Trim() != "")
            {
                if (txtsource.Text == "制造商")
                    return;

                if (connectState(serverFilePath))
                {

                    string dirServerPath = "";
                    dirServerPath = "\\\\" + this.serverFilePath + "\\代理证归档库" + "\\" + txtsupper.Text + "\\" + txtbrand.Text + "\\" + txtsource.Text;////+ "\\" + fileName[fileName.Length - 1] + ",";
                    DirectoryInfo theFolder = new DirectoryInfo(dirServerPath);
                    string savafilename = theFolder.ToString();
                    string srcfile = txtproxydoc.Text;
                    string srcfilefix = GetPostfixStr(srcfile);
                    string dstfile = filename + srcfilefix;
                    bool flag = TransportLocalToRemote(@srcfile, savafilename, dstfile);  //实现将本地文件写入到远程服务器  
                    if (flag)
                    {
                        string sql = " if not exists( select 1 from IQC_ProxyCardLibrary where supplier= '" + txtsupper.Text + "' and brand = '" + txtbrand.Text + "' and  materialsource = '" + txtsource.Text + "' and  filename= '" + filename + "' ) ";
                        sql += " insert into IQC_ProxyCardLibrary ( supplier,brand,materialsource,expirydate ,filename,remark,updateman,updatetime ) ";
                        sql += " values('" + txtsupper.Text + "','" + txtbrand.Text + "','" + txtsource.Text + "','" + ((DateTime)txtExpiryDate.EditValue).ToString("yyyy-MM-dd") + "' ,'" + dstfile + "','" + txtremark.Text + "','" + Login.username + "' ,GETDATE() ) ";
                        bool mark = DbAccess.ExecuteSql(sql);
                        if (mark)
                        {
                            MessageBox.Show("保存成功", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            sBtnselectproxy_Click(sender, e);

                            txtsupper.Text = "";
                            txtbrand.Text = "";
                            txtsource.Text = "";
                            txtExpiryDate.Text = "";
                            txtproxydoc.Text = "";

                        }
                        else
                        {
                            MessageBox.Show("保存失败", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }

                }
                else
                {
                    Console.WriteLine("未能连接成功！");
                }
            }
            else
            {
                if (txtsource.Text != "制造商")
                    return;

                string sql = " if not exists( select 1 from IQC_ProxyCardLibrary where supplier= '" + txtsupper.Text + "' and brand = '" + txtbrand.Text + "' and  materialsource = '制造商' and  filename= '' ) ";
                sql += " insert into IQC_ProxyCardLibrary ( supplier,brand,materialsource,expirydate ,filename,remark,updateman,updatetime ) ";
                sql += " values('" + txtsupper.Text + "','" + txtbrand.Text + "','制造商','"+ expirydate + "','','" + txtremark.Text + "','" + Login.username + "' ,GETDATE() ) ";   ////////  " + ((DateTime)txtExpiryDate.EditValue).ToString("yyyy-MM-dd") + "
                bool mark = DbAccess.ExecuteSql(sql);
                if (mark)
                {
                    MessageBox.Show("保存成功", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    sBtnselectproxy_Click(sender, e);
                    txtsupper.Text = "";
                    txtbrand.Text = "";
                    txtsource.Text = "";
                    txtExpiryDate.Text = "";
                    txtproxydoc.Text = "";

                }
                else
                {
                    MessageBox.Show("保存失败或已经存在", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

            }
        }

        private void sBtnselectproxy_Click(object sender, EventArgs e)
        {
            string where = " where 1=1 ";
            string supper = txtsupper.Text;
            string brand = txtbrand.Text;
            string source = txtsource.Text;
            string ExpiryDate = "";
            if (txtExpiryDate.Text != "")
            {
                 ExpiryDate = ((DateTime)txtExpiryDate.DateTime).ToString("yyyy-MM-dd");
            }

            if (!string.IsNullOrEmpty(supper))
            {
                where += " and supplier = '" + supper + "' ";
            }
            if (!string.IsNullOrEmpty(brand))
            {
                where += " and brand = '" + brand + "' ";
            }
            if (!string.IsNullOrEmpty(source))
            {
                where += " and materialsource = '" + source + "' ";
            }
            //if (!string.IsNullOrEmpty(ExpiryDate))
            //{
            //    where += " and expirydate <='" + ExpiryDate + " '";
            //}

            string sql = @"  select supplier 供应商,brand 品牌,materialsource 进货来源,expirydate 有效日期,filename 文件名,remark 备注,updateman 更新人,updatetime 更新时间 from IQC_ProxyCardLibrary  ";
            sql += where + " order by updatetime desc ";

            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];

            if (dt != null && dt.Rows.Count > 0)
            {
                gridControlproxy.DataSource = dt;        
            }
            else
            {
               //////MessageBox.Show("没有符合条件的记录");
                gridControlproxy.DataSource = null;
            }

        }

        private void sBtndeleteproxy_Click(object sender, EventArgs e)
        {
            DataTable dt = gridControlproxy.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
                return;
            if (gridViewproxy.RowCount < 1)
                return;
            if (gridViewproxy.FocusedRowHandle < 0)
                return;

            string supper = gridViewproxy.GetFocusedRowCellValue("供应商").ToString();
            string brand = gridViewproxy.GetFocusedRowCellValue("品牌").ToString();
            string source = gridViewproxy.GetFocusedRowCellValue("进货来源").ToString();
            string filename =  gridViewproxy.GetFocusedRowCellValue("文件名").ToString();
            string sql = @" delete IQC_ProxyCardLibrary where supplier = '"+ supper + "'and brand = '"+ brand + "' and materialsource= '"+ source + "' and filename = '"+ filename + "' ";

            bool mark = DbAccess.ExecuteSql(sql);
            if (mark)
            {
                if (source != "制造商")
                {
                    string deletefilepath = "\\\\192.168.0.204\\FilePath$\\代理证归档库" + "\\" + supper + "\\" + brand + "\\" + source;
                    del_prefile(deletefilepath, filename);
                }
                sBtnselectproxy_Click(sender,e);               
               //// gridControlproxy.DataSource = null;
            }
            else
            {
                MessageBox.Show("删除失败","提醒",MessageBoxButtons.OK,MessageBoxIcon.Information);
            }
        }

        private void sBtndicupdown_Click(object sender, EventArgs e)
        {
            DataTable dt = gridControlproxy.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
                return;
            if (gridViewproxy.RowCount < 1)
                return;
            if (gridViewproxy.FocusedRowHandle < 0)
                return;

            string supper = gridViewproxy.GetFocusedRowCellValue("供应商").ToString();
            string brand = gridViewproxy.GetFocusedRowCellValue("品牌").ToString();
            string source = gridViewproxy.GetFocusedRowCellValue("进货来源").ToString();
            string filename = gridViewproxy.GetFocusedRowCellValue("文件名").ToString();
            string srcdown = "";
            string srcfilefix = GetPostfixStr(filename);


            SaveFileDialog savefilename = new SaveFileDialog();
            savefilename.Filter = "文件(*"+srcfilefix+ ")|*"+ srcfilefix ;
            savefilename.ShowDialog();
            srcdown = savefilename.FileName;

            if (srcdown == "")
                return;

            if (source == "制造商")
                return;

            if (connectState(serverFilePath))
            {
                string dirServerPath = "";
                dirServerPath = "\\\\" + this.serverFilePath + "\\代理证归档库" + "\\" + txtsupper.Text + "\\" + txtbrand.Text + "\\" + txtsource.Text;
                DirectoryInfo theFolder = new DirectoryInfo(dirServerPath);
                string savafilename = theFolder.ToString();

               bool mark =  TransportRemoteToLocal(srcdown, savafilename, filename);
                if (mark)
                {
                    MessageBox.Show("下载成功", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("下载失败", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

            }
            else
            {
                Console.WriteLine("未能连接成功！");
            }
        }

        private void sBtnresetproxy_Click(object sender, EventArgs e)
        {
             txtsupper.Text = "";
             txtbrand.Text = "";
             txtsource.Text ="";
             txtExpiryDate.Text = "";
             txtproxydoc.Text = "";
            gridControlproxy.DataSource = null;

        }

        private void gridViewproxy_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {

            DataTable dt = gridControlproxy.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
                return;
            if (gridViewproxy.RowCount < 1)
                return;
            if (gridViewproxy.FocusedRowHandle < 0)
                return;
            try
            {
                txtsupper.Text = gridViewproxy.GetFocusedRowCellValue("供应商").ToString();
                txtbrand.Text = gridViewproxy.GetFocusedRowCellValue("品牌").ToString();
                txtsource.Text = gridViewproxy.GetFocusedRowCellValue("进货来源").ToString();
               //// string filename = gridViewproxy.GetFocusedRowCellValue("文件名").ToString();
            }
            catch
            {
            }
        }

        private void gridViewproxy_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {
            DataTable dt = gridControlproxy.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
                return;
            if (gridViewproxy.RowCount < 1)
                return;
            if (gridViewproxy.FocusedRowHandle < 0)
                return;

            string supper = gridViewproxy.GetFocusedRowCellValue("供应商").ToString();
            string brand = gridViewproxy.GetFocusedRowCellValue("品牌").ToString();
            string source = gridViewproxy.GetFocusedRowCellValue("进货来源").ToString();
            string filename = gridViewproxy.GetFocusedRowCellValue("文件名").ToString();

            if (e.Column.FieldName == "文件名")
            {          
                string openfilepath = "\\\\192.168.0.204\\FilePath$\\代理证归档库" + "\\" + supper + "\\" + brand.TrimEnd()+ "\\" + source;
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
        private string ShowSaveFileDialog(string title, string filter)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            string name = "代理证信息";
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
            //progressBarControl1.Position = 0;
        }
        void gridView_PrintExportProgress(object sender, ProgressChangedEventArgs e)
        {
            SetPosition(e.ProgressPercentage);
        }
        void SetPosition(int pos)
        {
            //progressBarControl1.Position = pos;
            this.Update();
        }
        private void btntoexcel_Click(object sender, EventArgs e)
        {
            if (gridViewproxy.RowCount < 1)
                return;
            string fileName = ShowSaveFileDialog("Microsoft Excel 2007 Document", "Microsoft Excel|*.xlsx");
            if (fileName == string.Empty) return;
            ExportToEx(fileName, "xlsx", gridViewproxy);
            OpenFile(fileName);
        }
    }
}