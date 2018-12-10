using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Windows.Forms;
using DevExpress.XtraBars;
using DX_QMS.Common;
using System.IO;
using DevExpress.XtraEditors.Controls;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraEditors;
using System.Diagnostics;

namespace DX_QMS
{
    public partial class QARedOrder : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private string serverFilePath = "192.168.0.204\\FilePath$";
        public QARedOrder()
        {
            InitializeComponent();
        }

        private void bindorg_id()
        {
            string sql = "select ORG_ID,ORG_NAME from ORG_INFO where ORG_TYPE='ORG' order by SORT ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            txtorg_id.Properties.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
                txtorg_id.Properties.Items.Add(row["ORG_NAME"]);
            }
            txtorg_id.SelectedIndex = 0;

        }

        private void bindcustomer()
        {
            string sql = "select distinct customer  from OQC_TestListNew  ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            txtcustomer.Properties.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
                txtcustomer.Properties.Items.Add(row["customer"]);
            }
            txtcustomer.SelectedIndex = 0;
        }


        string[,] exportData = new string[,] 
        {
            {"Excel Document", "Microsoft Excel|*.xlsx", "xlsx"},
            {"PDF Document", "PDF Files|*.pdf", "pdf"}
        };

        void InitExportData()
        {
            for (int i = 0; i < exportData.GetLength(0); i++)
                cbExport.Properties.Items.Add(exportData.GetValue(i, 0));
            cbExport.SelectedIndex = 0;
        }
        private void setRule()
        {
            string post = "";
            if (!string.IsNullOrEmpty(Login.manager))   
            {
                post = Login.manager;
            }
            else
            {
                post = Login.post;
            }
            Dictionary<string, bool> dic = GroupPermission.QMS_SelectRulesForForm(post, "出检红单");
            //sBtnsave.Enabled = bool.Parse(dic["hasInsert"].ToString());
            sBtndelete.Enabled = bool.Parse(dic["hasDelete"].ToString());
        }

        private void QARedOrder_Load(object sender, EventArgs e)
        {
            bindorg_id();
            bindcustomer();
            InitExportData();
            setRule();
            btnreset_Click(null, null);
            begindate.Text = DateTime.Now.ToString("yyyy-MM-dd");
            enddate.Text = DateTime.Now.ToString("yyyy-MM-dd");
            gridView.PrintExportProgress += new ProgressChangedEventHandler(gridView_PrintExportProgress);
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


        private bool CopyFileToServer(string folderType, string filePath,string workno,string item,string pictype)
        {
            try
            {
                string[] fileName = @filePath.Split('\\');
                string floerPath = "\\\\" + this.serverFilePath + "\\" + folderType + "\\" + workno + "\\" + item + "\\" + pictype;
                string fileServerPath = floerPath + "\\" + fileName[fileName.Length - 1];
                if (Directory.Exists(floerPath) == false)   //如果不存在就创建file文件夹
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

        private bool UploadFile(string filepath,string directype, string workno, string item, string pictype)
        {
            bool b = false;
            if (Connect(serverFilePath))
            {
                bool re = true;
                if (filepath != "")
                {             
                        if (!CopyFileToServer(directype,filepath, workno,item, pictype))
                        {
                            re = false;
                        }                    
                }
                if (re)
                {
                    MessageBox.Show("操作成功");
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
                        if (Directory.Exists(floerPath) == false)   //如果不存在就创建file文件夹
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
                    MessageBox.Show("操作成功");
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

        private void btnselect_Click(object sender, EventArgs e)
        {

            string where = " where testresult = 'NG' ";
            string org_id = txtorg_id.Text.Trim();
            string customer = txtcustomer.Text.Trim();
            string workno = txtworkno.Text.Trim();
            string PN = txtPN.Text.Trim();      
            string begintime = begindate.Text;
            string endtime = enddate.Text;

            if (!string.IsNullOrEmpty(org_id))
            {
                where += " and org_id = '" + org_id + "' ";
            }
            if (!string.IsNullOrEmpty(customer))
            {
                where += " and customer = '" + customer + "' ";
            }
            if (!string.IsNullOrEmpty(workno))
            {
                where += " and workno = '" + workno + "' ";
            }
            if (!string.IsNullOrEmpty(PN))
            {
                where += " and hytcode = '" + PN + "' ";
            }
            if (!string.IsNullOrEmpty(begintime))
            {
                where += " and checkdate >= '" + begindate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00 '";
            }
            if (!string.IsNullOrEmpty(endtime))
            {
                where += " and checkdate <='" + enddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59 '";
            }

            string sql = @"  select checkdate 发生时间,case when ifovertime is null and checkdate >'2018-11-03 00:00:00' and  datediff(hour,checkdate,GETDATE())>48  then '是'  when ifovertime is null and checkdate >'2018-11-03 00:00:00' and  datediff(hour,checkdate,GETDATE())<=48  then '否'  else ifovertime end  是否超期, ";
            sql += " item 序列号,workno 工单,hytcode 编码,model 机型,customer 客户,lineid 线别,sendqty 送检数量,sampleqty 抽样数量,NGQty 不良数量,badinformation 不良描述 ,testman  检验人,QE  QE工程师,masters 责任人,CauseAnalysis 原因分析, improvemeasures 改善措施  from OQC_TestListNew  ";
            sql += where + " order by checkdate desc ";

            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];

            if (dt != null && dt.Rows.Count > 0)
            {
                gridControl.DataSource = dt;
                repositoryItemMemoEdit1.LinesCount = 3;
            }
            else
            {
                MessageBox.Show("没有符合条件的记录");
                gridControl.DataSource = null;
            }

             
        }

        private void btnreset_Click(object sender, EventArgs e)
        {
            txtCauseAnalysis.Text = "";
            txtimprovemeasures.Text = "";
            txtgoodpicture.Text = "";
            txtbadpicture.Text = "";
            goodpicture.Image = null;
            badpicture.Image = null;
            txtorg_id.Text = "";
            txtcustomer.Text = "";
            txtworkno.Text = "";
            txtPN.Text = "";
            begindate.Text  = "";
            enddate.Text = "";
            txtQAredreport.Text = "";
            gridControl.DataSource = null;
        }

        private string ShowSaveFileDialog(string title, string filter)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            string name = "QA红单列表信息";
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
        private void btnexport_Click(object sender, EventArgs e)
        {

            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null)
                return;
            if (dt.Rows.Count <= 0)
                return;
            int index = cbExport.SelectedIndex;
            if (index < 0)
                return;
            string fileName = ShowSaveFileDialog(exportData.GetValue(index, 0).ToString(), exportData.GetValue(index, 1).ToString());
            if (fileName == string.Empty) return;
            ExportToEx(fileName, exportData.GetValue(index, 2).ToString(), gridView);
            OpenFile(fileName);

        }


        void export(string item)
        {
            string sql = "";


            int sheetCount = 1;
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = true;
            object missing = System.Reflection.Missing.Value;
            string templetFile = Environment.CurrentDirectory + @"\ReportFolder\OQC异常处理单.xlsx";
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


            sql = @"  select checkdate 发生时间,workno 工单,hytcode 编码,model 机型,lineid 线别,sendqty 送检数量,sampleqty 抽样数量,NGQty 不良数量,badinformation 不良描述 ,testman  检验人, CauseAnalysis 原因分析, improvemeasures 改善措施  
                     from OQC_TestListNew where item = '"+ item + "'";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            if (dt == null || dt.Rows.Count < 1)
                return;
  
            DateTime date = (DateTime)dt.Rows[0]["发生时间"];
            string datestring = date.ToString("yyyy年MM月dd日");
            sheet.Cells.get_Range("A2").Value = "发生时间："+ datestring + "                              文件编号：HYT-RA-WI-QC-OEM-267-F03 V4.0";

            sheet.Cells.get_Range("B3").Value = dt.Rows[0]["工单"].ToString();
            sheet.Cells.get_Range("B4").Value = dt.Rows[0]["编码"].ToString();
            sheet.Cells.get_Range("B5").Value = dt.Rows[0]["机型"].ToString();
            sheet.Cells.get_Range("C5").Value = dt.Rows[0]["线别"].ToString();
            sheet.Cells.get_Range("D5").Value = dt.Rows[0]["送检数量"].ToString();
            sheet.Cells.get_Range("E5").Value = dt.Rows[0]["抽样数量"].ToString();
            sheet.Cells.get_Range("F3").Value = "不良数量（PCS）:" + dt.Rows[0]["不良数量"].ToString();

            sheet.Cells.get_Range("A6").Value ="1、不良描述"+ "\r\n" + dt.Rows[0]["不良描述"].ToString();
            sheet.Cells.get_Range("M7").Value = dt.Rows[0]["检验人"].ToString();
            sheet.Cells.get_Range("A8").Value = sheet.Cells.get_Range("A8").Value+ "\r\n"+ (dt.Rows[0]["原因分析"] == null ? "": dt.Rows[0]["原因分析"].ToString());
            sheet.Cells.get_Range("A10").Value = sheet.Cells.get_Range("A10").Value+ "\r\n"+(dt.Rows[0]["改善措施"] == null ? "": dt.Rows[0]["改善措施"].ToString());



        }

        private void sBtnreportdoc_Click(object sender, EventArgs e)
        {
            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
                return;
            if (gridView.FocusedRowHandle < 0)
                return;
            string workno = gridView.GetFocusedRowCellValue("工单").ToString();
            string item = gridView.GetFocusedRowCellValue("序列号").ToString();
            try
            {
                export(item);
            }
            catch
            {
                MessageBox.Show("生成报告失败","提醒",MessageBoxButtons.OK ,MessageBoxIcon.Information);
            }
        }
        public Stream BytesToStream(byte[] bytes)
        {
            Stream stream = new MemoryStream(bytes);
            return stream;
        }

        public void StreamToFile(Image img, string dst, string fileName)
        {
            //Stream stream;
            //// 把 Stream 转换成 byte[] 
            //byte[] bytes = new byte[stream.Length];
            //stream.Read(bytes, 0, bytes.Length);
            //// 设置当前流的位置为流的开始 
            //stream.Seek(0, SeekOrigin.Begin);
            //// 把 byte[] 写入文件 
            //dst = dst + "\\" + fileName;                        
            //FileStream fs = new FileStream(dst, FileMode.OpenOrCreate);
            //BinaryWriter bw = new BinaryWriter(fs);
            //bw.Write(bytes);
            //bw.Close();
            //fs.Close();
        }


        /// <summary>  
        /// 将本地文件上传到远程服务器共享目录  
        /// </summary>  
        /// <param name="src">本地文件的绝对路径，包含扩展名</param>  
        /// <param name="dst">远程服务器共享文件路径，不包含文件扩展名</param>  
        /// <param name="fileName">上传到远程服务器后的文件扩展名</param>  
        public bool TransportLocalToRemote(Image img, string dst, string fileName)
        {
            try
            {
                //FileStream inFileStream = new FileStream(src, FileMode.Open);    //此处假定本地文件存在，不然程序会报错                           

                MemoryStream inMemoryStream = new MemoryStream();
                byte[] imagedata = null;
                img.Save(inMemoryStream, System.Drawing.Imaging.ImageFormat.Jpeg);
                imagedata = inMemoryStream.GetBuffer();

                if (!Directory.Exists(dst))        //判断上传到的远程服务器路径是否存在  
                {
                    Directory.CreateDirectory(dst);
                }
                dst = dst + "\\" + fileName;   //上传到远程服务器共享文件夹后文件的绝对路径              
                FileStream outFileStream = new FileStream(dst, FileMode.OpenOrCreate);

                byte[] buf = new byte[imagedata.Length];
                buf = imagedata;
                int byteCount;
             
                outFileStream.Write(buf, 0, imagedata.Length);

                inMemoryStream.Flush();
                inMemoryStream.Close();
                outFileStream.Flush();
                outFileStream.Close();
                return true;
            }
            catch
            {
                return false;
            }
        }

        //写byte[]到fileName
        private bool writeFile(byte[] pReadByte, string fileName)
        {
            FileStream pFileStream = null;
            try
            {
                pFileStream = new FileStream(fileName, FileMode.OpenOrCreate);
                pFileStream.Write(pReadByte, 0, pReadByte.Length);
            }
            catch
            {
                return false;
            }
            finally
            {
                if (pFileStream != null)
                    pFileStream.Close();
            }
            return true;
        }



        /// <summary>
        /// 图片转换成字节流
        /// </summary>
        /// <param name="img">要转换的Image对象</param>
        /// <returns>转换后返回的字节流</returns>
        public byte[] ImgToByt(Image img)
        {
            MemoryStream ms = new MemoryStream();
            byte[] imagedata = null;
            img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
            imagedata = ms.GetBuffer();
            return imagedata;
        }



        private void sBtnsave_Click(object sender, EventArgs e)
        {
            if (txtCauseAnalysis.Text.Trim () == "" || txtimprovemeasures.Text.Trim() == "")
            {

                DialogResult Rt = MessageBox.Show("原因分析和改善对策不完整, 是否继续?", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                if (DialogResult.Cancel == Rt)
                {
                    return;
                }
            }
            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
                return;
            if (gridView.FocusedRowHandle < 0)
                return;

            string workno = gridView.GetFocusedRowCellValue("工单").ToString();
            string item = gridView.GetFocusedRowCellValue("序列号").ToString();

            string goodpicServerDirec = "";
            if (txtgoodpicture.Text != "")
            {
                goodpicServerDirec = "\\\\" + serverFilePath + "\\QA红单图片" + "\\" + workno + "\\" + item + "\\" + "良品图片";
                string s = txtgoodpicture.Text;
                string[] fileName = @s.Split('\\');

                if (goodpicServerDirec != "")
                    del_prefile(goodpicServerDirec, "");
                UploadFile(txtgoodpicture.Text,"QA红单图片", workno, item, "良品图片");
            }
            else
            {
                if (goodpicture.Image != null)
                {
                    goodpicServerDirec = "\\\\" + serverFilePath + "\\QA红单图片" + "\\" + workno + "\\" + item + "\\" + "良品图片";
                    if (goodpicServerDirec != "")
                        del_prefile(goodpicServerDirec, "");
                    if (Connect(serverFilePath))
                    {
                       // string s = goodpicServerDirec + "\\" + DateTime.Now.ToString("yyMMddHHmmssfff") + ".jpg";
                       // goodpicture.Image.Save(s, System.Drawing.Imaging.ImageFormat.Jpeg);
                      bool flat =  TransportLocalToRemote(goodpicture.Image, goodpicServerDirec,DateTime.Now.ToString("yyMMddHHmmssfff") + ".jpg");
                        if (!flat)
                        {
                            MessageBox.Show("良品图片保存失败", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    else
                    {
                        MessageBox.Show("无法连接到服务器的共享目录");
                    }                       
                }
            }
            string badpicServerDirec = "";
            if (txtbadpicture.Text != "")
            {
                badpicServerDirec = "\\\\" + serverFilePath + "\\QA红单图片" + "\\" + workno + "\\" + item + "\\" + "不良品图片";
                string s = txtbadpicture.Text;
                string[] fileName = @s.Split('\\');
                if (badpicServerDirec != "")
                    del_prefile(badpicServerDirec, "");
                UploadFile(txtbadpicture.Text, "QA红单图片", workno, item, "不良品图片");
            }
            else
            {
                if (badpicture.Image != null)
                {
                    badpicServerDirec = "\\\\" + serverFilePath + "\\QA红单图片" + "\\" + workno + "\\" + item + "\\" + "不良品图片";
                    if (badpicServerDirec != "")
                        del_prefile(badpicServerDirec, "");
                    if (Connect(serverFilePath))
                    {
                        //badpicture.Image.Save(badpicServerDirec + "\\" + DateTime.Now.ToString("yyMMddHHmmssfff") + ".jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
                        bool flat = TransportLocalToRemote(badpicture.Image, badpicServerDirec, DateTime.Now.ToString("yyMMddHHmmssfff") + ".jpg");
                        if (!flat)
                        {
                            MessageBox.Show("不良品图片保存失败", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    else
                    {
                        MessageBox.Show("无法连接到服务器的共享目录");
                    }
                }
            }
            if (txtCauseAnalysis.Text.Trim() != "" || txtimprovemeasures.Text.Trim() != "")
            {
                string updatesql = @"  update OQC_TestListNew set CauseAnalysis = '" + txtCauseAnalysis.Text + "',improvemeasures = '" + txtimprovemeasures.Text + "',ifovertime = case when ifovertime is null and datediff(hour,checkdate,GETDATE())>48 then '是'  when ifovertime is null and datediff(hour,checkdate,GETDATE())<=48 then '否'  else ifovertime end  where workno = '" + workno + "' and item = '" + item + "'  ";
                bool falg = DbAccess.ExecuteSql(updatesql);
            }

            btnselect_Click(sender, e);

            //if (falg)
            //{
            //    MessageBox.Show("保存成功","提醒",MessageBoxButtons.OK ,MessageBoxIcon.Information);
            //}
            //else
            //{
            //     MessageBox.Show("保存失败", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //}
        }

        private void sBtngoodpicture_Click(object sender, EventArgs e)
        {

            OpenFileDialog ofdImport = new OpenFileDialog();
            ofdImport.Filter = "图像文件(*.jpg;bmp;png;jpeg)|*.jpg;*.bmp;*.png;*.jpeg";
            ofdImport.Multiselect = false;
            DialogResult dr = ofdImport.ShowDialog();
            if (dr == DialogResult.Cancel)                          
               return;

            txtgoodpicture.Text = ofdImport.FileName;
            FileStream fs = new FileStream(txtgoodpicture.Text, FileMode.Open);
            Image bt = Image.FromStream(fs);
            fs.Close();
            fs.Dispose();
            goodpicture.Image = bt;
            goodpicture.Properties.SizeMode = PictureSizeMode.Stretch;

        }

        private void sBtnbadpicture_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofdImport = new OpenFileDialog();
            ofdImport.Filter = "图像文件(*.jpg;bmp;png;jpeg)|*.jpg;*.bmp;*.png;*.jpeg";
            ofdImport.Multiselect = false;
            DialogResult dr = ofdImport.ShowDialog();
            if (dr == DialogResult.Cancel)
                return;
            txtbadpicture.Text = ofdImport.FileName;
            FileStream fs = new FileStream(txtbadpicture.Text, FileMode.Open);
            Image bt = Image.FromStream(fs);
            fs.Close();
            fs.Dispose();
            badpicture.Image = bt;
            badpicture.Properties.SizeMode = PictureSizeMode.Stretch;
        }

        private void browse_Click(object sender, EventArgs e)
        {

            OpenFileDialog ofdImport = new OpenFileDialog();
            ofdImport.Filter = "文件(*.jpg;png;jpeg;tif;pdf)|*.jpg;*.png;*.jpeg;*.tif;*.pdf";   /////  tif
            ofdImport.Multiselect = false;
            DialogResult dr = ofdImport.ShowDialog();
            if (dr == DialogResult.Cancel)
                return;
            this.txtQAredreport.Text = "";
            this.txtQAredreport.Text = ofdImport.FileName;
            
        }
        private void btnupload_Click(object sender, EventArgs e)
        {

            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
                return;
            if (gridView.FocusedRowHandle < 0)
                return;

            string workno = gridView.GetFocusedRowCellValue("工单").ToString();
            string item = gridView.GetFocusedRowCellValue("序列号").ToString();

            if (txtQAredreport.Text.Trim() == "")
            {
                return;
            }
            string fileDBServerPath = "";
            if (txtQAredreport.Text != "")
            {
                string s = txtQAredreport.Text;
                string[] fileName = @s.Split('\\');
                fileDBServerPath += "\\\\" + this.serverFilePath + "\\QA红单报告" + "\\" + workno + "\\" + item + "\\" + fileName[fileName.Length - 1];       
            }
            string floerPath = "";
            floerPath = "\\\\" + this.serverFilePath + "\\QA红单报告" + "\\" + workno + "\\" + item;
            if (floerPath != "")
               del_prefile(floerPath, "");
            bool b = UploadReportFile(txtQAredreport.Text,"QA红单报告",workno,item);    
            if (b)
            {
                string updatesql = @"  update OQC_TestListNew set CauseAnalysis = '" + txtCauseAnalysis.Text + "',improvemeasures = '" + txtimprovemeasures.Text + "',ifovertime = case when ifovertime is null and datediff(hour,checkdate,GETDATE())>48 then '是'  when ifovertime is null and datediff(hour,checkdate,GETDATE())<=48 then '否'   else ifovertime  end  where workno = '" + workno + "' and item = '" + item + "'  ";
                bool falg = DbAccess.ExecuteSql(updatesql);
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

        private void gridView_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {
            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
                return;
            if (gridView.FocusedRowHandle < 0)
                return;
            string workno = gridView.GetFocusedRowCellValue("工单").ToString();
            string item = gridView.GetFocusedRowCellValue("序列号").ToString();

            if (e.Column.FieldName == "序列号")
            {

                string floerPath = "\\\\" + serverFilePath + "\\QA红单报告" + "\\" + workno + "\\" + item;

                if (Connect(serverFilePath))
                {
                    if (Directory.Exists(floerPath) == false)
                    {
                        MessageBox.Show("还没有上传测试报告!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        return;
                    }

                    string[] str = Directory.GetFiles(floerPath);
                    if (str.Length > 0)
                    {
                        for (int i = 0; i < str.Length; i++)
                        {
                            if (str[i].ToString().Contains(".db"))
                                continue;
                            string filename = Path.GetFileName(str[i].ToString());

                            FileStream fs = new FileStream(str[i].ToString(), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
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
            if (e.Column.FieldName == "编码")
            {
                string goodfilePath = "\\\\" + serverFilePath + "\\QA红单图片" + "\\" + workno + "\\" + item + "\\良品图片";
                string badfilePath = "\\\\" + serverFilePath + "\\QA红单图片" + "\\" + workno + "\\" + item + "\\不良品图片";
                Cursor.Current = Cursors.WaitCursor;
                if (Connect(serverFilePath))
                {
                    if (Directory.Exists(goodfilePath))
                    {

                        string[] pt = Directory.GetFiles(goodfilePath);
                        if (pt.Length == 0)
                            return;
                        if (pt.Length > 0)
                        {
                            for (int i = 0; i < pt.Length; i++)
                            {
                                if (pt[i].ToString().Contains(".db"))
                                   continue;
                                FileStream fs = new FileStream(pt[i].ToString(), FileMode.Open);
                                Image goodpic = Image.FromStream(fs);
                                fs.Close();
                                fs.Dispose();
                                goodpicture.Image = goodpic;
                                goodpicture.Properties.SizeMode = PictureSizeMode.Stretch;

                            }

                         }

                    }

                    if (Directory.Exists(badfilePath))
                    {
                        string[] pt = Directory.GetFiles(badfilePath);
                        if (pt.Length == 0)
                            return;
                        if (pt.Length > 0)
                        {
                            for (int i = 0; i < pt.Length; i++)
                            {
                                if (pt[i].ToString().Contains(".db"))
                                    continue;
                                FileStream ft = new FileStream(pt[i].ToString(), FileMode.Open);
                                Image badpic = Image.FromStream(ft);
                                ft.Close();
                                ft.Dispose();
                                badpicture.Image = badpic;
                                badpicture.Properties.SizeMode = PictureSizeMode.Stretch;
                            }

                        }                  
                    }
                }
                Cursor.Current = Cursors.Default;
            }
           
        }

        private void gridView_CustomUnboundColumnData(object sender, CustomColumnDataEventArgs e)
        {

            //DataTable dt = gridControl.DataSource as DataTable;
            //if (dt == null || dt.Rows.Count < 1)
            //    return;
            //if (gridView.FocusedRowHandle < 0)
            //    return;

            // int i = 0;
            //for ( ; i < gridView.RowCount; i++)
            //{
            //    if (i == gridView.RowCount)
            //        break;

            //    string item = (string)((DataRowView)e.Row)["序列号"];
            //    string workno = (string)((DataRowView)e.Row)["工单"];

            //    string goodfilePath = "\\\\" + serverFilePath + "\\QA红单图片" + "\\" + workno + "\\" + item + "\\良品图片";
            //    string badfilePath = "\\\\" + serverFilePath + "\\QA红单图片" + "\\" + workno + "\\" + item + "\\不良品图片";

            //    if (e.Column.FieldName == "良品图片" && e.IsGetData)
            //    {
            //        try
            //        {
            //            Image img = null;
            //            if (Connect(serverFilePath))
            //            {
            //                if (Directory.Exists(goodfilePath))
            //                {
            //                    //FileStream fs = new FileStream(Directory.GetFiles(goodfilePath)[0].ToString(), FileMode.Open);
            //                    //Image goodpic = Image.FromStream(fs);
            //                    //fs.Close();
            //                    //fs.Dispose();
            //                    //gridView.SetRowCellValue(e.ListSourceRowIndex, gridView.Columns["良品图片"], goodpic);

            //                    img = Image.FromFile(Directory.GetFiles(goodfilePath)[0].ToString());
            //                    gridView.SetRowCellValue(e.ListSourceRowIndex, gridView.Columns["良品图片"], img);
            //                    // e.Value = img;


            //                }
            //            }
            //        }
            //        catch (Exception ex)
            //        {
            //            MessageBox.Show(ex.ToString());
            //        }
            //    }
            //    if (e.Column.FieldName == "不良品图片" && e.IsGetData)
            //    {
            //        try
            //        {
            //            if (Connect(serverFilePath))
            //            {
            //                if (Directory.Exists(badfilePath))
            //                {
            //                    FileStream ft = new FileStream(Directory.GetFiles(badfilePath)[0].ToString(), FileMode.Open);
            //                    Image badpic = Image.FromStream(ft);
            //                    ft.Close();
            //                    ft.Dispose();
            //                    gridView.SetRowCellValue(e.ListSourceRowIndex, gridView.Columns["不良品图片"], badpic);
            //                }
            //            }


            //        }
            //        catch (Exception ex)
            //        {
            //            MessageBox.Show(ex.ToString());
            //        }
            //    }
            //}


        }

        private void gridView_CustomDrawCell(object sender, RowCellCustomDrawEventArgs e)
        {

            //DataTable dt = gridControl.DataSource as DataTable;
            //if (dt == null || dt.Rows.Count < 1)
            //    return;
            //if (gridView.FocusedRowHandle < 0)
            //    return;
            //if (e.Column.FieldName == "良品图片" || e.Column.FieldName == "良品图片")
            //{
            //    try
            //    {
            //        Image img = null;

            //        string item = gridView.GetDataRow(e.RowHandle)["序列号"].ToString();
            //        string workno = gridView.GetDataRow(e.RowHandle)["工单"].ToString();

            //        string goodfilePath = "\\\\" + serverFilePath + "\\QA红单图片" + "\\" + workno + "\\" + item + "\\良品图片";
            //        string badfilePath = "\\\\" + serverFilePath + "\\QA红单图片" + "\\" + workno + "\\" + item + "\\不良品图片";

            //        if (Connect(serverFilePath))
            //        {
            //            if (Directory.Exists(goodfilePath))
            //            {
            //                FileStream fs = new FileStream(Directory.GetFiles(goodfilePath)[0].ToString(), FileMode.Open);
            //                Image goodpic = Image.FromStream(fs);
            //                fs.Close();
            //                fs.Dispose();
            //                gridView.SetRowCellValue(e.RowHandle, gridView.Columns["良品图片"], goodpic);

            //                //img = Image.FromFile(Directory.GetFiles(goodfilePath)[0].ToString());
            //                //gridView.SetRowCellValue(e.ListSourceRowIndex, gridView.Columns["良品图片"], img);
            //                // e.Value = img;


            //            }

            //            if (Directory.Exists(badfilePath))
            //            {
            //                FileStream ft = new FileStream(Directory.GetFiles(badfilePath)[0].ToString(), FileMode.Open);
            //                Image badpic = Image.FromStream(ft);
            //                ft.Close();
            //                ft.Dispose();
            //                gridView.SetRowCellValue(e.RowHandle, gridView.Columns["不良品图片"], badpic);

            //                //img = Image.FromFile(Directory.GetFiles(goodfilePath)[0].ToString());
            //                //gridView.SetRowCellValue(e.ListSourceRowIndex, gridView.Columns["良品图片"], img);
            //                // e.Value = img;


            //            }


            //        }
            //    }
            //    catch (Exception ex)
            //    {
            //        MessageBox.Show(ex.ToString());
            //    }
            //}

            ////if (e.Column.FieldName == "不良品图片" && e.IsGetData)
            ////{
            ////    try
            ////    {
            ////        if (Connect(serverFilePath))
            ////        {
            ////            if (Directory.Exists(badfilePath))
            ////            {
            ////                FileStream ft = new FileStream(Directory.GetFiles(badfilePath)[0].ToString(), FileMode.Open);
            ////                Image badpic = Image.FromStream(ft);
            ////                ft.Close();
            ////                ft.Dispose();
            ////                gridView.SetRowCellValue(e.ListSourceRowIndex, gridView.Columns["不良品图片"], badpic);
            ////            }
            ////        }


            ////    }
            ////    catch (Exception ex)
            ////    {
            ////        MessageBox.Show(ex.ToString());
            ////    }
            ////}
        }

        private void sBtndelete_Click(object sender, EventArgs e)
        {

            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
                return;
            if (gridView.FocusedRowHandle < 0)
                return;
            string workno = gridView.GetFocusedRowCellValue("工单").ToString();
            string item = gridView.GetFocusedRowCellValue("序列号").ToString();
            string goodpicServerDirec = "\\\\" + serverFilePath + "\\QA红单图片" + "\\" + workno + "\\" + item + "\\" + "良品图片";
            string badpicServerDirec = "\\\\" + serverFilePath + "\\QA红单图片" + "\\" + workno + "\\" + item + "\\" + "不良品图片";
            string reportPath = "\\\\" + this.serverFilePath + "\\QA红单报告" + "\\" + workno + "\\" + item;

            string updatesql = @"  update OQC_TestListNew set CauseAnalysis = '',improvemeasures = '' where workno = '" + workno + "' and item = '" + item + "'  ";
            bool falg = DbAccess.ExecuteSql(updatesql);

            if (falg == true)
            {
                try
                {
                    del_prefile(goodpicServerDirec,"");
                    del_prefile(badpicServerDirec,"");
                    del_prefile(reportPath,"");
                }
                catch
                {               
                }
                MessageBox.Show("删除成功", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        public DateTime beforeTime, afterTime;

        private void gridView_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
        {

            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
            {
                return;
            }
            if (gridView.RowCount < 1)
                return;
            if (gridView.GetDataRow(e.RowHandle)["是否超期"].ToString() == "是")
            {
                e.Appearance.BackColor = Color.Red;
            }

        }

        private void gridView_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
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
                txtCauseAnalysis.Text = gridView.GetFocusedRowCellValue("原因分析").ToString();
                txtimprovemeasures.Text = gridView.GetFocusedRowCellValue("改善措施").ToString();
            }
            catch
            {
            }
        }

        private void txtworkno_Leave(object sender, EventArgs e)
        {
            if (txtworkno.Text.Trim() == "")
                return;
            btnselect_Click(null , null);
        }

        private void txtworkno_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13 && txtworkno.Text != "")
            {
                txtworkno_Leave(sender, e);
            }
        }

        private void QARedOrder_FormClosing(object sender, FormClosingEventArgs e)
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