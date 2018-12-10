using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using DevExpress.XtraBars;
using System.Data;
using System.Data.SqlClient;
using DX_QMS.Common;
using DevExpress.XtraEditors.Controls;

namespace DX_QMS
{
    public partial class IQCProdPic : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        [DllImport("user32.dll")]
        public static extern int WindowFromPoint(int xPoint, int yPoint);

        private int hHwnd;
        private string Mano = "";
        private string serverFilePath = "";
        private int count = 0;
        private Microsoft.Office.Interop.Excel.ApplicationClass appClsExcel = null;
        IQC ic = new IQC();
        public IQCProdPic()
        {
            InitializeComponent();
            setRule();
            string path = System.Configuration.ConfigurationManager.AppSettings["ServerFilePath"].ToString();
            this.serverFilePath = @path;
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
            Dictionary<string, bool> dic = GroupPermission.QMS_SelectRulesForForm(post, "丝印图");
            btnsave.Enabled = bool.Parse(dic["hasInsert"].ToString());
            btndel.Enabled = bool.Parse(dic["hasDelete"].ToString());
        }
        void IQCProdPic_MouseWheel(object sender, MouseEventArgs e)
        {
            System.Drawing.Point p = PointToScreen(e.Location);
            if (WindowFromPoint(p.X, p.Y) == p1.Handle.ToInt32())
            {

                //向前
                if (e.Delta > 0)
                {
                    float w = this.p1.Width * 0.9f; //每次縮小 20%  
                    float h = this.p1.Height * 0.9f;
                    this.p1.Size = Size.Ceiling(new SizeF(w, h));

                }

                //向后
                else if (e.Delta < 0)
                {

                    float w = this.p1.Width * 1.1f; //每次放大 20%
                    float h = this.p1.Height * 1.1f;
                    this.p1.Size = Size.Ceiling(new SizeF(w, h));
                    p1.Invalidate();

                }
            }
        }
        private void IQCProdPic_Load(object sender, EventArgs e)
        {
            this.MouseWheel += new MouseEventHandler(IQCProdPic_MouseWheel);
            this.setRule();
            BtnStop.Enabled = false;
            btnphoto.Enabled = false;
        }

        private void txtPNo_Leave(object sender, EventArgs e)
        {
            if (txtPNo.Text == "") return;
            this.txtPNo.Leave -= this.txtPNo_Leave;
            string sql = "select ORGANIZATION_ID,segment1 materialcode,description materialname,en_des materialenname,INVENTORY_ITEM_ID,PRIMARY_UOM_CODE from cux_item_material_v  where " +
                                 " segment1 ='" + txtPNo.Text.Trim() + "'";
            DataSet ds = DbAccess.SelectByOracle(sql);

            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                txtname.Text = ds.Tables[0].Rows[0]["materialname"].ToString();
                Mano = ds.Tables[0].Rows[0]["materialcode"].ToString();
                txtsup.Focus();
                txtsup.SelectAll();

                string ssql = "select * from MaterialBitPIC where PNO like '" + txtPNo.Text.Trim() + "%' order  by eventtime asc ";
                DataSet dsinfo = DbAccess.SelectBySql(ssql);
                if (dsinfo != null && dsinfo.Tables.Count > 0 && dsinfo.Tables[0].Rows.Count > 0)
                {
                    txtsup.Text = dsinfo.Tables[0].Rows[0]["sup"].ToString();
                    txtchecklist.Text = dsinfo.Tables[0].Rows[0]["checklist"].ToString();
                    txtremarks.Text = dsinfo.Tables[0].Rows[0]["remarks"].ToString();
                }
                else
                {
                    txtsup.Text = "";
                    txtchecklist.Text = "";
                    txtremarks.Text = "";
                }
            }
            else
            {
                string ssql = "select materialcode from delivery where lotno='" + txtPNo.Text + "'";
                DataSet dslotno = DbAccess.SelectBySql(ssql);
                if (dslotno != null && ds.Tables.Count > 0 && dslotno.Tables[0].Rows.Count > 0)
                {
                    string materialcode = dslotno.Tables[0].Rows[0]["materialcode"].ToString();
                    string Orasqlbylotno = "select organization_id organization,segment1 materialcode,description materialname,en_des materialenname,INVENTORY_ITEM_ID,PRIMARY_UOM_CODE from cux_item_material_v where segment1='" + materialcode + "'";
                    ds = DbAccess.SelectByOracle(Orasqlbylotno);
                    if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                    {
                        txtPNo.Text = materialcode;
                        txtname.Text = ds.Tables[0].Rows[0]["materialname"].ToString();
                        txtname.ForeColor = Color.Blue;
                        txtsup.Focus();
                        txtsup.SelectAll();

                        string sqltemp = "select * from MaterialBitPIC where PNO like '" + txtPNo.Text.Trim() + "%' order  by eventtime asc ";
                        DataSet dsinfo = DbAccess.SelectBySql(sqltemp);
                        if (dsinfo != null && dsinfo.Tables.Count > 0 && dsinfo.Tables[0].Rows.Count > 0)
                        {
                            txtsup.Text = dsinfo.Tables[0].Rows[0]["sup"].ToString();
                            txtchecklist.Text = dsinfo.Tables[0].Rows[0]["checklist"].ToString();
                            txtremarks.Text = dsinfo.Tables[0].Rows[0]["remarks"].ToString();
                        }
                        else
                        {
                            txtsup.Text = "";
                            txtchecklist.Text = "";
                            txtremarks.Text = "";
                        }
                    }
                }
                else
                {
                    MessageBox.Show(txtPNo.Text + "不存在", "提醒");
                    txtPNo.Focus();
                    txtPNo.Text = "";
                    Mano = "";
                    txtsup.Text = "";
                    txtchecklist.Text = "";
                    txtremarks.Text = "";
                }
            }
            this.txtPNo.Leave += this.txtPNo_Leave;
        }

        private void btnpiclist_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofdImport = new OpenFileDialog();
            ofdImport.Filter = "图像文件(*.jpg;bmp;png;jpeg)|*.jpg;*.bmp;*.png;*.jpeg";
            ofdImport.Multiselect = true;
            DialogResult dr = ofdImport.ShowDialog();
            if (dr == DialogResult.Cancel) return;
            this.txtpiclist.Text = "";
            foreach (string str in ofdImport.FileNames)
            {
                //this.txtpiclist.Text = ofdImport.FileName;
                this.txtpiclist.Text += str + ",";
                if (txtpiclist.Text != "")
                {
                    string[] pt = @txtpiclist.Text.TrimEnd(',').Split(',');
                    if (pt.Length == 1)
                    {
                        FileStream fs = new FileStream(pt[0].ToString(), FileMode.Open);
                        Image bt = Image.FromStream(fs);
                        fs.Close();
                        fs.Dispose();
                        p1.Image = bt;
                        p1.Properties.SizeMode = PictureSizeMode.Stretch;
                    }
                    else if (pt.Length == 2)
                    {
                        FileStream fs = new FileStream(pt[0].ToString(), FileMode.Open);
                        Image bt = Image.FromStream(fs);
                        fs.Close();
                        fs.Dispose();
                        p1.Image = bt;
                        p1.Properties.SizeMode = PictureSizeMode.Stretch;

                        FileStream fs2 = new FileStream(pt[1].ToString(), FileMode.Open);
                        Image bt2 = Image.FromStream(fs2);
                        fs2.Close();
                        fs2.Dispose();
                        p2.Image = bt2;
                        p2.Properties.SizeMode = PictureSizeMode.Stretch;
                    }
                    else if (pt.Length == 3)
                    {
                        FileStream fs = new FileStream(pt[0].ToString(), FileMode.Open);
                        Image bt = Image.FromStream(fs);
                        fs.Close();
                        fs.Dispose();
                        p1.Image = bt;
                        p1.Properties.SizeMode = PictureSizeMode.Stretch;

                        FileStream fs2 = new FileStream(pt[1].ToString(), FileMode.Open);
                        Image bt2 = Image.FromStream(fs2);
                        fs2.Close();
                        fs2.Dispose();
                        p2.Image = bt2;
                        p2.Properties.SizeMode = PictureSizeMode.Stretch;

                        FileStream fs3 = new FileStream(pt[2].ToString(), FileMode.Open);
                        Image bt3 = Image.FromStream(fs3);
                        fs3.Close();
                        fs3.Dispose();
                        p3.Image = bt3;
                        p3.Properties.SizeMode = PictureSizeMode.Stretch;
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
                string floerPath = "\\\\" + this.serverFilePath + "\\" + folderType + "\\" + this.txtPNo.Text.Trim();
                string fileServerPath = floerPath + "\\" + fileName[fileName.Length - 1];
                if (Directory.Exists(floerPath) == false)//如果不存在就创建file文件夹
                {
                    Directory.CreateDirectory(floerPath);
                }

                File.Copy(filePath, fileServerPath, true);
                count++;
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
                count = 0;
                bool re = true;
                string msg = "";
                string te = txtpiclist.Text.TrimEnd(',');
                if (te != "")
                {
                    string[] copy = te.Split(',');
                    foreach (string ss in copy)
                    {
                        if (!CopyFileToServer("丝印图", ss))
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


        void AddRow(DevExpress.XtraGrid.Views.Grid.GridView view)
        {
            int prevDataRowIndex = view.GetFocusedDataSourceRowIndex();
            view.AddNewRow();
            if (view.GroupCount >= 0 && prevDataRowIndex >= 0)
            {
                foreach (DevExpress.XtraGrid.Columns.GridColumn groupColumn in view.GroupedColumns)
                {
                    object val = view.GetRowCellValue(prevDataRowIndex, groupColumn);
                    view.SetFocusedRowCellValue(groupColumn, val);
                }
                view.UpdateCurrentRow();
            }
            view.ShowEditor();
        }




        private void bindIQCProdPicInfo(string mano)
        {
            //this.databind.Rows.Add(1);
            //this.databind.Rows[this.databind.Rows.Count - 1].Cells["PNo"].Value = mano;
            //this.databind.Rows[this.databind.Rows.Count - 1].Cells["PName"].Value = txtname.Text;
            //this.databind.Rows[this.databind.Rows.Count - 1].Cells["sup"].Value = txtsup.Text;
            //this.databind.Rows[this.databind.Rows.Count - 1].Cells["checklist"].Value = txtchecklist.Text;
            //this.databind.Rows[this.databind.Rows.Count - 1].Cells["remarks"].Value = txtremarks.Text;

            string ssql = "select * from MaterialBitPIC where PNO = '"+mano+ "' order  by eventtime asc ";
            DataTable  dt = DbAccess.SelectBySql(ssql).Tables[0];
            if (dt != null && dt.Rows.Count > 0)
            {
                databind.DataSource = dt;
            }
           
        }
        private void clearform()
        {
            txtPNo.Text = "";
            txtname.Text = "";
            txtpiclist.Text = "";
            p1.Image = null;
            p2.Image = null;
            p3.Image = null;
            txtsup.Text = "";
            txtPackType.Text = "";
            txtchecklist.Text = "";
            txtremarks.Text = "";
        }


        public string Device_MaterialProdPicNew(string opertype, string PNo, string Pname, string userid, string sup,string PackType, string checklist, string remarks)
        {

            SqlParameter[] para = new SqlParameter[9];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@pno", PNo);
            para[2] = new SqlParameter("@PName", Pname);
            para[3] = new SqlParameter("@userid", userid);
            para[4] = new SqlParameter("@sup", sup);
            para[5] = new SqlParameter("@PackType", PackType);
            para[6] = new SqlParameter("@checklist", checklist);
            para[7] = new SqlParameter("@remarks", remarks);
            para[8] = new SqlParameter("@msg", SqlDbType.VarChar, 100);
            para[8].Direction = ParameterDirection.Output;
            DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "IQC_ProdPictureAddNew", para);
            return para[8].Value.ToString();
        }


        private void btnsave_Click(object sender, EventArgs e)
        {

            if (txtPNo.Text.Trim() == "") return;

            string pno = txtPNo.Text.Trim();
            string fileDBServerPath = "";
            if (txtpiclist.Text != "")
            {
                string[] fileqty = txtpiclist.Text.TrimEnd(',').Split(',');
                foreach (string s in fileqty)
                {
                    string[] fileName = @s.Split('\\');
                    fileDBServerPath += "\\\\" + this.serverFilePath + "\\丝印图" + "\\" + this.txtPNo.Text.Trim() + "\\" + fileName[fileName.Length - 1] + ",";
                }
            }

            string floerPath = "";
            floerPath = "\\\\" + this.serverFilePath + "\\丝印图" + "\\" + this.txtPNo.Text.Trim();

            //string msg = ic.Device_MaterialProdPic("新增", txtPNo.Text, txtname.Text, Login.username, txtsup.Text, txtchecklist.Text, txtremarks.Text);
            string msg = Device_MaterialProdPicNew("新增", txtPNo.Text, txtname.Text, Login.username, txtsup.Text,txtPackType.Text, txtchecklist.Text, txtremarks.Text);
            if (msg.IndexOf("上传成功") >= 0)
            {

                if (floerPath != "")
                    del_prefile(floerPath);

                UploadFile(fileDBServerPath);
                bindIQCProdPicInfo(Mano);
                clearform();
                txtPNo.Focus();
            }
            else if (msg.IndexOf("更新成功") >= 0)
            {
                if (txtpiclist.Text != "")
                {
                    if (floerPath != "")
                        del_prefile(floerPath);

                    UploadFile(fileDBServerPath);
                    //clearform();
                    //txtPNo.Focus();
                }
                clearform();
                txtPNo.Focus();
            }
            else
                MessageBox.Show(msg, "提醒", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void btndel_Click(object sender, EventArgs e)
        {
            // if (databind.SelectedRows.Count < 1) return;
            if (gridView.FocusedRowHandle < 0)
                return;
            if (gridView.GetSelectedRows().Length < 0)
                return;
            int j = 0;
            // //for (int i = gridView.GetSelectedRows().Length; i > 0; i--)
            // //{
            // gridView.GetDataRow(i - 1)  已存在的记录   gridView.GetFocusedRowCellValue("Remarks").ToString()
            for (int i = gridView.GetSelectedRows().Length; i > 0; i--)
            {
                DataRow dr = gridView.GetDataRow(gridView.GetSelectedRows()[i - 1]);

                string msg = ic.Device_MaterialProdPic("删除", dr["PNo"].ToString(), "", Login.username, "", "", "");
                if (msg.IndexOf("成功") >= 0)
                {
                    string floerPath = "";
                    floerPath = "\\\\" + this.serverFilePath + "\\丝印图" + "\\" + this.txtPNo.Text.Trim();
                    if (floerPath != "")
                        del_prefile(floerPath);

                    p1.Image = null;
                    p2.Image = null;
                    p3.Image = null;

                    gridView.DeleteRow(gridView.FocusedRowHandle);

                    j++;
                }
            }




            /*
            string msg = Device_MaterialProdPicNew("删除", gridView.GetFocusedRowCellValue("PNo").ToString(), gridView.GetFocusedRowCellValue("PName").ToString(), Login.username,
                                gridView.GetFocusedRowCellValue("sup").ToString(), gridView.GetFocusedRowCellValue("checklist").ToString(), gridView.GetFocusedRowCellValue("Remarks").ToString());
            if (msg.IndexOf("已存在的记录") >= 0)
            {
                p1.Image = null;
                p2.Image = null;
                p3.Image = null;
                gridView.DeleteRow(gridView.FocusedRowHandle);
                MessageBox.Show("删除已存在的记录成功！", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                clearform();
         
            }
            else if (msg.IndexOf("成功") >= 0)
            {
                string floerPath = "";
                floerPath = "\\\\" + this.serverFilePath + "\\丝印图" + "\\" + this.txtPNo.Text.Trim();
                if (floerPath != "")
                    del_prefile(floerPath);
                p1.Image = null;
                p2.Image = null;
                p3.Image = null;
                gridView.DeleteRow(gridView.FocusedRowHandle);
              
            }
            else
            {
                MessageBox.Show("删除失败！", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
         
            */
        }

        private void btnsearch_Click(object sender, EventArgs e)
        {
            // databind.Rows.Clear();
            if (txtPNo.Text.Trim() == "")
                return;
            databind.DataSource = null;
            string ssql = "	select top 1 * from MaterialBitPIC where PNo = '"+ txtPNo.Text.Trim() + "' order by eventtime desc  ";
            DataSet ds = DbAccess.SelectBySql(ssql);
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                databind.DataSource = ds.Tables[0];
            }
            else
            {
                MessageBox.Show("没有符合条件的记录", "提醒");
            }
        }

        private void btnallsearch_Click(object sender, EventArgs e)
        {
            // databind.Rows.Clear();
            //if (txtPNo.Text.Trim() == "")
            //    return;
            databind.DataSource = null;

            string where = " where 1=1 ";
            string PNo = txtPNo.Text.Trim();
            string checklist = txtchecklist.Text;

            if (!string.IsNullOrEmpty(PNo))
            {
                where += " and PNO = '" + PNo + "' ";
            }
            if (!string.IsNullOrEmpty(checklist))
            {
                where += " and checklist = '" + checklist + "' ";
            }

            string ssql = "select * from MaterialBitPIC ";
            ssql += where + "  order by eventtime desc  ";

            DataSet ds = DbAccess.SelectBySql(ssql);
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                databind.DataSource = ds.Tables[0];


                //int i = ds.Tables[0].Rows.Count;
                //string show = "";
                //txtinformation.Text = " 品牌商:" + ds.Tables[0].Rows[0]["sup"].ToString() + ",检查项:" + ds.Tables[0].Rows[0]["checklist"].ToString() + ",备注:" + ds.Tables[0].Rows[0]["Remarks"].ToString();
                //for (int j = 1; j < i; j++)
                //{
                //    show = show + "\r\n 【" + ds.Tables[0].Rows[j]["sup"].ToString() + "；" + ds.Tables[0].Rows[j]["checklist"].ToString() + "；" + ds.Tables[0].Rows[j]["Remarks"].ToString() + "】 \r\n";
                //}
                //txtinformation.Text = txtinformation.Text + show;
                //txtinformation.ForeColor = Color.Blue;


            }
            else
            {
                MessageBox.Show("没有符合条件的记录", "提醒");
            }

        }



        private void OpenFile(string filepath, string pdffile)
        {
            string filename = "";
            //filename = pdffile + ".bmp";
            filename = pdffile;
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
        private void p1_DoubleClick(object sender, EventArgs e)
        {
            string floerPath = "\\\\" + this.serverFilePath + "\\丝印图" + "\\" + txtPNo.Text;
            if (Directory.Exists(floerPath))
            {
                string[] pt = Directory.GetFiles(floerPath);
                pt = pt.Where(s => !s.EndsWith("Thumbs.db")).ToArray();
                if (pt.Length == 0)
                    return;
                string filename = Path.GetFileName(pt[0].ToString());
                OpenFile(floerPath, filename);
            }
            else
            {
                string pold = "";
                string sqlold = "select pold from UAT_ITEMNEW where Pnew='" + txtPNo.Text + "'";
                DataTable dtold = DbAccess.SelectBySql(sqlold).Tables[0];
                if (dtold.Rows.Count > 0)
                {
                    pold = dtold.Rows[0][0].ToString();
                    string floerPathold = "\\\\" + this.serverFilePath + "\\丝印图" + "\\" + pold;
                    if (Directory.Exists(floerPathold))
                    {
                        string[] pt = Directory.GetFiles(floerPathold);
                        pt = pt.Where(s => !s.EndsWith("Thumbs.db")).ToArray();
                        if (pt.Length == 0)
                            return;
                        string filename = Path.GetFileName(pt[0].ToString());
                        OpenFile(floerPathold, filename);
                    }
                }
            }
        }
        public struct videohdr_tag
        {
            public byte[] lpData;
            public int dwBufferLength;
            public int dwBytesUsed;
            public int dwTimeCaptured;
            public int dwUser;
            public int dwFlags;
            public int[] dwReserved;

        }
        public delegate bool CallBack(int hwnd, int lParam);
        //private System.ComponentModel.Container components = null;
        [DllImport("avicap32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int capCreateCaptureWindowA([MarshalAs(UnmanagedType.VBByRefStr)]   ref string lpszWindowName, int dwStyle, int x, int y, int nWidth, short nHeight, int hWndParent, int nID);
        [DllImport("avicap32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern bool capGetDriverDescriptionA(short wDriver, [MarshalAs(UnmanagedType.VBByRefStr)]   ref string lpszName, int cbName, [MarshalAs(UnmanagedType.VBByRefStr)]   ref string lpszVer, int cbVer);
        [DllImport("user32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern bool DestroyWindow(int hndw);
        [DllImport("user32", EntryPoint = "SendMessageA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int SendMessage(int hwnd, int wMsg, int wParam, [MarshalAs(UnmanagedType.AsAny)]   object lParam);
        [DllImport("user32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int SetWindowPos(int hwnd, int hWndInsertAfter, int x, int y, int cx, int cy, int wFlags);
        [DllImport("vfw32.dll")]
        public static extern string capVideoStreamCallback(int hwnd, videohdr_tag videohdr_tag);
        [DllImport("vicap32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern bool capSetCallbackOnFrame(int hwnd, string s);

        private void OpenCapture()
        {

            int intWidth = this.panel_Vedio.Width;
            int intHeight = this.panel_Vedio.Height;
            int intDevice = 0;
            string refDevice = intDevice.ToString();
            hHwnd = IQCProdPic.capCreateCaptureWindowA(ref refDevice, 1342177280, 0, 0, 640, 480, this.panel_Vedio.Handle.ToInt32(), 0);
            if (IQCProdPic.SendMessage(hHwnd, 0x40a, intDevice, 0) > 0)
            {
                IQCProdPic.SendMessage(this.hHwnd, 0x435, -1, 0);
                IQCProdPic.SendMessage(this.hHwnd, 0x434, 0x42, 0);
                IQCProdPic.SendMessage(this.hHwnd, 0x432, -1, 0);
                IQCProdPic.SetWindowPos(this.hHwnd, 1, 0, 0, intWidth, intHeight, 6);
                this.BtnCapTure.Enabled = false;
                this.BtnStop.Enabled = true;
            }
            else
            {
                IQCProdPic.DestroyWindow(this.hHwnd);
                this.BtnCapTure.Enabled = false;
                this.BtnStop.Enabled = true;
            }
        }
        private void BtnCapTure_Click(object sender, EventArgs e)
        {
            btnphoto.Enabled = true;
            this.OpenCapture();
        }

        private void BtnStop_Click(object sender, EventArgs e)
        {
            IQCProdPic.SendMessage(this.hHwnd, 0x40b, 0, 0);
            IQCProdPic.DestroyWindow(this.hHwnd);
            this.BtnCapTure.Enabled = true;
            btnphoto.Enabled = false;
            this.BtnStop.Enabled = false;
        }

        private void btnphoto_Click(object sender, EventArgs e)
        {
            try
            {
                IQCProdPic.SendMessage(this.hHwnd, 0x41e, 0, 0);
                IDataObject obj1 = Clipboard.GetDataObject();
                if (obj1.GetDataPresent(typeof(Bitmap)))
                {
                    Image image1 = (Image)obj1.GetData(typeof(Bitmap));

                    SaveFileDialog SaveFileDialog1 = new SaveFileDialog();
                    SaveFileDialog1.FileName = DateTime.Now.ToString("yyyyMMddhhmmss");
                    SaveFileDialog1.Filter = "Image Files(*.JPG;*.GIF)|*.JPG;*.GIF|All files (*.*)|*.*";
                    if (SaveFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        image1.Save(SaveFileDialog1.FileName, ImageFormat.Bmp);
                    }
                }
            }
            catch
            {
            }
        }

        private void databind_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //DataGridView dgv = (DataGridView)sender;
            ////如果是"PNo"列，被点击
            ////if (dgv.Columns[e.ColumnIndex].Name == "PNo")
            //{
            //    txtPNo.Text = databind.CurrentRow.Cells["PNo"].Value.ToString();
            //    txtname.Text = databind.CurrentRow.Cells["PName"].Value.ToString();
            //    txtsup.Text = databind.CurrentRow.Cells["sup"].Value.ToString();
            //    txtchecklist.Text = databind.CurrentRow.Cells["checklist"].Value.ToString();
            //    txtremarks.Text = databind.CurrentRow.Cells["remarks"].Value.ToString();

            //    p1.Image = null;
            //    p2.Image = null;
            //    p3.Image = null;
            //    Cursor.Current = Cursors.WaitCursor;
            //    string pno = databind.CurrentRow.Cells["PNo"].Value.ToString();

            //    string floerPath = "\\\\" + this.serverFilePath + "\\丝印图" + "\\" + pno;
            //    if (Connect(serverFilePath))
            //    {
            //        if (Directory.Exists(floerPath))
            //        {
            //            string[] pt = Directory.GetFiles(floerPath);

            //            pt = pt.Where(s => !s.EndsWith("Thumbs.db")).ToArray();

            //            if (pt.Length == 0) return;
            //            if (pt.Length == 1)
            //            {
            //                FileStream fs = new FileStream(pt[0].ToString(), FileMode.Open);
            //                Image bt = Image.FromStream(fs);
            //                fs.Close();
            //                fs.Dispose();
            //                p1.Image = bt;
            //                this.p1.SizeMode = PictureBoxSizeMode.Zoom;
            //            }
            //            else if (pt.Length == 2)
            //            {
            //                FileStream fs = new FileStream(pt[0].ToString(), FileMode.Open);
            //                Image bt = Image.FromStream(fs);
            //                fs.Close();
            //                fs.Dispose();
            //                p1.Image = bt;
            //                this.p1.SizeMode = PictureBoxSizeMode.Zoom;

            //                FileStream fs2 = new FileStream(pt[1].ToString(), FileMode.Open);
            //                Image bt2 = Image.FromStream(fs2);
            //                fs2.Close();
            //                fs2.Dispose();
            //                p2.Image = bt2;
            //                this.p2.SizeMode = PictureBoxSizeMode.Zoom;
            //            }
            //            else if (pt.Length == 3)
            //            {
            //                FileStream fs = new FileStream(pt[0].ToString(), FileMode.Open);
            //                Image bt = Image.FromStream(fs);
            //                fs.Close();
            //                fs.Dispose();
            //                p1.Image = bt;
            //                this.p1.SizeMode = PictureBoxSizeMode.Zoom;

            //                FileStream fs2 = new FileStream(pt[1].ToString(), FileMode.Open);
            //                Image bt2 = Image.FromStream(fs2);
            //                fs2.Close();
            //                fs2.Dispose();
            //                p2.Image = bt2;
            //                this.p2.SizeMode = PictureBoxSizeMode.Zoom;

            //                FileStream fs3 = new FileStream(pt[2].ToString(), FileMode.Open);
            //                Image bt3 = Image.FromStream(fs3);
            //                fs3.Close();
            //                fs3.Dispose();
            //                p3.Image = bt3;
            //                this.p3.SizeMode = PictureBoxSizeMode.Zoom;
            //            }
            //        }
            //        else
            //        {
            //            string pold = "";
            //            string sqlold = "select pold from UAT_ITEMNEW where Pnew='" + pno + "'";
            //            DataTable dtold = DbAccess.SelectBySql(sqlold).Tables[0];
            //            if (dtold.Rows.Count > 0)
            //            {
            //                pold = dtold.Rows[0][0].ToString();
            //                string floerPathold = "\\\\" + this.serverFilePath + "\\丝印图" + "\\" + pold;
            //                if (Directory.Exists(floerPathold))
            //                {
            //                    string[] pt = Directory.GetFiles(floerPathold);

            //                    pt = pt.Where(s => !s.EndsWith("Thumbs.db")).ToArray();

            //                    if (pt.Length == 0) return;
            //                    if (pt.Length == 1)
            //                    {
            //                        FileStream fs = new FileStream(pt[0].ToString(), FileMode.Open);
            //                        Image bt = Image.FromStream(fs);
            //                        fs.Close();
            //                        fs.Dispose();
            //                        p1.Image = bt;
            //                        this.p1.SizeMode = PictureBoxSizeMode.Zoom;
            //                    }
            //                    else if (pt.Length == 2)
            //                    {
            //                        FileStream fs = new FileStream(pt[0].ToString(), FileMode.Open);
            //                        Image bt = Image.FromStream(fs);
            //                        fs.Close();
            //                        fs.Dispose();
            //                        p1.Image = bt;
            //                        this.p1.SizeMode = PictureBoxSizeMode.Zoom;

            //                        FileStream fs2 = new FileStream(pt[1].ToString(), FileMode.Open);
            //                        Image bt2 = Image.FromStream(fs2);
            //                        fs2.Close();
            //                        fs2.Dispose();
            //                        p2.Image = bt2;
            //                        this.p2.SizeMode = PictureBoxSizeMode.Zoom;
            //                    }
            //                    else if (pt.Length == 3)
            //                    {
            //                        FileStream fs = new FileStream(pt[0].ToString(), FileMode.Open);
            //                        Image bt = Image.FromStream(fs);
            //                        fs.Close();
            //                        fs.Dispose();
            //                        p1.Image = bt;
            //                        this.p1.SizeMode = PictureBoxSizeMode.Zoom;

            //                        FileStream fs2 = new FileStream(pt[1].ToString(), FileMode.Open);
            //                        Image bt2 = Image.FromStream(fs2);
            //                        fs2.Close();
            //                        fs2.Dispose();
            //                        p2.Image = bt2;
            //                        this.p2.SizeMode = PictureBoxSizeMode.Zoom;

            //                        FileStream fs3 = new FileStream(pt[2].ToString(), FileMode.Open);
            //                        Image bt3 = Image.FromStream(fs3);
            //                        fs3.Close();
            //                        fs3.Dispose();
            //                        p3.Image = bt3;
            //                        this.p3.SizeMode = PictureBoxSizeMode.Zoom;
            //                    }
            //                }
            //            }
            //        }
            //    }
            //    Cursor.Current = Cursors.Default;
            //}
        }

        private void gridView_Click(object sender, EventArgs e)
        {
           // //txtPNo.Text = databind.CurrentRow.Cells["PNo"].Value.ToString();
           // //txtname.Text = databind.CurrentRow.Cells["PName"].Value.ToString();
           // //txtsup.Text = databind.CurrentRow.Cells["sup"].Value.ToString();
           // //txtchecklist.Text = databind.CurrentRow.Cells["checklist"].Value.ToString();
           // //txtremarks.Text = databind.CurrentRow.Cells["remarks"].Value.ToString();
           // DataTable de = databind.DataSource as DataTable;
           // if (de == null || de.Rows.Count < 1)
           //     return;
           // txtPNo.Text = gridView.GetFocusedRowCellValue("PNo").ToString();
           // txtname.Text = gridView.GetFocusedRowCellValue("PName").ToString();
           // txtsup.Text = gridView.GetFocusedRowCellValue("sup").ToString();
           // txtchecklist.Text = gridView.GetFocusedRowCellValue("checklist").ToString();
           // txtremarks.Text = gridView.GetFocusedRowCellValue("Remarks").ToString();
           // p1.Image = null;
           // p2.Image = null;
           // p3.Image = null;
           // Cursor.Current = Cursors.WaitCursor;
           //// string pno = databind.CurrentRow.Cells["PNo"].Value.ToString();
           // string pno = gridView.GetFocusedRowCellValue("PNo").ToString();

           // string floerPath = "\\\\" + this.serverFilePath + "\\丝印图" + "\\" + pno;
           // if (Connect(serverFilePath))
           // {
           //     if (Directory.Exists(floerPath))
           //     {
           //         string[] pt = Directory.GetFiles(floerPath);

           //         pt = pt.Where(s => !s.EndsWith("Thumbs.db")).ToArray();

           //         if (pt.Length == 0)
           //             return;
           //         if (pt.Length == 1)
           //         {
           //             FileStream fs = new FileStream(pt[0].ToString(), FileMode.Open);
           //             Image bt = Image.FromStream(fs);
           //             fs.Close();
           //             fs.Dispose();
           //             p1.Image = bt;
           //             p1.Properties.SizeMode = PictureSizeMode.Stretch;
           //         }
           //         else if (pt.Length == 2)
           //         {
           //             FileStream fs = new FileStream(pt[0].ToString(), FileMode.Open);
           //             Image bt = Image.FromStream(fs);
           //             fs.Close();
           //             fs.Dispose();
           //             p1.Image = bt;
           //             p1.Properties.SizeMode = PictureSizeMode.Stretch;

           //             FileStream fs2 = new FileStream(pt[1].ToString(), FileMode.Open);
           //             Image bt2 = Image.FromStream(fs2);
           //             fs2.Close();
           //             fs2.Dispose();
           //             p2.Image = bt2;
           //             p2.Properties.SizeMode = PictureSizeMode.Stretch;
           //         }
           //         else if (pt.Length == 3)
           //         {
           //             FileStream fs = new FileStream(pt[0].ToString(), FileMode.Open);
           //             Image bt = Image.FromStream(fs);
           //             fs.Close();
           //             fs.Dispose();
           //             p1.Image = bt;
           //             p1.Properties.SizeMode = PictureSizeMode.Stretch;

           //             FileStream fs2 = new FileStream(pt[1].ToString(), FileMode.Open);
           //             Image bt2 = Image.FromStream(fs2);
           //             fs2.Close();
           //             fs2.Dispose();
           //             p2.Image = bt2;
           //             p2.Properties.SizeMode = PictureSizeMode.Stretch;

           //             FileStream fs3 = new FileStream(pt[2].ToString(), FileMode.Open);
           //             Image bt3 = Image.FromStream(fs3);
           //             fs3.Close();
           //             fs3.Dispose();
           //             p3.Image = bt3;
           //             p3.Properties.SizeMode = PictureSizeMode.Stretch;
           //         }
           //     }
           //     else
           //     {
           //         string pold = "";
           //         string sqlold = "select pold from UAT_ITEMNEW where Pnew='" + pno + "'";
           //         DataTable dtold = DbAccess.SelectBySql(sqlold).Tables[0];
           //         if (dtold.Rows.Count > 0)
           //         {
           //             pold = dtold.Rows[0][0].ToString();
           //             string floerPathold = "\\\\" + this.serverFilePath + "\\丝印图" + "\\" + pold;
           //             if (Directory.Exists(floerPathold))
           //             {
           //                 string[] pt = Directory.GetFiles(floerPathold);

           //                 pt = pt.Where(s => !s.EndsWith("Thumbs.db")).ToArray();

           //                 if (pt.Length == 0)
           //                     return;
           //                 if (pt.Length == 1)
           //                 {
           //                     FileStream fs = new FileStream(pt[0].ToString(), FileMode.Open);
           //                     Image bt = Image.FromStream(fs);
           //                     fs.Close();
           //                     fs.Dispose();
           //                     p1.Image = bt;
           //                     p1.Properties.SizeMode = PictureSizeMode.Stretch;
           //                 }
           //                 else if (pt.Length == 2)
           //                 {
           //                     FileStream fs = new FileStream(pt[0].ToString(), FileMode.Open);
           //                     Image bt = Image.FromStream(fs);
           //                     fs.Close();
           //                     fs.Dispose();
           //                     p1.Image = bt;
           //                     p1.Properties.SizeMode = PictureSizeMode.Stretch;

           //                     FileStream fs2 = new FileStream(pt[1].ToString(), FileMode.Open);
           //                     Image bt2 = Image.FromStream(fs2);
           //                     fs2.Close();
           //                     fs2.Dispose();
           //                     p2.Image = bt2;
           //                     p2.Properties.SizeMode = PictureSizeMode.Stretch;
           //                 }
           //                 else if (pt.Length == 3)
           //                 {
           //                     FileStream fs = new FileStream(pt[0].ToString(), FileMode.Open);
           //                     Image bt = Image.FromStream(fs);
           //                     fs.Close();
           //                     fs.Dispose();
           //                     p1.Image = bt;
           //                     p1.Properties.SizeMode = PictureSizeMode.Stretch;

           //                     FileStream fs2 = new FileStream(pt[1].ToString(), FileMode.Open);
           //                     Image bt2 = Image.FromStream(fs2);
           //                     fs2.Close();
           //                     fs2.Dispose();
           //                     p2.Image = bt2;
           //                     p2.Properties.SizeMode = PictureSizeMode.Stretch;

           //                     FileStream fs3 = new FileStream(pt[2].ToString(), FileMode.Open);
           //                     Image bt3 = Image.FromStream(fs3);
           //                     fs3.Close();
           //                     fs3.Dispose();
           //                     p3.Image = bt3;
           //                     p3.Properties.SizeMode = PictureSizeMode.Stretch;
           //                 }
           //             }
           //         }
           //     }
           // }
           // Cursor.Current = Cursors.Default;
        }

        private void txtPNo_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13)
            {
                txtPNo_Leave(sender, e);
            }
        }

        private void gridView_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {
            DataTable de = databind.DataSource as DataTable;
            if (de == null || de.Rows.Count < 1)
                return;
            if (gridView.RowCount < 1)
                return;      
            txtPNo.Text = gridView.GetFocusedRowCellValue("PNo").ToString();
            txtname.Text = gridView.GetFocusedRowCellValue("PName").ToString();
            txtsup.Text = gridView.GetFocusedRowCellValue("sup").ToString();
            txtPackType.Text = gridView.GetFocusedRowCellValue("PackType").ToString();
            txtchecklist.Text = gridView.GetFocusedRowCellValue("checklist").ToString();
            txtremarks.Text = gridView.GetFocusedRowCellValue("Remarks").ToString();
            p1.Image = null;
            p2.Image = null;
            p3.Image = null;
            Cursor.Current = Cursors.WaitCursor;
            // string pno = databind.CurrentRow.Cells["PNo"].Value.ToString();
            string pno = gridView.GetFocusedRowCellValue("PNo").ToString();

            string floerPath = "\\\\" + this.serverFilePath + "\\丝印图" + "\\" + pno;
            if (Connect(serverFilePath))
            {
                if (Directory.Exists(floerPath))
                {
                    string[] pt = Directory.GetFiles(floerPath);

                    pt = pt.Where(s => !s.EndsWith("Thumbs.db")).ToArray();

                    if (pt.Length == 0)
                        return;
                    if (pt.Length == 1)
                    {
                        FileStream fs = new FileStream(pt[0].ToString(), FileMode.Open);
                        Image bt = Image.FromStream(fs);
                        fs.Close();
                        fs.Dispose();
                        p1.Image = bt;
                        p1.Properties.SizeMode = PictureSizeMode.Stretch;
                    }
                    else if (pt.Length == 2)
                    {
                        FileStream fs = new FileStream(pt[0].ToString(), FileMode.Open);
                        Image bt = Image.FromStream(fs);
                        fs.Close();
                        fs.Dispose();
                        p1.Image = bt;
                        p1.Properties.SizeMode = PictureSizeMode.Stretch;

                        FileStream fs2 = new FileStream(pt[1].ToString(), FileMode.Open);
                        Image bt2 = Image.FromStream(fs2);
                        fs2.Close();
                        fs2.Dispose();
                        p2.Image = bt2;
                        p2.Properties.SizeMode = PictureSizeMode.Stretch;
                    }
                    else if (pt.Length == 3)
                    {
                        FileStream fs = new FileStream(pt[0].ToString(), FileMode.Open);
                        Image bt = Image.FromStream(fs);
                        fs.Close();
                        fs.Dispose();
                        p1.Image = bt;
                        p1.Properties.SizeMode = PictureSizeMode.Stretch;

                        FileStream fs2 = new FileStream(pt[1].ToString(), FileMode.Open);
                        Image bt2 = Image.FromStream(fs2);
                        fs2.Close();
                        fs2.Dispose();
                        p2.Image = bt2;
                        p2.Properties.SizeMode = PictureSizeMode.Stretch;

                        FileStream fs3 = new FileStream(pt[2].ToString(), FileMode.Open);
                        Image bt3 = Image.FromStream(fs3);
                        fs3.Close();
                        fs3.Dispose();
                        p3.Image = bt3;
                        p3.Properties.SizeMode = PictureSizeMode.Stretch;
                    }
                }
                else
                {
                    string pold = "";
                    string sqlold = "select pold from UAT_ITEMNEW where Pnew='" + pno + "'";
                    DataTable dtold = DbAccess.SelectBySql(sqlold).Tables[0];
                    if (dtold.Rows.Count > 0)
                    {
                        pold = dtold.Rows[0][0].ToString();
                        string floerPathold = "\\\\" + this.serverFilePath + "\\丝印图" + "\\" + pold;
                        if (Directory.Exists(floerPathold))
                        {
                            string[] pt = Directory.GetFiles(floerPathold);

                            pt = pt.Where(s => !s.EndsWith("Thumbs.db")).ToArray();

                            if (pt.Length == 0)
                                return;
                            if (pt.Length == 1)
                            {
                                FileStream fs = new FileStream(pt[0].ToString(), FileMode.Open);
                                Image bt = Image.FromStream(fs);
                                fs.Close();
                                fs.Dispose();
                                p1.Image = bt;
                                p1.Properties.SizeMode = PictureSizeMode.Stretch;
                            }
                            else if (pt.Length == 2)
                            {
                                FileStream fs = new FileStream(pt[0].ToString(), FileMode.Open);
                                Image bt = Image.FromStream(fs);
                                fs.Close();
                                fs.Dispose();
                                p1.Image = bt;
                                p1.Properties.SizeMode = PictureSizeMode.Stretch;

                                FileStream fs2 = new FileStream(pt[1].ToString(), FileMode.Open);
                                Image bt2 = Image.FromStream(fs2);
                                fs2.Close();
                                fs2.Dispose();
                                p2.Image = bt2;
                                p2.Properties.SizeMode = PictureSizeMode.Stretch;
                            }
                            else if (pt.Length == 3)
                            {
                                FileStream fs = new FileStream(pt[0].ToString(), FileMode.Open);
                                Image bt = Image.FromStream(fs);
                                fs.Close();
                                fs.Dispose();
                                p1.Image = bt;
                                p1.Properties.SizeMode = PictureSizeMode.Stretch;

                                FileStream fs2 = new FileStream(pt[1].ToString(), FileMode.Open);
                                Image bt2 = Image.FromStream(fs2);
                                fs2.Close();
                                fs2.Dispose();
                                p2.Image = bt2;
                                p2.Properties.SizeMode = PictureSizeMode.Stretch;

                                FileStream fs3 = new FileStream(pt[2].ToString(), FileMode.Open);
                                Image bt3 = Image.FromStream(fs3);
                                fs3.Close();
                                fs3.Dispose();
                                p3.Image = bt3;
                                p3.Properties.SizeMode = PictureSizeMode.Stretch;
                            }
                        }
                    }
                }
            }
            Cursor.Current = Cursors.Default;
        }
    }
}