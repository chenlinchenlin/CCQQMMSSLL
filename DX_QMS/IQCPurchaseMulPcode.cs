using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.IO;
using System.Windows.Forms;
using DevExpress.XtraBars;
using DX_QMS.Common;

namespace DX_QMS
{
    public partial class IQCPurchaseMulPcode : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        [DllImport("user32.dll")]
        public static extern int WindowFromPoint(int xPoint, int yPoint);

        private int hHwnd;
        private string serverFilePath = "";

        public IQCPurchaseMulPcode()
        {
            InitializeComponent();
            setRule();
            string path = System.Configuration.ConfigurationManager.AppSettings["ServerFilePath"].ToString();
            this.serverFilePath = @path;
            txttype.SelectedIndex = 0;
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
            Dictionary<string, bool> dic = GroupPermission.QMS_SelectRulesForForm(post, "图片样");
            this.btnsave.Enabled = bool.Parse(dic["hasInsert"].ToString());
            this.btndel.Enabled = bool.Parse(dic["hasDelete"].ToString());
        }
        void IQCPurchaseMulPcode_MouseWheel(object sender, MouseEventArgs e)
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
        private void IQCPurchaseMulPcode_Load(object sender, EventArgs e)
        {
            this.MouseWheel += new MouseEventHandler(IQCPurchaseMulPcode_MouseWheel);
        }

        private void txtlotno_KeyUp(object sender, KeyEventArgs e)
        {
            if (txtlotno.Text == "") return;
            if (e.KeyValue == 13)
                txtlotno_Leave(sender, null);
        }
        private DataSet bindData(string productcode)
        {
            string ssql = "select productcode 产品编码,pname 描述,mtype 是否多件,psubname 附件,operdate 维护日期,operuser 维护人 from IQC_PurchaseMulPcode where productcode like '" + productcode + "%'";
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
                lblinfo.Text = txtproductcode.Text + "正确,请添加子项";
                lblinfo.ForeColor = Color.Blue;
                txtpurchasedesc.Text = "";
                txtpurchasedesc.Focus();
                txtpurchasep.Text = txtproductcode.Text;
                 DataTable dce = bindData(txtproductcode.Text).Tables[0];
                databind.DataSource = dce;
                lblinfo2.Text = txtproductcode.Text + "总共有子项目数:" + dce.Rows.Count;
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
                    lblinfo.Text = txtproductcode.Text + "正确,请添加子项";
                    lblinfo.ForeColor = Color.Blue;
                    txtpurchasedesc.Text = "";
                    txtpurchasedesc.Focus();
                    txtpurchasep.Text = txtproductcode.Text;
                    DataTable des = bindData(txtproductcode.Text).Tables[0];
                    databind.DataSource = des;
                    lblinfo2.Text = txtproductcode.Text + "总共有子项目数:" + des.Rows.Count;
                    lblinfo2.ForeColor = Color.Blue;
                    txtlotno.Leave += txtlotno_Leave;
                    return;
                }
            }
            else
            {
                lblinfo.Text = txtlotno.Text + "不正确,请重新输入";
                lblinfo.ForeColor = Color.Red;
                txtpurchasedesc.Text = "";
                txtlotno.SelectAll();
                txtlotno.Enabled = true;
                txtlotno.Focus();
            }

            txtlotno.Leave += txtlotno_Leave;
        }

        private void txtpurchasep_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13)
                txtpurchasedesc.Focus();
        }

        private void txtpurchasedesc_KeyUp(object sender, KeyEventArgs e)
        {
            if (txtpurchasedesc.Text.Trim() == "") return;
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
                if (txtpiclist.Text != "" && !txtpiclist.Text.Contains("pdf"))
                {
                    string[] pt = @txtpiclist.Text.TrimEnd(',').Split(',');
                    if (pt.Length == 1)
                    {
                        FileStream fs = new FileStream(pt[0].ToString(), FileMode.Open);
                        Image bt = Image.FromStream(fs);
                        fs.Close();
                        fs.Dispose();
                        p1.Image = bt;
                        this.p1.SizeMode = PictureBoxSizeMode.Zoom;
                    }
                }
            }
        }
        IQCProdPic ic = new IQCProdPic();
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
        private void UploadFile(string filepath)
        {

            if (ic.Connect(serverFilePath))
            {
                bool re = true;
                string msg = "";
                string te = txtpiclist.Text.TrimEnd(',');
                if (te != "")
                {
                    string[] copy = te.Split(',');
                    foreach (string ss in copy)
                    {
                        if (!CopyFileToServer("外购件", ss))
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
        protected void del_prefile(string filepath, string filename)
        {
            if (Directory.Exists(filepath))
            {
                if (ic.Connect(serverFilePath))
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
        private void btnsave_Click(object sender, EventArgs e)
        {
            if (txtproductcode.Text.Trim() == "") return;

            string fileDBServerPath = "";
            if (txtpiclist.Text != "")
            {
                string filename = Path.GetFileNameWithoutExtension(txtpiclist.Text);
                if (filename != txtpurchasep.Text.Trim())
                {
                    MessageBox.Show("上传的文件名:" + filename + "与子项目号:" + txtpurchasep.Text + "不一致");
                    return;
                }
                string[] fileqty = txtpiclist.Text.TrimEnd(',').Split(',');
                foreach (string s in fileqty)
                {
                    string[] fileName = @s.Split('\\');
                    fileDBServerPath += "\\\\" + this.serverFilePath + "\\外购件" + "\\" + this.txtproductcode.Text.Trim() + "\\" + fileName[fileName.Length - 1] + ",";
                }
            }
            string floerPath = "";
            floerPath = "\\\\" + this.serverFilePath + "\\外购件" + "\\" + this.txtproductcode.Text.Trim();


            if (this.txtproductcode.Text == "" || txtpurchasep.Text == "" || txtpurchasedesc.Text == "")
                return;
            string sql = "if not exists(select 1 from IQC_PurchaseMulPcode where productcode='" + txtproductcode.Text.Trim() + "' and item='" + txtpurchasep.Text + "'" + ")";
            sql += " insert into IQC_PurchaseMulPcode(productcode,mtype,pname,item,psubname,operuser,operdate) values('" + txtproductcode.Text + "','" + txttype.Text + "','";
            sql += txtproductdesc.Text.Replace("'", "") + "','" + txtpurchasep.Text + "','" + txtpurchasedesc.Text.Replace("'", "") + "','" + Login.username + "',getdate())";
            if (DbAccess.ExecuteSql(sql))
            {
                if (floerPath != "")
                    del_prefile(floerPath, txtpurchasep.Text);
                UploadFile(fileDBServerPath);

                DataTable des = bindData(txtproductcode.Text).Tables[0];

                databind.DataSource =des;
                lblinfo2.Text = txtproductcode.Text + "总共有子项目数:" + des.Rows.Count;
                lblinfo2.ForeColor = Color.Blue;
                txtpurchasedesc.Text = "";
                txtpurchasep.Text = "";
                txtpiclist.Text = "";
                txtpurchasep.Focus();
            }
            else
            {
                MessageBox.Show("新增失败," + txtproductcode.Text + "," + txtpurchasep.Text + "该编码已存在!");
                txtpurchasedesc.Text = "";
                txtpurchasep.Text = "";
                txtpurchasep.Focus();
            }
        }

        private void btndel_Click(object sender, EventArgs e)
        {
          //  if (databind.SelectedRows.Count <= 0) return;

            if (gridView.FocusedRowHandle < 0)
                return;
            if (gridView.GetSelectedRows().Length < 0)
                return;

            string sql = "";
            for (int i = 0; i < gridView.GetSelectedRows().Length; i++)
            {
                //sql = "delete IQC_PurchaseMulPcode where productcode='" + databind.SelectedRows[i].Cells["产品编码"].Value.ToString() + "' and item='" + databind.SelectedRows[i].Cells["子项目编码"].Value.ToString() + "'";
                sql = "delete IQC_PurchaseMulPcode where productcode='" + gridView.GetDataRow(gridView.GetSelectedRows()[i])["产品编码"].ToString() + "' and item='" + gridView.GetDataRow(gridView.GetSelectedRows()[i])["产品编码"].ToString() + "'";
                DbAccess.ExecuteSql(sql);

                string floerPath = "";
                floerPath = "\\\\" + this.serverFilePath + "\\外购件" + "\\" + gridView.GetDataRow(gridView.GetSelectedRows()[i])["产品编码"].ToString();
                string filename = gridView.GetDataRow(gridView.GetSelectedRows()[i])["产品编码"].ToString();
                del_prefile(floerPath, filename);
            }
            DataTable dse = bindData(txtproductcode.Text).Tables[0];
            int k = dse.Rows.Count;
            databind.DataSource = dse;
            lblinfo2.Text = txtproductcode.Text + "总共有子项目数:" + k.ToString();
            lblinfo2.ForeColor = Color.Blue;
            txtpurchasep.Focus();
        }

        private void btnreset_Click(object sender, EventArgs e)
        {
            txtpurchasep.Text = "";
            txtpurchasedesc.Text = "";
            txtlotno.Text = "";
            txtproductcode.Text = "";
            txtproductdesc.Text = "";
            txtpiclist.Text = "";
            txtlotno.Enabled = true;
            txtlotno.Focus();
        }

        private void btnsearch_Click(object sender, EventArgs e)
        {
            DataSet ds = bindData(txtlotno.Text);
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                databind.DataSource = ds.Tables[0];
                lblinfo2.Text = txtproductcode.Text + "总共有子项目数:" + ds.Tables[0].Rows.Count;
                lblinfo2.ForeColor = Color.Blue;
            }
            else
                MessageBox.Show("没有符合条件的记录");
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
            hHwnd = IQCPurchaseMulPcode.capCreateCaptureWindowA(ref refDevice, 1342177280, 0, 0, 640, 480, this.panel_Vedio.Handle.ToInt32(), 0);
            if (IQCPurchaseMulPcode.SendMessage(hHwnd, 0x40a, intDevice, 0) > 0)
            {
                IQCPurchaseMulPcode.SendMessage(this.hHwnd, 0x435, -1, 0);
                IQCPurchaseMulPcode.SendMessage(this.hHwnd, 0x434, 0x42, 0);
                IQCPurchaseMulPcode.SendMessage(this.hHwnd, 0x432, -1, 0);
                IQCPurchaseMulPcode.SetWindowPos(this.hHwnd, 1, 0, 0, intWidth, intHeight, 6);
                this.BtnCapTure.Enabled = false;
                this.BtnStop.Enabled = true;
            }
            else
            {
                IQCPurchaseMulPcode.DestroyWindow(this.hHwnd);
                this.BtnCapTure.Enabled = false;
                this.BtnStop.Enabled = true;
            }
        }
        private void BtnCapTure_Click(object sender, EventArgs e)
        {
            this.OpenCapture();
        }

        private void BtnStop_Click(object sender, EventArgs e)
        {
            IQCPurchaseMulPcode.SendMessage(this.hHwnd, 0x40b, 0, 0);
            IQCPurchaseMulPcode.DestroyWindow(this.hHwnd);
            this.BtnCapTure.Enabled = true;
            this.BtnStop.Enabled = false;
        }

        private void btnphoto_Click(object sender, EventArgs e)
        {
            try
            {
                IQCPurchaseMulPcode.SendMessage(this.hHwnd, 0x41e, 0, 0);
                IDataObject obj1 = Clipboard.GetDataObject();
                if (obj1.GetDataPresent(typeof(Bitmap)))
                {
                    Image image1 = (Image)obj1.GetData(typeof(Bitmap));
                    SaveFileDialog SaveFileDialog1 = new SaveFileDialog();
                    SaveFileDialog1.FileName = txtpurchasep.Text;
                    SaveFileDialog1.Filter = "Image Files(*.JPG;*.GIF)|*.JPG;*.GIF|All files (*.*)|*.*";
                    if (SaveFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        image1.Save(SaveFileDialog1.FileName, ImageFormat.Jpeg);
                    }
                }
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
        private void databind_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //this.txtpurchasedesc.Text = databind.CurrentRow.Cells["附件"].Value.ToString();
            //string pno = databind.CurrentRow.Cells["产品编码"].Value.ToString();
            //string item = databind.CurrentRow.Cells["产品编码"].Value.ToString();
            //string floerPath = "\\\\" + this.serverFilePath + "\\外购件" + "\\" + pno;
            //if (ic.Connect(serverFilePath))
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

            //                if (filename == item)
            //                {
            //                    FileStream fs = new FileStream(pt[i].ToString(), FileMode.Open);
            //                    BinaryReader reader = new BinaryReader(fs);
            //                    string fileclass = "";
            //                    try
            //                    {
            //                        for (int j = 0; j < 2; j++)
            //                        {
            //                            fileclass += reader.ReadByte().ToString();
            //                        }

            //                    }
            //                    catch (Exception)
            //                    {

            //                        throw;
            //                    }

            //                    if (fileclass == "3780")
            //                    {
            //                        OpenFile(floerPath, filename);
            //                        fs.Close();
            //                        fs.Dispose();
            //                        reader.Close();
            //                        return;
            //                    }

            //                    Image bt = Image.FromStream(fs);
            //                    fs.Close();
            //                    fs.Dispose();
            //                    p1.Image = bt;
            //                    this.p1.SizeMode = PictureBoxSizeMode.Zoom;
            //                    break;
            //                }
            //            }
            //        }
            //    }
            //    else
            //    {
            //        string pnoold = "";
            //        string sqlold = "select pold from UAT_ITEMNEW where Pnew='" + pno + "'";
            //        DataTable dtold = DbAccess.SelectBySql(sqlold).Tables[0];
            //        if (dtold.Rows.Count > 0)
            //        {
            //            pnoold = dtold.Rows[0][0].ToString();
            //            string floerPathold = "\\\\" + this.serverFilePath + "\\外购件" + "\\" + pnoold;
            //            if (Directory.Exists(floerPathold))
            //            {
            //                string[] pt = Directory.GetFiles(floerPathold);
            //                if (pt.Length == 0) return;
            //                if (pt.Length > 0)
            //                {
            //                    for (int i = 0; i < pt.Length; i++)
            //                    {
            //                        string filename = Path.GetFileNameWithoutExtension(pt[i].ToString());

            //                        if (filename == pnoold)
            //                        {
            //                            FileStream fs = new FileStream(pt[i].ToString(), FileMode.Open);
            //                            BinaryReader reader = new BinaryReader(fs);
            //                            string fileclass = "";
            //                            try
            //                            {
            //                                for (int j = 0; j < 2; j++)
            //                                {
            //                                    fileclass += reader.ReadByte().ToString();
            //                                }

            //                            }
            //                            catch (Exception)
            //                            {

            //                                throw;
            //                            }

            //                            if (fileclass == "3780")
            //                            {
            //                                OpenFile(floerPathold, filename);
            //                                fs.Close();
            //                                fs.Dispose();
            //                                reader.Close();
            //                                return;
            //                            }

            //                            Image bt = Image.FromStream(fs);
            //                            fs.Close();
            //                            fs.Dispose();
            //                            p1.Image = bt;
            //                            this.p1.SizeMode = PictureBoxSizeMode.Zoom;
            //                            break;
            //                        }
            //                    }
            //                }
            //            }
            //        }
            //    }
            //}
        }

        private void gridView_DoubleClick(object sender, EventArgs e)
        {
            //this.txtpurchasedesc.Text = databind.CurrentRow.Cells["附件"].Value.ToString();
            //string pno = databind.CurrentRow.Cells["产品编码"].Value.ToString();
            //string item = databind.CurrentRow.Cells["产品编码"].Value.ToString();
            this.txtpurchasedesc.Text = gridView.GetFocusedRowCellValue("附件").ToString();
            string pno = gridView.GetFocusedRowCellValue("产品编码").ToString();
            string item = gridView.GetFocusedRowCellValue("产品编码").ToString();
            string floerPath = "\\\\" + this.serverFilePath + "\\外购件" + "\\" + pno;
            if (ic.Connect(serverFilePath))
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

                            if (filename == item)
                            {
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

                                if (fileclass == "3780")
                                {
                                    OpenFile(floerPath, filename);
                                    fs.Close();
                                    fs.Dispose();
                                    reader.Close();
                                    return;
                                }

                                Image bt = Image.FromStream(fs);
                                fs.Close();
                                fs.Dispose();
                                p1.Image = bt;
                                this.p1.SizeMode = PictureBoxSizeMode.Zoom;
                                break;
                            }
                        }
                    }
                }
                else
                {
                    string pnoold = "";
                    string sqlold = "select pold from UAT_ITEMNEW where Pnew='" + pno + "'";
                    DataTable dtold = DbAccess.SelectBySql(sqlold).Tables[0];
                    if (dtold.Rows.Count > 0)
                    {
                        pnoold = dtold.Rows[0][0].ToString();
                        string floerPathold = "\\\\" + this.serverFilePath + "\\外购件" + "\\" + pnoold;
                        if (Directory.Exists(floerPathold))
                        {
                            string[] pt = Directory.GetFiles(floerPathold);
                            if (pt.Length == 0) return;
                            if (pt.Length > 0)
                            {
                                for (int i = 0; i < pt.Length; i++)
                                {
                                    string filename = Path.GetFileNameWithoutExtension(pt[i].ToString());

                                    if (filename == pnoold)
                                    {
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

                                        if (fileclass == "3780")
                                        {
                                            OpenFile(floerPathold, filename);
                                            fs.Close();
                                            fs.Dispose();
                                            reader.Close();
                                            return;
                                        }

                                        Image bt = Image.FromStream(fs);
                                        fs.Close();
                                        fs.Dispose();
                                        p1.Image = bt;
                                        this.p1.SizeMode = PictureBoxSizeMode.Zoom;
                                        break;
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