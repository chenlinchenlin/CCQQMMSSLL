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
using Microsoft.Win32;
using System.IO.Ports;
using System.Runtime.InteropServices;
using System.Data.SqlClient;
using System.Data.OracleClient;
using DX_QMS.Common;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using DevExpress.XtraEditors.Controls;
using DX_QMS.IQCFilePosition;

namespace DX_QMS
{
    public partial class IQCTestReturn : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private string tranid = "", item_id = "", org_id = "";
        private string srecept = "", productcode_1 = "";
        private string checktype = "NGToOK";
        [DllImport("user32.dll", EntryPoint = "FindWindow", CharSet = CharSet.Auto)]
        private extern static IntPtr FindWindow(string lpClassName, string lpWindowName);
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int PostMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
        public const int WM_CLOSE = 0x10;
        [DllImport("user32.dll")]
        public static extern int WindowFromPoint(int xPoint, int yPoint);
        private static SerialPort ssp = new SerialPort();
        public delegate void HandleInterfaceUpdateDelegate(string text);
        private HandleInterfaceUpdateDelegate updateSmallBox;
        Queue<float> valQueue = new Queue<float>();
        Queue<float> valQueue2 = new Queue<float>();
        bool sBoxEnable = false;
        int errorid = 0;
        int testcount = 0;
        public DataTable dttitem = null, dtalready = new DataTable();
        private string serverFilePath = "";
        public int lotqty = 0, checkcycle = 0;
        public string testtype = "";
        public string AQLValue = "", sproductcode = "", sCheckType = "";
        string StrSql = "", MaterialState ="";
        IQC ic = new IQC();
        public IQCTestReturn()
        {
            InitializeComponent();
            //InitializeComponent();
            string path = System.Configuration.ConfigurationSettings.AppSettings["ServerFilePath"].ToString();
            this.serverFilePath = @path;

            dtalready.Columns.Add("id", typeof(string));
            dtalready.Columns.Add("qty", typeof(int));

            ini_com_list();
            setRule();
            GetRSNO();

          
            //labelControl1.Enabled = false;
            //labelControl1.Visible = false;
            Badcategory.Enabled = false;
            Badcategory.Visible = false;

        }               
        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



        DataTable BadSituation()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("不良类别", System.Type.GetType("System.String"));
            dt.Columns.Add("不良现象", System.Type.GetType("System.String"));
            string[] bad = { "包装", "标签", "数量", "错料", "外观", "结构尺寸", "试装不良", "功能", "环保", "其它" };
            string[] situation =
                {
                "散料、漏气、不符合ESD、MSL包装要求",
                 "五要素缺失、超期、批次、版本、型号",
                 "叉板超标、多装、短装、打叉板未减数、少配件等不良情况",
                 "与BOM描述不一致、混料",
                 "划痕、丝印、色差、引脚、打点、漏工序等",
                 "长宽厚、平面度、翘曲",
                 "螺孔打不到底，造成滑牙，或者装配间隙大，无法匹配，造成卡塞或者装不进去",
                 "可靠性测试（丝印粘着力测试、低温跌落实验，48H盐雾等）、阻容感值、发光测试",
                 "RoHS",
                 "材质问题或其他出现问题"
                 };
            for (int i = 0; i < 10; i++)
            {
                dt.Rows.Add(bad[i], situation[i]);
            }

            return dt;
        }

        private void ini_com_list()
        {
            RegistryKey keyCom = Registry.LocalMachine.OpenSubKey("Hardware\\DeviceMap\\SerialComm");
            if (keyCom != null)
            {
                string[] sSubKeys = keyCom.GetValueNames();

                this.edtCom.Properties.Items.Clear();
                foreach (string sName in sSubKeys)
                {
                    string sValue = (string)keyCom.GetValue(sName);
                    this.edtCom.Properties.Items.Add(sValue);
                }
                edtCom.SelectedIndex = 0;
            }
            updateSmallBox = new HandleInterfaceUpdateDelegate(updateSmallBoxInfo);
            ssp.DataReceived += new SerialDataReceivedEventHandler(sspDataReceived);
            ssp.RtsEnable = true;
        }

        private void tuiliaojianyan_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (ssp.IsOpen)
                ssp.Close();
            this.Dispose();
            this.Close();
        }
        void TestListReturn_MouseWheel(object sender, MouseEventArgs e)
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
            if (p2.Image != null)
            {
                if (WindowFromPoint(p.X, p.Y) == p2.Handle.ToInt32())
                {
                    //向前
                    if (e.Delta > 0)
                    {
                        float w = this.p2.Width * 0.9f; //每次縮小 20%  
                        float h = this.p2.Height * 0.9f;
                        this.p2.Size = Size.Ceiling(new SizeF(w, h));

                    }

                    //向后
                    else if (e.Delta < 0)
                    {

                        float w = this.p2.Width * 1.1f; //每次放大 20%
                        float h = this.p2.Height * 1.1f;
                        this.p2.Size = Size.Ceiling(new SizeF(w, h));
                        p2.Invalidate();

                    }
                }
            }
        }
        private void xtraTabControl1_Click(object sender, EventArgs e)
        {
            this.MouseWheel += new MouseEventHandler(TestListReturn_MouseWheel);
        }
        public DataTable dtpub = null;
        public DataTable dtother = null;
        public DataTable dtEMS = null;
        public DataSet dsinfo = null;
        int m = 0;
        private void getInfo(string lotno)
        {
            this.dsinfo = null;
            //20150210新增,清空表,以免下一检验批次为有接收单号的批次,而没有清空杂项检验的批次信息
            this.dtother = null;
            m = 0;
            string sSql = "select deliveryid,materialcode,qty,vendorname,lot_number from delivery  where lotno='" + lotno + "'";
            DataSet ds = DbAccess.SelectBySql(sSql);
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                string sDeliveryID = ds.Tables[0].Rows[0]["deliveryid"].ToString();

                string ssql = "select max(deliveryid) deliveryid ,d.materialcode,sum(qty) qty,max(vendorname) vendorname,max(INVENTORY_ITEM_ID) INVENTORY_ITEM_ID,max(unit) unit,max(org_id) org_id,'' lot_number from delivery d inner join materialspec m on d.materialcode=m.materialcode where deliveryid='" + ds.Tables[0].Rows[0]["deliveryid"].ToString() + "' and d.materialcode='" + ds.Tables[0].Rows[0]["materialcode"].ToString() + "'  group by d.materialcode ";
                DataSet dsREC = DbAccess.SelectBySql(ssql);
                dsinfo = dsREC;
                dtother = dsREC.Tables[0];
                if (dtother.Rows.Count > 0)
                    m = int.Parse(dtother.Rows[0]["qty"].ToString());
            }
            else
            {
                lblinfo.Text = txtlotno.Text + ":批次号错误或不存在!";
                lblinfo.ForeColor = Color.Red;
                MessageBox.Show(lblinfo.Text, "提示", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                txtlotno.Text = "";
                txtlotno.Focus();
            }
        }

        public void sspDataReceived(object sender, System.IO.Ports.SerialDataReceivedEventArgs e)
        {
            byte[] readBuffer = new byte[ssp.BytesToRead];
            ssp.Read(readBuffer, 0, readBuffer.Length);
            this.Invoke(updateSmallBox, new object[] { Encoding.UTF8.GetString(readBuffer) });
        }
        private void StartKiller()
        {
            Timer timer = new Timer();
            timer.Interval = 3000; //3秒启动   
            timer.Tick += new EventHandler(Timer_Tick);
            timer.Start();
        }
        private void Timer_Tick(object sender, EventArgs e)
        {
            KillMessageBox();
            ((Timer)sender).Stop();
        }
        private void KillMessageBox()
        {
            IntPtr ptr = FindWindow(null, "MessageBox");
            if (ptr != IntPtr.Zero)
            {
                PostMessage(ptr, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
            }

        }
        private void updateSmallBoxInfo(string text)
        {

            if (text.Length < 30 || (text.Substring(0, 1) != "{" && text.Substring(29, 1) != "}"))
            {
                errorid++;
                ssp.ReadExisting();
                ssp.ReadTo("}");
                Application.DoEvents();
                return;
            }
            string unit = "";
            string T = text.Substring(1, 1).Trim();
            string u = text.Substring(26, 1).Trim();
            if (T == "0")
            {
                if (u == "0")
                    unit = "uH";
                else if (u == "1")
                    unit = "mH";
                else if (u == "2")
                    unit = "H";
            }
            else if (T == "1")
            {
                if (u == "0")
                    unit = "pF";
                else if (u == "1")
                    unit = "nF";
                else if (u == "2")
                    unit = "uF";
            }
            else if (T == "2" || T == "3")
            {
                if (u == "0")
                    unit = "Ω";
                else if (u == "1")
                    unit = "KΩ";
                else if (u == "2")
                    unit = "MΩ";
            }
            lblunit.Text = unit;
            lblunit.ForeColor = Color.Green;
            lblTestValueShow.Text = (text.Substring(14, 6).Trim());
            string tval = text.Substring(14, 6).Trim();
            if (txtlotno.Text.Trim() == "") return;
            float n = 0;
            if (float.TryParse(tval, out n))
            {
                if (float.Parse(tval) <= 0)
                {
                    lblTestValueShow.ForeColor = Color.Red;
                    return;
                }
                float lowlittle = float.Parse(txtscope1.Text == "" ? "0" : txtscope1.Text) * (1 - (float.Parse(txtupscope.Text == "" ? "0" : txtupscope.Text) / 100));
                float upbig = float.Parse(txtscope2.Text == "" ? "0" : txtscope2.Text) * (1 + (float.Parse(txtupscope.Text == "" ? "0" : txtupscope.Text) / 100));
                if (float.Parse(tval) < lowlittle)
                {
                    lblTestValueShow.ForeColor = Color.Red;
                    return;
                }
                if (float.Parse(tval) > upbig)
                {
                    lblTestValueShow.ForeColor = Color.Red;
                    return;
                }
                if (txtsampleqty.Text == txtsamplefactqty.Text)
                {
                    return;
                }
                if ((float.Parse(tval) >= float.Parse(txtscope1.Text == "" ? "0" : txtscope1.Text)) && (float.Parse(tval) <= float.Parse(txtscope2.Text == "" ? "0" : txtscope2.Text)))
                {
                    lblTestValueShow.ForeColor = Color.Green;
                    valQueue.Enqueue(float.Parse(tval));
                    valQueue2.Clear();
                    testcount++;

                    if (valQueue.Count == 5)
                    {
                        txttestvalue.Text = valQueue.Last().ToString();
                        cbOK.Checked = true;
                        cbOK.Enabled = false;
                        cbNG.Enabled = false;
                        StartKiller();
                        if (MessageBox.Show("测试值:" + txttestvalue.Text + ",检验结果:" + cbOK.Text, "MessageBox", MessageBoxButtons.OK, MessageBoxIcon.Information) == DialogResult.Cancel)
                        {
                            return;
                        }
                        if (btnOK.Enabled)
                            btnOK_Click(null, null);
                        valQueue.Clear();
                        testcount = 0;
                    }
                }
                else if ((float.Parse(tval) >= lowlittle && float.Parse(tval) <= float.Parse(txtscope1.Text == "" ? "0" : txtscope1.Text))
                          || (float.Parse(tval) >= float.Parse(txtscope2.Text == "" ? "0" : txtscope2.Text) && float.Parse(tval) <= upbig))
                {
                    lblTestValueShow.ForeColor = Color.YellowGreen;
                    valQueue2.Enqueue(float.Parse(tval));
                    valQueue.Clear();
                    testcount++;

                    if (valQueue2.Count == 5)
                    {
                        txttestvalue.Text = valQueue2.Max().ToString();

                        cbOK.Checked = false;
                        cbOK.Enabled = false;
                        cbNG.Checked = true;
                        cbNG.Enabled = false;
                        StartKiller();
                        if (MessageBox.Show("测试值:" + txttestvalue.Text + ",检验结果:" + cbOK.Text, "MessageBox", MessageBoxButtons.OK, MessageBoxIcon.Warning) == DialogResult.Cancel)
                        {
                            return;
                        }
                        if (btnOK.Enabled)
                            btnOK_Click(null, null);
                        valQueue2.Clear();
                        testcount = 0;
                    }
                }
                else if (testcount > 10)
                {
                    lblTestValueShow.ForeColor = Color.Red;
                    valQueue.Clear();
                    valQueue2.Clear();
                    txttestvalue.Text = "0";
                    MessageBox.Show("测试不成功,请检查测量值范围或校准后再测!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    testcount = 0;
                    txttestvalue.Focus();
                    cbOK.Checked = false;
                    cbOK.Enabled = false;
                    cbNG.Checked = false;
                    cbNG.Enabled = false;
                }
            }
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
            Dictionary<string, bool> dic = GroupPermission.QMS_SelectRulesForForm(post, "退料检验");
            btninstock.Enabled = bool.Parse(dic["hasInsert"].ToString());
            btninstock.Enabled = false;
        }
        private void GetRSNO()
        {
            string sql = "select '' Definevalue union select Definevalue from  OQC_TypeDefine where Definetype='资产编号'";
            DataTable dt = Common.DbAccess.SelectBySql(sql).Tables[0];
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow row in dt.Rows)
                {
                    txtrsno.Properties.Items.Add(row["Definevalue"]);
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
                //string dosLine = @"net use \\" + remoteHost + " " + passWord + " " + " /user:" + userName + ">NUL";
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
        private void showpic(string pcode)
        {
            Cursor.Current = Cursors.WaitCursor;

            string pno = pcode;

            string floerPath = "\\\\" + this.serverFilePath + "\\丝印图" + "\\" + pno;
            if (Connect(serverFilePath))
            {
                if (Directory.Exists(floerPath))
                {
                    string[] pt = Directory.GetFiles(floerPath);
                    if (pt.Length == 0) return;
                    try
                    {
                        if (pt.Length == 1)
                        {
                            FileStream fs = new FileStream(pt[0].ToString(), FileMode.Open);
                            Image bt = Image.FromStream(fs);
                            fs.Close();
                            fs.Dispose();
                            p1.Image = bt;
                            //this.p1.SizeMode = PictureBoxSizeMode.Zoom;
                            p1.Properties.SizeMode = PictureSizeMode.Stretch;

                            p2.Image = null;
                            p3.Image = null;
                        }
                        else if (pt.Length == 2)
                        {
                            FileStream fs = new FileStream(pt[0].ToString(), FileMode.Open);
                            Image bt = Image.FromStream(fs);
                            fs.Close();
                            fs.Dispose();
                            p1.Image = bt;
                           //this.p1.SizeMode = PictureBoxSizeMode.Zoom;
                            p1.Properties.SizeMode = PictureSizeMode.Stretch;

                            FileStream fs2 = new FileStream(pt[1].ToString(), FileMode.Open);
                            Image bt2 = Image.FromStream(fs2);
                            fs2.Close();
                            fs2.Dispose();
                            p2.Image = bt2;
                            //this.p2.SizeMode = PictureBoxSizeMode.Zoom;
                            p2.Properties.SizeMode = PictureSizeMode.Zoom;
                            p3.Image = null;
                        }
                        else if (pt.Length == 3)
                        {
                            FileStream fs = new FileStream(pt[0].ToString(), FileMode.Open);
                            Image bt = Image.FromStream(fs);
                            fs.Close();
                            fs.Dispose();
                            p1.Image = bt;
                            //this.p1.SizeMode = PictureBoxSizeMode.Zoom;
                            p1.Properties.SizeMode = PictureSizeMode.Stretch;

                            FileStream fs2 = new FileStream(pt[1].ToString(), FileMode.Open);
                            Image bt2 = Image.FromStream(fs2);
                            fs2.Close();
                            fs2.Dispose();
                            p2.Image = bt2;
                            //this.p2.SizeMode = PictureBoxSizeMode.Zoom;
                            p2.Properties.SizeMode = PictureSizeMode.Zoom;

                            FileStream fs3 = new FileStream(pt[2].ToString(), FileMode.Open);
                            Image bt3 = Image.FromStream(fs3);
                            fs3.Close();
                            fs3.Dispose();
                            p3.Image = bt3;
                            //this.p3.SizeMode = PictureBoxSizeMode.Zoom;
                            p3.Properties.SizeMode = PictureSizeMode.Zoom;
                        }
                    }
                    catch { }
                }
                else
                {
                    string s = "select pold from UAT_ITEMNEW where Pnew='" + pcode + "'";
                    DataTable dtold = DbAccess.SelectBySql(s).Tables[0];
                    if (dtold.Rows.Count > 0)
                    {
                        pno = dtold.Rows[0][0].ToString();
                        string floerPathold = "\\\\" + this.serverFilePath + "\\丝印图" + "\\" + pno;
                        if (Directory.Exists(floerPathold))
                        {
                            string[] pt = Directory.GetFiles(floerPathold);
                            if (pt.Length == 0) return;
                            try
                            {
                                if (pt.Length == 1)
                                {
                                    FileStream fs = new FileStream(pt[0].ToString(), FileMode.Open);
                                    Image bt = Image.FromStream(fs);
                                    fs.Close();
                                    fs.Dispose();
                                    p1.Image = bt;
                                    //this.p1.SizeMode = PictureBoxSizeMode.Zoom;
                                    p1.Properties.SizeMode = PictureSizeMode.Stretch;

                                    p2.Image = null;
                                    p3.Image = null;
                                }
                                else if (pt.Length == 2)
                                {
                                    FileStream fs = new FileStream(pt[0].ToString(), FileMode.Open);
                                    Image bt = Image.FromStream(fs);
                                    fs.Close();
                                    fs.Dispose();
                                    p1.Image = bt;
                                    //this.p1.SizeMode = PictureBoxSizeMode.Zoom;
                                    p1.Properties.SizeMode = PictureSizeMode.Stretch;

                                    FileStream fs2 = new FileStream(pt[1].ToString(), FileMode.Open);
                                    Image bt2 = Image.FromStream(fs2);
                                    fs2.Close();
                                    fs2.Dispose();
                                    p2.Image = bt2;
                                    //this.p2.SizeMode = PictureBoxSizeMode.Zoom;
                                    p2.Properties.SizeMode = PictureSizeMode.Zoom;
                                    p3.Image = null;
                                }
                                else if (pt.Length == 3)
                                {
                                    FileStream fs = new FileStream(pt[0].ToString(), FileMode.Open);
                                    Image bt = Image.FromStream(fs);
                                    fs.Close();
                                    fs.Dispose();
                                    p1.Image = bt;
                                    //this.p1.SizeMode = PictureBoxSizeMode.Zoom;
                                    p1.Properties.SizeMode = PictureSizeMode.Stretch;

                                    FileStream fs2 = new FileStream(pt[1].ToString(), FileMode.Open);
                                    Image bt2 = Image.FromStream(fs2);
                                    fs2.Close();
                                    fs2.Dispose();
                                    p2.Image = bt2;
                                   // this.p2.SizeMode = PictureBoxSizeMode.Zoom;
                                    p2.Properties.SizeMode = PictureSizeMode.Zoom;

                                    FileStream fs3 = new FileStream(pt[2].ToString(), FileMode.Open);
                                    Image bt3 = Image.FromStream(fs3);
                                    fs3.Close();
                                    fs3.Dispose();
                                    p3.Image = bt3;
                                    //this.p3.SizeMode = PictureBoxSizeMode.Zoom;
                                    p3.Properties.SizeMode = PictureSizeMode.Zoom;
                                }
                            }
                            catch { }
                        }
                        else
                        {
                            p1.Image = null;
                            p2.Image = null;
                            p3.Image = null;
                        }
                    }
                    else
                    {
                        p1.Image = null;
                        p2.Image = null;
                        p3.Image = null;
                    }
                }
                //string ssql = "select PNo, PName, userid, eventtime, sup, checklist, Remarks from MaterialBitPIC where PNo='" + pcode + "'";
                //DataSet ds = DbAccess.SelectBySql(ssql);
                //if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                //{
                //    lblpicdes.Text = "制造商:" + ds.Tables[0].Rows[0]["sup"].ToString() + ",检验项目:" + ds.Tables[0].Rows[0]["checklist"].ToString() + ",备注:" + ds.Tables[0].Rows[0]["Remarks"].ToString();
                //    lblpicdes.ForeColor = Color.Blue;
                //}
                //else
                //{
                //    lblpicdes.Text = "";
                //}
                //Cursor.Current = Cursors.Default;

                string ssql = "select PNo, PName, userid, eventtime, sup, checklist, Remarks from MaterialBitPIC where PNo='" + pcode + "' order  by eventtime asc";
                DataSet ds = Common.DbAccess.SelectBySql(ssql);
                if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                {
                    int i = ds.Tables[0].Rows.Count;
                    string show = "";
                    lblpicdes.Text = "品牌商:" + ds.Tables[0].Rows[0]["sup"].ToString() + ",检验项目:" + ds.Tables[0].Rows[0]["checklist"].ToString() + ",备注:" + ds.Tables[0].Rows[0]["Remarks"].ToString();
                    for (int j = 1; j < i; j++)
                    {
                        show = show + " 【" + ds.Tables[0].Rows[j]["sup"].ToString() + "；" + ds.Tables[0].Rows[j]["checklist"].ToString() + "；" + ds.Tables[0].Rows[j]["Remarks"].ToString() + "】 ";
                    }
                    lblpicdes.Text = lblpicdes.Text + show;
                    lblpicdes.ForeColor = Color.Blue;

                }
                else
                {
                    lblpicdes.Text = "";
                }
                Cursor.Current = Cursors.Default;

            }
        }
        private void showsampledirectory(string pcode)
        {
            IQC ic = new IQC();
            DataSet ds = ic.SelectTestSamplePosition("查询", pcode, "", "", "", "", "", "", "", "", "", "", "", Login.username,"");
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                lbldirectory.Text = "【" + ds.Tables[0].Rows[0]["状态"].ToString() + "】,样品位置:" + ds.Tables[0].Rows[0]["存放位置"].ToString() + ",图纸位置:" + ds.Tables[0].Rows[0]["图纸位置"].ToString();
                if (ds.Tables[0].Rows[0]["状态"].ToString() == "正常")
                    lbldirectory.ForeColor = Color.Green;
                else
                    lbldirectory.ForeColor = Color.Red;
            }
            else
            {
                lbldirectory.Text = "";
            }
        }
        private void showRevExcept(string pcode)
        {
            string sql = "select productcode, pname, supplier, Reason, operuser, operdate from IQC_RevExcept where productcode='" + pcode + "'";
            DataTable dt = Common.DbAccess.SelectBySql(sql).Tables[0];
            if (dt.Rows.Count > 0)
            {
                DialogResult drt = MessageBox.Show(pcode + ":有【" + dt.Rows.Count.ToString() + "】项不良现象,请去查看", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                if (DialogResult.OK == drt)
                {
                    IQCRevExcept_2 IQCE = new IQCRevExcept_2(pcode);
                    IQCE.ShowDialog();
                }
            }
        }
        private string IfCheck(string productcode)
        {
            string Flag = "不需审核";
            string sql = "select min(case when IFYiQi='否' then '不需审核' else ISNULL(states,'未审') end) states from  IQC_TestProgSet s inner join  IQC_TestType i on s.TestType=i.TestType where Productcode='" + productcode + "'";
            DataSet ds = Common.DbAccess.SelectBySql(sql);
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
        private string GetF1Supp(string pcode)
        {
            string s = "";
            string sqlORA = "select MANUFACTURER_NAME,MFG_PART_NUM from apps.CUX_MTL_MANUFACTURERS_v where ITEM_NUMBER='" + pcode + "'";
            DataTable dt = Common.DbAccess.SelectByOracle(sqlORA).Tables[0];
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    s += dt.Rows[i]["MANUFACTURER_NAME"].ToString() + " " + dt.Rows[i]["MFG_PART_NUM"].ToString() + ";";
                }
            }
            return s;
        }
        private void TestListbind(string testtype, string testitem, string testsubitem, string packtype, string lotno, string sampleqty)
        {
          //  DataSet ds = ic.SelectTestList("退料查询", testtype, testitem, testsubitem, "", "", packtype, "", 0, 0, 0, lotno, 0, Login.username, 0, 1, "", "", "", "", "", 0, "", "", "", "", "", "", "");


            DataSet ds= ic.SelectTestListNew("退料查询", testtype, testitem, testsubitem,"", "", packtype, "", 0, 0, 0,lotno, 0, Login.username, 0, 1, "", "", "", "", "", 0, "", "", "", "", "", "", "", "", "", "", 0, "");


            databind.DataSource = ds.Tables[0];
            {
                txtstockqty.Text = "";
                txttestvalue.Focus();
                txttestvalue.Enabled = true;
                txtremarks.Enabled = true;
                btnOK.Enabled = true;
            }
        }
        private string showAQLRcvalue(string AQLvalue)
        {
            string Reqty = "0";
            string ssql = "select AQLLevel, AQL, AQLValue, s.Code,c.Sampleqty, Ac, Re from IQC_TestAQLRcSet s left join IQC_TestSTD105ECode c on s.Code=c.Code where AQLValue='" + AQLvalue + "'";
            DataSet ds = Common.DbAccess.SelectBySql(ssql);
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                Reqty = ds.Tables[0].Rows[0]["Re"].ToString();
            }
            return Reqty;
        }



        private string IfNoCheck(string testtype, string testitem, string testsubitem, string productcode)
        {
            string Flag = "";
            DataSet ds = ic.IfNorequireCheck(testtype, testitem, testsubitem, productcode);
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                //表示超过了检验周期,需要重新检测
                if (int.Parse(ds.Tables[0].Rows[0]["checkcycle"].ToString()) > 0 && int.Parse(ds.Tables[0].Rows[0]["dys"].ToString()) > 0)
                {
                    Flag = "已超周期,需重测" + ",上次测试时间" + ds.Tables[0].Rows[0]["testtime"].ToString();
                }
                else if (int.Parse(ds.Tables[0].Rows[0]["checkcycle"].ToString()) > 0 && int.Parse(ds.Tables[0].Rows[0]["dys"].ToString()) < 0)
                    Flag = "未超周期无需重测" + ds.Tables[0].Rows[0]["Testvalue"].ToString() + "," + ds.Tables[0].Rows[0]["testtime"].ToString();
                else if (int.Parse(ds.Tables[0].Rows[0]["checkcycle"].ToString()) == 0)
                    Flag = "无检验周期,要录测试值";
            }
            else
            {
                //20151123新增加
                if (testitem.Contains("可靠性") || testitem.Contains("RoHS"))
                {
                    string sql = "select Productcode from IQC_TestProgSet where Productcode='" + productcode + "' and TestType='" + testtype + "' and (TestItem like '%RoHS%' OR TestItem like '%可靠性%')";
                    DataTable dt = Common.DbAccess.SelectBySql(sql).Tables[0];
                    if (dt.Rows.Count > 0)
                    {
                        Flag = productcode + "有维护送测该项,需重测";
                    }
                }
                else
                    //表示无测试数据,此次是第一次测试
                    Flag = "首次测试，要录测试值";
            }
            return Flag;
        }
        private DataTable bindTypeSet(string testtype, string testqty)
        {
            DataTable dt = null;
            DataSet ds = ic.SelectTestTypeRecord("查询", testtype, "测试个数", "");
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                dt = ds.Tables[0];
            }
            return dt;
        }
        private void bindSubitemInfo()
        {
            if (txttestitem.Text.Trim() == null || txttestsubitem.Text.Trim() == null) return;
            DataTable dt = dttitem.Clone();
            DataRow[] arrDr = dttitem.Select("TestItem='" + txttestitem.Text + "' and TestSubItem='" + txttestsubitem.Text + "'");
            for (int i = 0; i < arrDr.Length; i++)
            {
                dt.Rows.Add(arrDr[i].ItemArray);
            }
            testtype = dt.Rows[0]["TestType"].ToString();
            lblAQLValue.Text = dt.Rows[0]["AQLValue"].ToString();
            AQLValue = dt.Rows[0]["AQLValue"].ToString();

            //lblAQLRe.Text = showAQLRcvalue(AQLValue);
            textRe.Text = showAQLRcvalue(AQLValue);
            sproductcode = dt.Rows[0]["Productcode"].ToString();
            sCheckType = dt.Rows[0]["CheckType"].ToString();

            txttestdes.Text = dt.Rows[0]["TestDesc"].ToString();
            txttesttools.Text = dt.Rows[0]["TestTool"].ToString();
            lblifyiqi.Text = dt.Rows[0]["IFYiQi"].ToString();
            txtPacktype.Text = dt.Rows[0]["PackType"].ToString();
            txtsampletype.Text = dt.Rows[0]["SampleType"].ToString();
            txtAQL.Text = dt.Rows[0]["AQL"].ToString();
            txtscope1.Text = dt.Rows[0]["LowValue"].ToString();
            txtscope2.Text = dt.Rows[0]["UpValue"].ToString();
            txtupscope.Text = dt.Rows[0]["UpScope"].ToString();
            lblsetunit.Text = dt.Rows[0]["unit"].ToString();
            int testvalueqty = int.Parse(dt.Rows[0]["testvalueqty"].ToString());
            if (txtsampletype.Text == "ISO2859-1")
            {
                string ssampleqty = @"Select case when Sampleqty>lotfactqty then lotfactqty else Sampleqty end as Sampleqty,Code,AQLValue,AQL,Ac,Re,DirectCode,CheckLevel from IQC_TestSTD105ECode c ";
                ssampleqty += "  inner join ";
                ssampleqty += " (Select " + lotqty + " lotfactqty,AQLValue,AQL,CheckLevel,Ac,Re,DirectCode from IQC_TestSTD105ECheckSet i inner join IQC_TestAQLRcSet s on i.Code=s.Code  where LotSizemin<=" + lotqty.ToString() + " and LotSizemax>=" + lotqty.ToString() + " and CheckLevel='II' and AQLValue='" + AQLValue + "') s on c.Code=s.DirectCode";

                DataSet dssampleqty = Common.DbAccess.SelectBySql(ssampleqty);
                txtsampleqty.Text = dssampleqty.Tables[0].Rows[0]["Sampleqty"].ToString();
            }
            else if (txtsampletype.Text == "全检")
            {
                string scheckqty = "SELECT count(lotno) qty,SUM(qty) totalqty from delivery where deliveryid='" + dsinfo.Tables[0].Rows[0]["deliveryid"].ToString() + "' and materialcode='" + dsinfo.Tables[0].Rows[0]["materialcode"].ToString() + "'";

                DataTable dtqty = Common.DbAccess.SelectBySql(scheckqty).Tables[0];
                txtsampleqty.Text = dtqty.Rows[0][0].ToString();
                txttotalqty.Text = dtqty.Rows[0][1].ToString();
            }
            else
            {
                txtsampleqty.Text = dt.Rows[0]["Samplevalue"].ToString();
            }


            TestListbind(testtype, txttestitem.Text, txttestsubitem.Text, txtPacktype.Text, txtlotno.Text, "");
            if (txtlotno.Text.Trim() == "")
            {
                txtlotno.Focus();
                return;
            }

            string sFlag = IfNoCheck(testtype, txttestitem.Text, txttestsubitem.Text, sproductcode);
            if (sFlag.IndexOf("无需重测") >= 0)
            {
                int j = sFlag.IndexOf("无需重测");
                int i = sFlag.IndexOf(",");

                txttestvalue.Text = sFlag.Substring(j + 4, i - (j + 4));
                txtremarks.Text = "上次测试时间为:" + sFlag.Substring(i + 1);
                txtremarks.BackColor = Color.Yellow;
                txttestvalue.Enabled = false;
                txtremarks.Enabled = false;
                //return;
            }
            else if (sFlag.IndexOf("需重测") >= 0)
            {
                MessageBox.Show(sFlag, "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txttestvalue.Text = "";
                txtremarks.Text = "";
                txttestvalue.Enabled = true;
                txtremarks.Enabled = true;
                txtlotno.Text = "";
                txtlotno.Focus();
                txtlotno.SelectAll();
                return;

            }

            DataTable dtscop = bindTypeSet("", "");
            DataRow[] rw = dtscop.Select("TestType='" + txttestitem.Text + "'");
            if (rw.Length > 0 && int.Parse(txtsampleqty.Text) > int.Parse(txtsamplefactqty.Text == "" ? "0" : txtsamplefactqty.Text))
            {
                TestValueQtyList TVQ = new TestValueQtyList(testvalueqty, txtscope1.Text, txtscope2.Text, txtsampleqty.Text, txtsamplefactqty.Text, txtlotno.Text, int.Parse(txtlotqty.Text), testtype, txttestitem.Text, sproductcode, dt, txtrsno.Text == null ? "" : txtrsno.Text);
                TVQ.ShowDialog();

                txtsamplefactqty.Text = TVQ.sfq;
                TestListbind(testtype, txttestitem.Text, txttestsubitem.Text, txtPacktype.Text, txtlotno.Text, TVQ.sfq);
            }
        }

        void showExpiryDate(string lotno)
        {
            string sql = "   ";


        }

        private void txtlotno_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13 && txtlotno.Text.Trim() != "")
            {            
                             
                if (txtlotno.Text.Trim() == "")
                {
                    databind.DataSource = null;
                   // DbAccess.SetControlEmpty(this);
                    return;
                }
                getInfo(txtlotno.Text.Trim());

                //returnreason.Text = "";
                txtreturnqty.Text = "";

                DataSet ds = new DataSet();
                DataSet dsother = dsinfo;

                if (dsother != null && dsother.Tables.Count > 0 && dsother.Tables[0].Rows.Count > 0)
                {
                    string ssql = "";
                    if (dsother.Tables[0].Rows[0]["deliveryid"].ToString().Contains("REC"))
                    {
                        ds = dsother;
                    }
                    else
                    {
                        ssql = "select materialcode,sum(qty) qty,max(vendorname) vendorname,max(lot_number) lot_number  from delivery where deliveryid='" + dsother.Tables[0].Rows[0]["deliveryid"].ToString() + "' and materialcode='" + dsother.Tables[0].Rows[0]["materialcode"].ToString() + "'  group by materialcode ";
                        ds = DbAccess.SelectBySql(ssql);
                    }
                }
                else
                {
                    lblinfo.Text = txtlotno.Text + "批次号不存在!";
                    lblinfo.ForeColor = Color.Red;
                    txtlotno.Text = "";
                    txtlotno.Focus();
                    txtlotno.SelectAll();
                    return;
                }
                if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                {
                    lotqty = int.Parse(ds.Tables[0].Rows[0]["qty"].ToString());
                    txtlotqty.Text = ds.Tables[0].Rows[0]["qty"].ToString();

                    string pre = "";
                    string ssql = "select code from OQC_TypeDefine where Definetype='测试频率' and Definevalue='" + ds.Tables[0].Rows[0]["materialcode"].ToString() + "'";
                    DataTable dtpre = DbAccess.SelectBySql(ssql).Tables[0];
                    if (dtpre.Rows.Count > 0)
                        pre = dtpre.Rows[0]["code"].ToString();

                    if (ds.Tables[0].Rows[0]["materialcode"].ToString() != sproductcode)
                    {
                        showpic(ds.Tables[0].Rows[0]["materialcode"].ToString());
                        showsampledirectory(ds.Tables[0].Rows[0]["materialcode"].ToString());
                        showRevExcept(ds.Tables[0].Rows[0]["materialcode"].ToString());
                    }                   
                    string sprog = "select Productcode,productname from IQC_TestProgSet where productcode='" + ds.Tables[0].Rows[0]["materialcode"].ToString() + "'";
                    DataSet dsprog = DbAccess.SelectBySql(sprog);
                    if (dsprog != null && dsprog.Tables.Count > 0 && dsprog.Tables[0].Rows.Count > 0)
                    {                  
                        string F = IfCheck(ds.Tables[0].Rows[0]["materialcode"].ToString());
                        if (F == "未审")
                        {
                            lblinfo.Text = ds.Tables[0].Rows[0]["materialcode"].ToString() + F + ",还没有经过SQE审核,不能检验!";
                            lblinfo.ForeColor = Color.Red;
                            txtlotno.Focus();
                            txtlotno.SelectAll();
                            txtlotno.Text = "";
                            return;
                        }
                        else if (F == "NG")
                        {
                            lblinfo.Text = ds.Tables[0].Rows[0]["materialcode"].ToString() + F + ",SQE审核为NG,需要你修改重新发送审核!";
                            lblinfo.ForeColor = Color.Red;
                            txtlotno.Focus();
                            txtlotno.SelectAll();
                            txtlotno.Text = "";
                            return;
                        }
                        else
                        {
                            lblinfo.Text = ds.Tables[0].Rows[0]["materialcode"].ToString() + F;
                            lblinfo.ForeColor = Color.Blue;
                        }

                        lblsup.Text = "供应商:" + ds.Tables[0].Rows[0]["vendorname"].ToString() + ",【频点:" + ds.Tables[0].Rows[0]["lot_number"].ToString() + "】" + "频率:" + pre; ;
                        lblsup.ForeColor = Color.Blue;
                        if (dsprog.Tables[0].Rows[0]["productname"].ToString().Contains("F1专用") || dsprog.Tables[0].Rows[0]["productname"].ToString().Contains("F1客供"))
                        {
                            lblprodinfo.Text = "编码:" + dsprog.Tables[0].Rows[0]["Productcode"].ToString() + ",描述:" + dsprog.Tables[0].Rows[0]["productname"].ToString() + GetF1Supp(dsprog.Tables[0].Rows[0]["Productcode"].ToString());
                        }
                        else
                            lblprodinfo.Text = "编码:" + dsprog.Tables[0].Rows[0]["Productcode"].ToString() + ",描述:" + dsprog.Tables[0].Rows[0]["productname"].ToString();
                        if (ds.Tables[0].Rows[0]["vendorname"].ToString().Contains("临采"))
                            lblprodinfo.ForeColor = Color.Red;
                        else
                            lblprodinfo.ForeColor = Color.Blue;

                        string stestitem = "select  t.TestType, t.TestItem, t.TestSubItem, t.TestDesc, t.TestTool, t.PackType, t.SampleType, t.UpValue, t.LowValue,t.UpScope,t.AQL, t.AQLValue,Samplevalue,Productcode,IFYiQi,item,subitem,CheckType,isnull(testvalueqty,1) testvalueqty,unit,isnull(checkcycle,0) checkcycle from  IQC_TestProgSet t ";
                        stestitem += "  left join  IQC_TestSampleType s on t.SampleType=s.SampleType left join IQC_TestType ity on ity.TestType=t.TestTool where TTypes='测试工具' and Productcode='" + dsprog.Tables[0].Rows[0]["Productcode"].ToString() + "' order by item";

                        DataSet dstestitem = Common.DbAccess.SelectBySql(stestitem);
                        if (dstestitem != null && dstestitem.Tables.Count > 0 && dstestitem.Tables[0].Rows.Count > 0)
                        {
                            dttitem = dstestitem.Tables[0];

                            DataTable dt = dstestitem.Tables[0];
                            testtype = dt.Rows[0]["TestType"].ToString();

                            DataTable newdt = dt.Clone();
                            newdt.Rows.Add(dt.Rows[0].ItemArray);
                            for (int i = 1; i < dt.Rows.Count; i++)
                            {
                                bool flag = true;
                                foreach (DataRow dr in newdt.Rows)
                                {
                                    if (dt.Rows[i]["TestItem"].ToString() == dr["TestItem"].ToString())
                                    {
                                        flag = false;
                                        continue;
                                    }
                                }
                                if (flag)
                                    newdt.Rows.Add(dt.Rows[i].ItemArray);
                            }
                            
                            txttestitem.SelectedIndexChanged -= new EventHandler(txttestitem_SelectedIndexChanged);

                            txttestitem.Properties.Items.Clear();
                            foreach (DataRow row in newdt.Rows)
                            {
                                txttestitem.Properties.Items.Add(row["TestItem"]);
                            }
                            txttestitem.SelectedIndex = 0;

                            txttestitem.SelectedIndexChanged += new EventHandler(txttestitem_SelectedIndexChanged);
                            if (txttestitem.Properties.Items.Count == 1)
                            {
                                DataTable tb = dttitem.Clone();
                                DataRow[] arrDr = dttitem.Select("TestItem='" + txttestitem.Text + "'", "subitem ASC");
                                for (int i = 0; i < arrDr.Length; i++)
                                {
                                    tb.Rows.Add(arrDr[i].ItemArray);
                                }
                                txttestsubitem.SelectedIndexChanged -= new EventHandler(txttestsubitem_SelectedIndexChanged);

                                txttestsubitem.Properties.Items.Clear();
                                foreach (DataRow row in tb.Rows)
                                {
                                    txttestsubitem.Properties.Items.Add(row["TestSubItem"]);
                                }
                                txttestsubitem.SelectedIndex = 0;

                                txttestsubitem.SelectedIndexChanged += new EventHandler(txttestsubitem_SelectedIndexChanged);
                                if (txttestsubitem.Properties.Items.Count == 1)
                                {
                                    TestListbind(testtype, txttestitem.Text, txttestsubitem.Text, txtPacktype.Text, txtlotno.Text, "");
                                    showpic(sproductcode);
                                    showsampledirectory(sproductcode);
                                }
                                else
                                {
                                    txttestsubitem.Focus();
                                    //显示第一子项目的信息
                                    bindSubitemInfo();
                                }
                            }
                            else
                            {
                                DataTable tb = dttitem.Clone();
                                DataRow[] arrDr = dttitem.Select("TestItem='" + txttestitem.Text+ "'", "subitem ASC");
                                for (int i = 0; i < arrDr.Length; i++)
                                {
                                    tb.Rows.Add(arrDr[i].ItemArray);
                                }
                                txttestsubitem.SelectedIndexChanged -= new EventHandler(txttestsubitem_SelectedIndexChanged);

                                txttestsubitem.Properties.Items.Clear();
                                foreach (DataRow row in tb.Rows)
                                {
                                    txttestsubitem.Properties.Items.Add(row["TestSubItem"]);
                                }

                                txttestsubitem.SelectedIndex = 0;
                                txttestsubitem.SelectedIndexChanged += new EventHandler(txttestsubitem_SelectedIndexChanged);
                                txttestitem.Focus();
                                lblinfo.Text = "";
                                //显示第一项目、第一子项目的信息
                                bindSubitemInfo();
                            }
                            lblinfo.Text = "";
                            txtreturnqty.Enabled = true;






                        }
                        else
                        {
                            lblinfo.Text = dsprog.Tables[0].Rows[0]["Productcode"].ToString() + "还没有维护相应的测试项目!";
                            lblinfo.ForeColor = Color.Red;
                            txtlotno.Focus();
                            txtlotno.SelectAll();
                        }
                    }
                    else
                    {
                        lblinfo.Text = ds.Tables[0].Rows[0]["materialcode"].ToString() + "还没有维护相应的测试程序!";
                        lblinfo.ForeColor = Color.Red;
                        txtlotno.Focus();
                        txtlotno.SelectAll();
                    }

                }
            }
                       
        }

        private void txtlotno_Leave(object sender, EventArgs e)
        {
             txtreturnqty.Enabled = true;
        }

        private void txttestitem_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable dt = dttitem.Clone();
            DataRow[] arrDr = dttitem.Select("TestItem='" + txttestitem.Text+ "'");
            for (int i = 0; i < arrDr.Length; i++)
            {
                dt.Rows.Add(arrDr[i].ItemArray);
            }
            txttestsubitem.SelectedIndexChanged -= new EventHandler(txttestsubitem_SelectedIndexChanged);
            txttestsubitem.Properties.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
                txttestsubitem.Properties.Items.Add(row["TestSubItem"]);
            }
            txttestsubitem.SelectedIndex = 0;

            txttestsubitem.SelectedIndexChanged += new EventHandler(txttestsubitem_SelectedIndexChanged);
            if (txttestsubitem.Properties.Items.Count == 1)
            {

                testtype = dt.Rows[0]["TestType"].ToString();
                lblAQLValue.Text = dt.Rows[0]["AQLValue"].ToString();
                AQLValue = dt.Rows[0]["AQLValue"].ToString();

                //lblAQLRe.Text = showAQLRcvalue(AQLValue);
                textRe.Text = showAQLRcvalue(AQLValue);

                sproductcode = dt.Rows[0]["Productcode"].ToString();
                sCheckType = dt.Rows[0]["CheckType"].ToString();

                txttestdes.Text = dt.Rows[0]["TestDesc"].ToString();
                txttesttools.Text = dt.Rows[0]["TestTool"].ToString();
                lblifyiqi.Text = dt.Rows[0]["IFYiQi"].ToString();
                txtPacktype.Text = dt.Rows[0]["PackType"].ToString();
                txtsampletype.Text = dt.Rows[0]["SampleType"].ToString();
                txtAQL.Text = dt.Rows[0]["AQL"].ToString();
                txtscope1.Text = dt.Rows[0]["LowValue"].ToString();
                txtscope2.Text = dt.Rows[0]["UpValue"].ToString();
                txtupscope.Text = dt.Rows[0]["UpScope"].ToString();
                lblsetunit.Text = dt.Rows[0]["unit"].ToString();
                if (txtsampletype.Text == "ISO2859-1")
                {
                    string ssampleqty = @"Select case when Sampleqty>lotfactqty then lotfactqty else Sampleqty end as Sampleqty,Code,AQLValue,AQL,Ac,Re,DirectCode,CheckLevel from IQC_TestSTD105ECode c ";
                    ssampleqty += "  inner join ";
                    ssampleqty += " (Select " + lotqty + " lotfactqty,AQLValue,AQL,CheckLevel,Ac,Re,DirectCode from IQC_TestSTD105ECheckSet i inner join IQC_TestAQLRcSet s on i.Code=s.Code  where LotSizemin<=" + lotqty.ToString() + " and LotSizemax>=" + lotqty.ToString() + " and CheckLevel='II' and AQLValue='" + AQLValue + "') s on c.Code=s.DirectCode";

                    DataSet dssampleqty = Common.DbAccess.SelectBySql(ssampleqty);
                    txtsampleqty.Text = dssampleqty.Tables[0].Rows[0]["Sampleqty"].ToString();
                                   
                }
                else if (txtsampletype.Text == "C=0")
                {
                    if (lotqty >= 500000)
                    {
                        lotqty = 500000;
                    }
                    string cosample = @"select case when sampleqty = '*' then '*' else sampleqty end as Sampleqty from IHPS_QUALITY_SPC_AQLC0 ";
                    cosample += " where ( Lowervalue <=" + lotqty + "and Uppervalue >=" + lotqty + "and AQLValue = " + float.Parse(AQLValue) + ")";
                    DataTable dttqty = DbAccess.SelectBySql(cosample).Tables[0];
                    string Sampleqty = dttqty.Rows[0]["Sampleqty"].ToString();
                    if (Sampleqty.Contains("*"))
                    {
                        txtsampleqty.Text = lotqty.ToString();
                    }
                    else
                    {
                        txtsampleqty.Text = Sampleqty;
                    }
                      textAc.Text = "0";
                      textRe.Text = "1";

                }
                else if (txtsampletype.Text == "全检")
                {
                    string scheckqty = "SELECT count(lotno) qty,SUM(qty) totalqty from delivery where  deliveryid='" + dsinfo.Tables[0].Rows[0]["deliveryid"].ToString() + "' and materialcode='" + dsinfo.Tables[0].Rows[0]["materialcode"].ToString() + "'";
                    DataTable dtqty = Common.DbAccess.SelectBySql(scheckqty).Tables[0];
                    txtsampleqty.Text = dtqty.Rows[0][0].ToString();
                    txttotalqty.Text = dtqty.Rows[0][1].ToString();
                }

                else
                {
                    txtsampleqty.Text = dt.Rows[0]["Samplevalue"].ToString();
                }

                int testvalueqty = int.Parse(dt.Rows[0]["testvalueqty"].ToString());

                DataTable dtscop = bindTypeSet("", "");
                DataRow[] rw = dtscop.Select("TestType='" + txttestitem.Text + "'");
                if (rw.Length > 0 && int.Parse(txtsampleqty.Text) > int.Parse(txtsamplefactqty.Text == "" ? "0" : txtsamplefactqty.Text))
                {
                  TestValueQtyList TVQ = new TestValueQtyList(testvalueqty, txtscope1.Text, txtscope2.Text, txtsampleqty.Text, txtsamplefactqty.Text, txtlotno.Text, int.Parse(txtlotqty.Text), testtype, txttestitem.Text, sproductcode, dt, txtrsno.Text == null ? "" : txtrsno.Text);
                    TVQ.ShowDialog();

                    txtsamplefactqty.Text = TVQ.sfq;
                    TestListbind(testtype, txttestitem.Text, txttestsubitem.Text, txtPacktype.Text, txtlotno.Text, TVQ.sfq);
                    return;
                }

                TestListbind(testtype, txttestitem.Text, txttestsubitem.Text, txtPacktype.Text, txtlotno.Text, "");

                string sFlag = IfNoCheck(testtype, txttestitem.Text, txttestsubitem.Text, sproductcode);
                if (sFlag.IndexOf("无需重测") >= 0)
                {
                    int j = sFlag.IndexOf("无需重测");
                    int i = sFlag.IndexOf(",");

                    txttestvalue.Text = sFlag.Substring(j + 4, i - (j + 4));
                    txtremarks.Text = "上次测试时间为:" + sFlag.Substring(i + 1);
                    txtremarks.BackColor = Color.Yellow;
                    txttestvalue.Enabled = false;
                    txtremarks.Enabled = false;
                }
                else if (sFlag.IndexOf("需重测") >= 0)
                {
                    MessageBox.Show(sFlag, "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    txttestvalue.Text = "";
                    txtremarks.Text = "";
                    txttestvalue.Enabled = true;
                    txtremarks.Enabled = true;
                    txtlotno.Text = "";
                    txtlotno.Focus();
                    txtlotno.SelectAll();
                    return;
                }
            }
            else
            {
                txttestsubitem.Focus();
                bindSubitemInfo();
            }
        }
        private void txttestitem_SelectedValueChanged(object sender, EventArgs e)
        {

            txttestitem_SelectedIndexChanged(sender, e);

        }
        private void txttestsubitem_SelectedIndexChanged(object sender, EventArgs e)
        {
            bindSubitemInfo();
        }

        private void edtStart_Click(object sender, EventArgs e)
        {
            if (sBoxEnable)
            {
                edtStart.Text = "开始";
                ssp.Close();
                cbOK.Enabled = true;
                cbNG.Enabled = true;
                lblunit.Text = "";
            }
            else
            {
                cbOK.Enabled = false;
                cbNG.Enabled = false;

                edtStart.Text = "停止";
                ssp.PortName = edtCom.Text;
                ssp.ReadTimeout = 3000;
                ssp.BaudRate = 19200;
                ssp.StopBits = StopBits.One;
                ssp.DataBits = 8;
                ssp.Parity = Parity.None;

                ssp.ReceivedBytesThreshold = 30;
                try
                {
                    if (ssp.IsOpen)
                    {

                    }
                    else
                    {
                        ssp.Open();
                        ssp.ReadExisting();
                    }
                }
                catch
                {
                    MessageBox.Show("无法打开串口！");
                }
            }
            sBoxEnable = !sBoxEnable;
        }
        private int dev = 0;
        public bool OpenInstrument(int addr)
        {
            //Open and intialize an GPIB instrument  

            dev = GPIB.ibdev(0, addr, 0, (int)GPIB.gpib_timeout.T1s, 1, 0);
            GPIB.ibclr(dev);
            return true;
        }
        public bool write(int addr, string strWrite)
        {
            //Open and intialize an GPIB instrument  

            //Write a string command to a GPIB instrument using the ibwrt() command  
            GPIB.ibwrt(dev, strWrite, strWrite.Length);

            return true;
        }
        public bool read(int addr, ref string strRead)
        {
            //int dev = GPIB.ibdev(0, addr, 0, (int)GPIB.gpib_timeout.T1s, 1, 0);

            StringBuilder str = new StringBuilder(100);

            GPIB.ibrd(dev, str, 100);
            strRead = str.ToString();
            return true;
        }
        public bool CloseInstrument(int addr)
        {
            //Offline the GPIB interface card  
            GPIB.ibonl(dev, 0);
            return true;
        }
        private void IfCorrect(string tval)
        {

            lblTestValueShow.Text = tval;
            //下限再往上浮动%(-)
            float lowlittle = float.Parse(txtscope1.Text == "" ? "0" : txtscope1.Text) * (1 - (float.Parse(txtupscope.Text == "" ? "0" : txtupscope.Text) / 100));
            //上限再往上浮动%(+)
            float upbig = float.Parse(txtscope2.Text == "" ? "0" : txtscope2.Text) * (1 + (float.Parse(txtupscope.Text == "" ? "0" : txtupscope.Text) / 100));
            //测试值在范围之外,不处理
            if (float.Parse(tval) < lowlittle)
            {
                lblTestValueShow.ForeColor = Color.Red;
                return;
            }
            if (float.Parse(tval) > upbig)
            {
                lblTestValueShow.ForeColor = Color.Red;
                return;
            }
            //在测试值范围内,则进行记录测试值
            if ((float.Parse(tval) >= float.Parse(txtscope1.Text == "" ? "0" : txtscope1.Text)) && (float.Parse(tval) <= float.Parse(txtscope2.Text == "" ? "0" : txtscope2.Text)))
            {
                lblTestValueShow.ForeColor = Color.Green;

                txttestvalue.Text = tval.ToString();

                cbOK.Checked = true;
                cbOK.Enabled = false;
                cbNG.Enabled = false;

                //在测试值范围之内,系统弹出一对话框,停留5秒后再写入系统,以免连续写入多笔数据
                StartKiller();
                if (MessageBox.Show("测试值:" + txttestvalue.Text + ",检验结果:" + cbOK.Text, "MessageBox", MessageBoxButtons.OK, MessageBoxIcon.Information) == DialogResult.Cancel)
                {
                    return;
                }
                //如果项目未测试完(即"确定"可按)
                if (btnOK.Enabled)
                    btnOK_Click(null, null);
            }
            else if ((float.Parse(tval) >= lowlittle && float.Parse(tval) <= float.Parse(txtscope1.Text == "" ? "0" : txtscope1.Text))
                           || (float.Parse(tval) >= float.Parse(txtscope2.Text == "" ? "0" : txtscope2.Text) && float.Parse(tval) <= upbig))
            {
                lblTestValueShow.ForeColor = Color.YellowGreen;

                txttestvalue.Text = tval.ToString();

                cbOK.Checked = false;
                cbOK.Enabled = false;
                cbNG.Checked = true;
                cbNG.Enabled = false;

                //在测试值范围之内,系统弹出一对话框,停留5秒后再写入系统,以免连续写入多笔数据
                StartKiller();
                if (MessageBox.Show("测试值:" + txttestvalue.Text + ",检验结果:" + cbOK.Text, "MessageBox", MessageBoxButtons.OK, MessageBoxIcon.Warning) == DialogResult.Cancel)
                {
                    return;
                }
                //如果项目未测试完(即"确定"可按)
                if (btnOK.Enabled)
                    btnOK_Click(null, null);
            }
            else
            {
                lblTestValueShow.ForeColor = Color.Red;
                txttestvalue.Text = "0";
                MessageBox.Show("测试不成功,请检查测量值范围或校准后再测!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                txttestvalue.Focus();
                cbOK.Checked = false;
                cbOK.Enabled = false;
                cbNG.Checked = false;
                cbNG.Enabled = false;
            }
        }
        private void btnElect_Click(object sender, EventArgs e)
        {
            try
            {
                int addres = 5;
                OpenInstrument(addres);
                write(addres, "CALC1:MARK1:Y?");
                string res = "";
                read(addres, ref res);
                int ends = res.IndexOf(',');
                int hposition = 0;
                if (res.Substring(0, ends).IndexOf('-') >= 0)
                    hposition = res.Substring(0, ends).IndexOf('-');
                else if (res.Substring(1, ends).IndexOf('+') >= 0)
                    hposition = res.Substring(1, ends).IndexOf('+');

                int Eposition = res.IndexOf('E');

                string num = res.Substring(hposition + 1, ends - hposition - 1);

                string s = int.Parse(num).ToString();
                int l = s.Length + 1;
                if (lblprodinfo.Text.ToUpper().Contains("NH") && lblprodinfo.Text.ToUpper().Contains("MH"))
                {
                    lblunit.Text = "nH";
                    txttestvalue.Text = (Math.Round(double.Parse(res.Substring(0, hposition - 1)) * 1000000000 * (Math.Pow(10, -(int.Parse(num)))), 3)).ToString();
                }
                else if (lblprodinfo.Text.ToUpper().Contains("MH") && (!lblprodinfo.Text.ToUpper().Contains("OHM") && !lblprodinfo.Text.ToUpper().Contains("BLM") && !lblprodinfo.Text.ToUpper().Contains("磁珠")))
                {
                    lblunit.Text = "mH";
                    txttestvalue.Text = (Math.Round(double.Parse(res.Substring(0, hposition - 1)) * 1000 * (Math.Pow(10, -(int.Parse(num)))), 3)).ToString();
                }
                else if (lblprodinfo.Text.ToUpper().Contains("UH"))
                {
                    lblunit.Text = "uH";
                    txttestvalue.Text = (Math.Round(double.Parse(res.Substring(0, hposition - 1)) * 1000000 * (Math.Pow(10, -(int.Parse(num)))), 3)).ToString();
                }
                else if (lblprodinfo.Text.ToUpper().Contains("NH"))
                {
                    lblunit.Text = "nH";
                    txttestvalue.Text = (Math.Round(double.Parse(res.Substring(0, hposition - 1)) * 1000000000 * (Math.Pow(10, -(int.Parse(num)))), 3)).ToString();
                }
                else if (lblprodinfo.Text.ToUpper().Contains("OHM") || lblprodinfo.Text.ToUpper().Contains("BLM") || lblprodinfo.Text.ToUpper().Contains("磁珠"))
                {
                    lblunit.Text = "Ω";
                    txttestvalue.Text = (Math.Round(double.Parse(res.Substring(0, hposition - 1)) * (Math.Pow(10, (int.Parse(num)))), 3)).ToString();
                }
                else if (lblprodinfo.Text.ToUpper().Contains("电感"))
                {
                    lblunit.Text = "nH";
                    txttestvalue.Text = (Math.Round(double.Parse(res.Substring(0, hposition - 1)) * 1000000000 * (Math.Pow(10, -(int.Parse(num)))), 3)).ToString();
                }
                else
                {
                    lblunit.Text = "H";
                    txttestvalue.Text = (Math.Round(double.Parse(res.Substring(0, hposition - 1)) * (Math.Pow(10, -(int.Parse(num)))), 3)).ToString();
                }
                if ("电阻类,电感类,电容类".Contains(testtype) && txttestsubitem.Text.Contains("测试"))
                {
                    if (lblunit.Text.ToUpper() != lblsetunit.Text.ToUpper())
                    {
                        MessageBox.Show("测试单位与配置的单位不相同", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        return;
                    }
                }

                IfCorrect(txttestvalue.Text);
            }
            catch (Exception ex)
            {
                {
                    cbOK.Enabled = false;
                    cbNG.Enabled = false;
                    ssp.ReadTimeout = 5000;
                    ssp.BaudRate = 9600;
                    ssp.StopBits = StopBits.One;
                    ssp.DataBits = 8;
                    ssp.Parity = Parity.None;

                    ssp.ReceivedBytesThreshold = 30;
                    try
                    {
                        if (ssp.IsOpen)
                        {

                        }
                        else
                        {
                            ssp.Open();
                        }
                        ssp.WriteLine("fetch?");
                        System.Threading.Thread.Sleep(1000);
                        string s = ssp.ReadExisting();
                        ssp.WriteLine("PARA?");
                        System.Threading.Thread.Sleep(1000);
                        string u = ssp.ReadExisting();
                        if (s.ToUpper().Substring(0, s.IndexOf(",")).Contains("E"))
                        {
                            double n = 0, m = 0;
                            int k = 0;
                            s = s.ToUpper().Substring(0, s.IndexOf(","));
                            n = float.Parse(s.Substring(0, s.IndexOf("E")));
                            if (s.IndexOf("+") >= 0)
                                k = int.Parse(s.Substring(s.IndexOf("+") + 1));
                            else if (s.IndexOf("-") >= 0)
                                k = -int.Parse(s.Substring(s.IndexOf("-") + 1));
                            if (u.StartsWith("R") || u.StartsWith("Z"))
                            {
                                m = n * Math.Pow(10, k);
                                s = (n * Math.Pow(10, k)).ToString();
                            }
                            else if (u.StartsWith("C"))
                            {
                                m = n * Math.Pow(10, 9) * Math.Pow(10, k);
                                s = (n * Math.Pow(10, 9) * Math.Pow(10, k)).ToString();
                            }
                            else if (u.StartsWith("L"))
                            {
                                m = n * Math.Pow(10, k);
                                s = (n * Math.Pow(10, k)).ToString();
                            }
                            if (u.StartsWith("R") || u.StartsWith("Z"))
                            {
                                if (m < 1000)
                                {
                                    txttestvalue.Text = s.Substring(0, s.IndexOf(".") + 4);
                                    this.lblunit.Text = "Ω";
                                }
                                else if (m > 1000 && m < 1000000)
                                {
                                    txttestvalue.Text = (m / 1000).ToString().Substring(0, (m / 1000).ToString().IndexOf(".") + 4);
                                    this.lblunit.Text = "KΩ";
                                }
                                else
                                {
                                    txttestvalue.Text = (m / 1000000).ToString().Substring(0, (m / 1000000).ToString().IndexOf(".") + 4);
                                    this.lblunit.Text = "MΩ";
                                }
                            }
                            else if (u.StartsWith("C"))
                            {
                                if (m < 1000)
                                {
                                    txttestvalue.Text = s.Substring(0, s.IndexOf(".") + 4);
                                    this.lblunit.Text = "nF";
                                }
                                else if (m > 1000 && m < 1000000)
                                {
                                    txttestvalue.Text = (m / 1000).ToString().Substring(0, (m / 1000).ToString().IndexOf(".") + 4);
                                    this.lblunit.Text = "uF";
                                }
                                else
                                {
                                    txttestvalue.Text = (m / 1000000).ToString().Substring(0, (m / 1000000).ToString().IndexOf(".") + 4);
                                    this.lblunit.Text = "F";
                                }
                            }
                            else if (u.StartsWith("L"))
                            {
                                if (m < 1000)
                                {
                                    txttestvalue.Text = s.Substring(0, s.IndexOf(".") + 4);
                                    this.lblunit.Text = "uH";
                                }
                                else if (m > 1000 && m < 1000000)
                                {
                                    txttestvalue.Text = (m / 1000).ToString();
                                    this.lblunit.Text = "mH";
                                }
                                else
                                {
                                    txttestvalue.Text = (m / 1000000).ToString();
                                    this.lblunit.Text = "H";
                                }
                            }
                        }
                        if ("电阻类,电感类,电容类".Contains(testtype) && txttestsubitem.Text.Contains("测试"))
                        {
                            if (lblunit.Text.ToUpper() != lblsetunit.Text.ToUpper())
                            {
                                MessageBox.Show("测试单位与配置的单位不相同", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                                return;
                            }
                        }
                        IfCorrect(txttestvalue.Text);
                        ssp.Close();
                    }
                    catch
                    {
                        MessageBox.Show("无法打开串口！");
                    }
                }
            }
        }
        private void bindMdAndExDate(string DC)
        {
            string sql1 = "select dbo.Week2DayFun('" + DC + "') Mdate";
            DataTable dt1 = Common.DbAccess.SelectBySql(sql1).Tables[0];
            if (dt1.Rows.Count <= 0 || dt1.Rows[0][0].ToString() == "格式错误")
            {
                lblinfo.Text = "D/C格式错误";
                lblinfo.ForeColor = Color.Red;
                txtDateCode.Text = "";
                return;
            }
            string sql = "select dbo.Week2DayFun('" + DC + "') Mdate,convert(varchar(10),dateadd(day,isnull(usefullife,90),dbo.Week2DayFun('" + DC + "')),121) ExpiryDate from MaterialSpec where materialcode='" + sproductcode + "'";
            DataTable dt = Common.DbAccess.SelectBySql(sql).Tables[0];
            txtMdate.Text = dt.Rows[0][0].ToString();
            txtExpiryDate.Text = dt.Rows[0][1].ToString();
        }
        private void txtDateCode_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13 && txtDateCode.Text.Trim() != "")
            {
                if (txtDateCode.Text.Trim() == "NA")
                {

                }
                else
                {
                    this.bindMdAndExDate(txtDateCode.Text);
                }
            }
        }

        private void txtDateCode_Leave(object sender, EventArgs e)
        {
            if (txtDateCode.Text.Trim() == "") return;
            if (txtDateCode.Text.Trim() == "NA")
            {

            }
            else
            {
                txtDateCode.Leave -= txtDateCode_Leave;
                bindMdAndExDate(txtDateCode.Text);
                txtDateCode.Leave += txtDateCode_Leave;
            }
        }
        private void cbOK2_CheckedChanged(object sender, EventArgs e)
        {
            if (cbOK2.Checked)
            {
                cbNG2.Checked = false;
                txtstockqty.Enabled = true;
                if (lblprodinfo.Text.ToUpper().Contains("D1专用") || lblprodinfo.Text.ToUpper().Contains("D1客供"))
                {
                    if (txtmadecode.Text == "")
                    {
                        txtmadecode.Focus();
                        txtmadecode.SelectAll();
                        return;
                    }
                }
                txtstockqty.SelectAll();
                txtstockqty.Focus();
            }
        }
        private void cbNG2_CheckedChanged(object sender, EventArgs e)
        {
            if (cbNG2.Checked)
            {
                cbOK2.Checked = false;
                txtstockqty.Enabled = true;
                txtremarks.Enabled = true;
                if (lblprodinfo.Text.ToUpper().Contains("D1专用") || lblprodinfo.Text.ToUpper().Contains("D1客供"))
                {
                    if (txtmadecode.Text == "")
                    {
                        txtmadecode.Focus();
                        txtmadecode.SelectAll();
                        return;
                    }

                }
                txtremarks.SelectAll();
                txtremarks.Focus();
            }
        }
        private void btninstock_Click(object sender, EventArgs e)
        {

        }
        int sampleqty,sreturnqty = 0;

        private void txtreturnqty_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13 && txtreturnqty.Text.Trim() != "")
            {

                txtreturnqty_Leave(sender,e);

                //lblreturnqtyshow.Text = "";
                //if (!int.TryParse(txtreturnqty.Text, out sreturnqty))
                //{
                //    lblreturnqtyshow.Text = "请输入正确的数字";
                //    lblreturnqtyshow.ForeColor = Color.Red;
                //    return;
                //}
                //if (sreturnqty <= 1)
                //{
                //    lblreturnqtyshow.Text = "请输入正确的数字";
                //    lblreturnqtyshow.ForeColor = Color.Red;
                //    return;
                //}
                //string ssampleqty = @"Select case when Sampleqty>lotfactqty then lotfactqty else Sampleqty end as Sampleqty,s.Code from IQC_TestSTD105ECode c ";
                //ssampleqty += "  inner join ";
                //ssampleqty += " (Select " + sreturnqty + " lotfactqty,AQLValue,AQL,CheckLevel,Ac,Re,s.Code from IQC_TestSTD105ECheckSet i inner join IHPS_QUALITY_SPC_AQLIS02859 s on i.Code=s.Code  where LotSizemin<=" + sreturnqty + " and LotSizemax>=" + sreturnqty + " and CheckLevel='II') s on c.Code=s.Code";
                //DataSet dssampleqty = DbAccess.SelectBySql(ssampleqty);
                //txtsampleqty.Text = dssampleqty.Tables[0].Rows[0]["Sampleqty"].ToString();
                //sampleqty = int.Parse(txtsampleqty.Text);

                //string txtCode = " Select a.Code from IQC_TestSTD105ECheckSet i inner join IHPS_QUALITY_SPC_AQLIS02859 a on i.Code=a.Code  where ( LotSizemin<=" + sreturnqty + " and LotSizemax>=" + sreturnqty + " and CheckLevel='II')";
                //DataTable dds = DbAccess.SelectBySql(txtCode).Tables[0];
                //string Code = dds.Rows[0]["Code"].ToString();

                //string sql = @"select Ac,Re from IHPS_QUALITY_SPC_AQLIS02859 where Type ='正常检验' and AQLValue = " + float.Parse(AQLValue) + " and Code = '" + Code + "'";

                //DataTable djt = DbAccess.SelectBySql(sql).Tables[0];

                //textAc.Text = djt.Rows[0]["Ac"].ToString();
                //textRe.Text = djt.Rows[0]["Re"].ToString();
            }
        }

        private void gridView_CustomDrawCell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
        {
            //DataGridViewRow dgr = databind.Rows[e.RowIndex];
            //try
            //{
            //    if (dgr.Cells["TestResult"].Value.ToString() == "NG")
            //    {
            //        dgr.DefaultCellStyle.BackColor = Color.Red;
            //    }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}
            try
            {
                if (e.Column.FieldName == "TestResult")
                {
                    GridCellInfo GridCell = e.Cell as GridCellInfo;
                    if (GridCell.CellValue.ToString() == "NG")
                    {
                        e.Appearance.BackColor = Color.Red;
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
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
            if (sproductcode == "")
                return;
            string floerPath = "\\\\" + this.serverFilePath + "\\丝印图" + "\\" + sproductcode;
            if (Connect(serverFilePath))
            {
                if (Directory.Exists(floerPath))
                {
                    string[] pt = Directory.GetFiles(floerPath);
                    string filename = Path.GetFileName(pt[0].ToString());
                    OpenFile(floerPath, filename);
                }
                else
                {
                    string pold = "";
                    string sqlold = "select pold from UAT_ITEMNEW where Pnew='" + sproductcode + "'";
                    DataTable dtold = DbAccess.SelectBySql(sqlold).Tables[0];
                    if (dtold.Rows.Count > 0)
                    {
                        pold = dtold.Rows[0][0].ToString();
                        string floerPathold = "\\\\" + this.serverFilePath + "\\丝印图" + "\\" + pold;
                        if (Directory.Exists(floerPathold))
                        {
                            string[] pt = Directory.GetFiles(floerPathold);
                            string filename = Path.GetFileName(pt[0].ToString());
                            OpenFile(floerPathold, filename);
                        }
                    }
                }
            }

        }

        private void gridView_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
        {
            DataTable dt = databind.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 0)
            {
                return;
            }
            if (gridView.GetDataRow(e.RowHandle)["TestResult"].ToString() == "NG")
            {
                e.Appearance.BackColor = Color.Red;
            }
        }

        private void txtreturnqty_Leave(object sender, EventArgs e)
        {
            //lblreturnqtyshow.Text = "";
            //if (!int.TryParse(txtreturnqty.Text, out sreturnqty))
            //{
            //    lblreturnqtyshow.Text = "请输入正确的数字";
            //    lblreturnqtyshow.ForeColor = Color.Red;
            //    return;
            //}
            //if (sreturnqty <= 1)
            //{
            //    lblreturnqtyshow.Text = "请输入正确的数字";
            //    lblreturnqtyshow.ForeColor = Color.Red;
            //    return;
            //}
            //string ssampleqty = @"Select case when Sampleqty>lotfactqty then lotfactqty else Sampleqty end as Sampleqty,s.Code from IQC_TestSTD105ECode c ";
            //ssampleqty += "  inner join ";
            //ssampleqty += " (Select " + sreturnqty + " lotfactqty,AQLValue,AQL,CheckLevel,Ac,Re,s.Code from IQC_TestSTD105ECheckSet i inner join IHPS_QUALITY_SPC_AQLIS02859 s on i.Code=s.Code  where LotSizemin<=" +sreturnqty + " and LotSizemax>=" +sreturnqty + " and CheckLevel='II') s on c.Code=s.Code";
            //DataSet dssampleqty = DbAccess.SelectBySql(ssampleqty);
            //txtsampleqty.Text = dssampleqty.Tables[0].Rows[0]["Sampleqty"].ToString();
            //sampleqty = int.Parse(txtsampleqty.Text);

            //string txtCode = " Select a.Code from IQC_TestSTD105ECheckSet i inner join IHPS_QUALITY_SPC_AQLIS02859 a on i.Code=a.Code  where ( LotSizemin<=" + sreturnqty + " and LotSizemax>=" + sreturnqty + " and CheckLevel='II')";
            //DataTable dds = DbAccess.SelectBySql(txtCode).Tables[0];
            //string Code = dds.Rows[0]["Code"].ToString();

            //string sql = @"select Ac,Re from IHPS_QUALITY_SPC_AQLIS02859 where Type ='正常检验' and AQLValue = " + float.Parse(AQLValue) + " and Code = '" + Code + "'";

            //DataTable djt = DbAccess.SelectBySql(sql).Tables[0];

            //textAc.Text = djt.Rows[0]["Ac"].ToString();
            //textRe.Text = djt.Rows[0]["Re"].ToString();

            if (txtreturnqty.Text.Trim() == "")
                return;

            lblreturnqtyshow.Text = "";
            if (!int.TryParse(txtreturnqty.Text, out sreturnqty))
            {
                lblreturnqtyshow.Text = "请输入正确的数字";
                lblreturnqtyshow.ForeColor = Color.Red;
                return;
            }
            if (sreturnqty <= 1)
            {
                lblreturnqtyshow.Text = "请输入正确的数字";
                lblreturnqtyshow.ForeColor = Color.Red;
                return;
            }
            string ssampleqty = @"Select case when Sampleqty>lotfactqty then lotfactqty else Sampleqty end as Sampleqty,s.Code from IQC_TestSTD105ECode c ";
            ssampleqty += "  inner join ";
            ssampleqty += " (Select " + sreturnqty + " lotfactqty,AQLValue,AQL,CheckLevel,Ac,Re,s.Code from IQC_TestSTD105ECheckSet i inner join IHPS_QUALITY_SPC_AQLIS02859 s on i.Code=s.Code  where LotSizemin<=" + sreturnqty + " and LotSizemax>=" + sreturnqty + " and CheckLevel='II') s on c.Code=s.Code";
            DataSet dssampleqty = DbAccess.SelectBySql(ssampleqty);
            txtsampleqty.Text = dssampleqty.Tables[0].Rows[0]["Sampleqty"].ToString();
            sampleqty = int.Parse(txtsampleqty.Text);

            string txtCode = " Select a.Code from IQC_TestSTD105ECheckSet i inner join IHPS_QUALITY_SPC_AQLIS02859 a on i.Code=a.Code  where ( LotSizemin<=" + sreturnqty + " and LotSizemax>=" + sreturnqty + " and CheckLevel='II')";
            DataTable dds = DbAccess.SelectBySql(txtCode).Tables[0];
            string Code = dds.Rows[0]["Code"].ToString();

            string sql = @"select Ac,Re from IHPS_QUALITY_SPC_AQLIS02859 where Type ='正常检验' and AQLValue = " + float.Parse(AQLValue) + " and Code = '" + Code + "'";

            DataTable djt = DbAccess.SelectBySql(sql).Tables[0];

            textAc.Text = djt.Rows[0]["Ac"].ToString();
            textRe.Text = djt.Rows[0]["Re"].ToString();

        }

        private void cbOK_CheckedChanged(object sender, EventArgs e)
        {
            if (cbOK.Checked)
            {
                cbNG.Checked = false;
               
                Badcategory.Visible = false;
                Badcategory.Enabled = false;

            }
            else
            {              
                Badcategory.Visible = false;
                Badcategory.Enabled = false;
            }
        }

        private void cbNG_CheckedChanged(object sender, EventArgs e)
        {
            if (cbNG.Checked)
            {
                cbOK.Checked = false;
                Badcategory.Visible = true;
                Badcategory.Enabled = true;
                Badcategory.Properties.DataSource = BadSituation();
                Badcategory.Properties.DisplayMember = "不良类别";
                Badcategory.Properties.ValueMember = "不良类别";
            }
            else
            {
                Badcategory.Enabled = false;
                Badcategory.Visible = false;
            }
        }
        private bool IfInputCorrect(string tval)
        {
            bool f = false;
            lblTestValueShow.Text = tval;
            //下限再往上浮动%(-)
            float lowlittle = float.Parse(txtscope1.Text == "" ? "0" : txtscope1.Text) * (1 - (float.Parse(txtupscope.Text == "" ? "0" : txtupscope.Text) / 100));
            //上限再往上浮动%(+)
            float upbig = float.Parse(txtscope2.Text == "" ? "0" : txtscope2.Text) * (1 + (float.Parse(txtupscope.Text == "" ? "0" : txtupscope.Text) / 100));
            //测试值在范围之外,不处理
            if (float.Parse(tval) < lowlittle)
            {
                lblTestValueShow.ForeColor = Color.Red;
            }
            if (float.Parse(tval) > upbig)
            {
                lblTestValueShow.ForeColor = Color.Red;
            }


            //在测试值范围内,则进行记录测试值
            if ((float.Parse(tval) >= float.Parse(txtscope1.Text == "" ? "0" : txtscope1.Text)) && (float.Parse(tval) <= float.Parse(txtscope2.Text == "" ? "0" : txtscope2.Text)))
            {
                lblTestValueShow.ForeColor = Color.Green;

                txttestvalue.Text = tval.ToString();

                cbOK.Checked = true;
                cbOK.Enabled = false;
                cbNG.Enabled = false;
                f = true;
            }


            else if ((float.Parse(tval) >= lowlittle && float.Parse(tval) <= float.Parse(txtscope1.Text == "" ? "0" : txtscope1.Text))
                          || (float.Parse(tval) >= float.Parse(txtscope2.Text == "" ? "0" : txtscope2.Text) && float.Parse(tval) <= upbig))
            {
                lblTestValueShow.ForeColor = Color.YellowGreen;

                txttestvalue.Text = tval.ToString();

                cbOK.Checked = false;
                cbOK.Enabled = false;
                cbNG.Checked = true;
                cbNG.Enabled = false;
                f = true;
            }
            else
            {
                lblTestValueShow.ForeColor = Color.Red;
                txttestvalue.Text = "0";
                MessageBox.Show("测试不成功,请检查测量值范围或校准后再测!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                txttestvalue.Focus();
                cbOK.Checked = false;
                cbOK.Enabled = false;
                cbNG.Checked = false;
                cbNG.Enabled = false;
                f = false;
            }
            return f;
        }
        private void btnOK_Click(object sender, EventArgs e)
        {

            //20150323增加:电阻类、电感类、电容类,如果配置的单位与实测的单位不一致,则不让测试
            if ("电阻类,电感类,电容类".Contains(testtype) && txttestsubitem.Text.Contains("测试"))
            {
                if (lblunit.Text.ToUpper() != lblsetunit.Text.ToUpper())
                {
                    MessageBox.Show("测试单位与配置的单位不相同", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }
                bool flag = IfInputCorrect(txttestvalue.Text);
                if (!flag)
                    return;
            }
            //20150323增加:电阻类、电感类、电容类,如果配置的单位与实测的单位不一致,则不让测试

            //if (txtreturnqty.Text.Trim() =="" && txtsampletype.Text == "ISO2859-1")
            //{
            //    MessageBox.Show("请输入退料数量", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //    return; 
            //}

            if (txtreturnqty.Text.Trim() == "")
            {
                MessageBox.Show("请输入退料数量", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (returnreason.Text.Trim()=="")
            {
                MessageBox.Show("请选择退料原因", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (txttestitem.Text.Contains("包装确认")&&txtDateCode.Text == "")
            {
                lblinfo.Text = "请输入D/C";
                lblinfo.ForeColor = Color.Red;
                return;
            }

            //if (lblprodinfo.Text.ToUpper().Contains("F1专用") || lblprodinfo.Text.ToUpper().Contains("F1客供"))
            //{
            //    string sqlf1 = "   select convert(varchar(10),ExpiryDate,121) ExpiryDate from materialRelation  where reelid = '" + txtlotno.Text.Trim() + "'  ";
            //    DataTable dtf1 = DbAccess.SelectBySql(sqlf1).Tables[0];
            //    if (dtf1 != null && dtf1.Rows.Count > 0)
            //    {
            //        string expirydatef1 = dtf1.Rows[0]["ExpiryDate"].ToString();
            //        if (!string.IsNullOrEmpty(expirydatef1) && (DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")) > DateTime.Parse(expirydatef1)))
            //        {
            //            IQCTestExpiryDate F1extendExpiryDate = new IQCTestExpiryDate(txtlotno.Text.Trim());
            //            F1extendExpiryDate.ShowDialog();
            //        }
            //    }
            //}


            //if (lblprodinfo.Text.ToUpper().Contains("G2专用") || lblprodinfo.Text.ToUpper().Contains("G2客供"))
            //{
            //    string sqlf1 = "   select convert(varchar(10),ExpiryDate,121) ExpiryDate from materialRelation  where reelid = '" + txtlotno.Text.Trim() + "'  ";
            //    DataTable dtf1 = DbAccess.SelectBySql(sqlf1).Tables[0];
            //    if (dtf1 != null && dtf1.Rows.Count > 0)
            //    {
            //        string expirydatef1 = dtf1.Rows[0]["ExpiryDate"].ToString();
            //        if (!string.IsNullOrEmpty(expirydatef1) && (DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")) > DateTime.Parse(expirydatef1)))
            //        {
            //            IQCTestExpiryDate F1extendExpiryDate = new IQCTestExpiryDate(txtlotno.Text.Trim());
            //            F1extendExpiryDate.ShowDialog();
            //        }
            //    }
            //}


            string badcategory = Badcategory.Text;
            string returnrea = returnreason.Text;

            string sstate = "NG";
            if (cbOK.Checked == false && cbNG.Checked == false)
            {
                lblinfo.Text = "请选择一个结果";
                lblinfo.ForeColor = Color.Red;
                return;
            }
            if (cbOK.Checked)
            {
                sstate = "OK";
                badcategory = "";
                returnrea = "";
            }
            else if (cbNG.Checked)
            {
                sstate = "NG";
                if (cbNG.Enabled == true && (returnreason.Text == "" || Badcategory.Text == ""))
                {
                    MessageBox.Show("请确认退料原因和不良类别！", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
            }
            //if (txtsampleqty.Text == txtsamplefactqty.Text) return;
            if (txtlotno.Text == "")
                return;

            string msg = "";
            try
            {
                //msg = ic.AddNewTestList("退料测试", testtype, txttestitem.SelectedValue.ToString(), txttestsubitem.SelectedValue.ToString(), txttestdes.Text, txttesttools.Text, txtPacktype.Text, txtsampletype.Text, float.Parse(txtscope2.Text == "" ? "0" : txtscope2.Text),
                //     float.Parse(txtscope1.Text == "" ? "0" : txtscope1.Text), float.Parse(AQLValue == "" ? "0" : AQLValue), txtlotno.Text, int.Parse(txtlotqty.Text), Login.username, int.Parse(txtsampleqty.Text), 1, sproductcode, txtAQL.Text, sstate, "", txtremarks.Text, 0, txttestvalue.Text,
                //     lblunit.Text == "" ? lblsetunit.Text : lblunit.Text, sCheckType, txtDateCode.Text, txtMdate.Text, txtExpiryDate.Text, txtrsno.SelectedValue == null ? "" : txtrsno.SelectedValue.ToString());
                msg = ic.AddNewTestListNew("退料测试", testtype, txttestitem.Text, txttestsubitem.Text, txttestdes.Text, txttesttools.Text, txtPacktype.Text, txtsampletype.Text, float.Parse(txtscope2.Text == "" ? "0" : txtscope2.Text),
                     float.Parse(txtscope1.Text == "" ? "0" : txtscope1.Text), float.Parse(AQLValue == "" ? "0" : AQLValue), txtlotno.Text.Trim(), int.Parse(txtlotqty.Text), Login.username, int.Parse(txtsampleqty.Text), 1, sproductcode, txtAQL.Text, sstate, "", txtremarks.Text, 0, txttestvalue.Text,
                     lblunit.Text == "" ? lblsetunit.Text : lblunit.Text, sCheckType, txtDateCode.Text, txtMdate.Text, txtExpiryDate.Text, txtrsno.Text == null ? "" : txtrsno.Text, badcategory, "","",sreturnqty == 0 ? 0 : sreturnqty, returnrea);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            if (msg.IndexOf("OK") >= 0)
            {
                lblinfo.Text = msg;
                lblinfo.ForeColor = Color.Blue;
                TestListbind(testtype, txttestitem.Text, txttestsubitem.Text, txtPacktype.Text, txtlotno.Text, "");
                txttestvalue.Text = "";
                txtremarks.Text = "";
                txtDateCode.Text = "";
                cbOK.Checked = false;
                cbNG.Checked = false;                
                sreturnqty = 0;
                this.txtDateCode.Text = "";
                txtMdate.Text = "";
                txtExpiryDate.Text = "";
                GetRSNO();
            }
            else
            {
                MessageBox.Show(msg, "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                DbAccess.SetControlEmpty(this);
                sreturnqty = 0;
                txtlotno.Focus();
            }

        }

        private void btndel_Click(object sender, EventArgs e)
        {
            //if (databind.SelectedRows.Count <= 0) return;
            //for (int i = databind.SelectedRows.Count; i > 0; i--)
            //{
            //    string k = ic.AddNewTestList("退料删除", databind.SelectedRows[i - 1].Cells["TestType"].Value.ToString(), databind.SelectedRows[i - 1].Cells["TestItem"].Value.ToString(), databind.SelectedRows[i - 1].Cells["TestSubItem"].Value.ToString(), "", "", databind.SelectedRows[i - 1].Cells["PackType"].Value.ToString(),
            //                      "", 0, 0, 0, databind.SelectedRows[i - 1].Cells["LotNo"].Value.ToString(), 0, Login.username, 0, 0, "", "", "", "", "", int.Parse(databind.SelectedRows[i - 1].Cells["Items"].Value.ToString()), "", "", "", "", "", "", "");

            //    if (k.IndexOf("OK") >= 0)
            //        this.databind.Rows.RemoveAt(databind.SelectedRows[i - 1].Index);
            //}

            //txtsamplefactqty.Text = databind.Rows.Count.ToString();
            //if (txtsampleqty.Text == txtsamplefactqty.Text)
            //{
            //    txttestvalue.Enabled = false;
            //    txtremarks.Enabled = false;
            //    btnOK.Enabled = false;
            //}
            //else
            //{
            //    txttestvalue.Enabled = true;
            //    txtremarks.Enabled = true;
            //    btnOK.Enabled = true;
            //}

            DataTable dt = databind.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
                return;
            if (gridView.FocusedRowHandle < 0)
                return;
            if (gridView.GetSelectedRows().Length < 0)
                return;

            for (int i = gridView.GetSelectedRows().Length; i > 0; i--)
            {
                string k = ic.AddNewTestList("退料删除", gridView.GetDataRow(gridView.GetSelectedRows()[i - 1])["TestType"].ToString(), gridView.GetDataRow(gridView.GetSelectedRows()[i - 1])["TestItem"].ToString(), gridView.GetDataRow(gridView.GetSelectedRows()[i - 1])["TestSubItem"].ToString(), "", "", gridView.GetDataRow(gridView.GetSelectedRows()[i - 1])["PackType"].ToString(),
                                  "", 0, 0, 0, gridView.GetDataRow(gridView.GetSelectedRows()[i - 1])["LotNo"].ToString(), 0, Login.username, 0, 0, "", "", "", "", "", int.Parse(gridView.GetDataRow(gridView.GetSelectedRows()[i - 1])["Items"].ToString()), "", "", "", "", "", "", "");

                if (k.IndexOf("OK") >= 0)
                gridView.DeleteRow(gridView.GetSelectedRows()[i - 1]);
            }
            DataTable dts = databind.DataSource as DataTable;
            txtsamplefactqty.Text = dts.Rows.Count.ToString();
            if (txtsampleqty.Text == txtsamplefactqty.Text)
            {
                txttestvalue.Enabled = false;
                txtremarks.Enabled = false;
                btnOK.Enabled = false;
            }
            else
            {
                txttestvalue.Enabled = true;
                txtremarks.Enabled = true;
                btnOK.Enabled = true;
            }

        }

        private void btnsearch_Click(object sender, EventArgs e)
        {
            TestListbind(testtype, txttestitem.Text, txttestsubitem.Text, txtPacktype.Text,txtlotno.Text.Trim(), "");
        }

        private void btnfinished_Click(object sender, EventArgs e)
        {
            txtlotno.Text = "";
            txtlotno.Focus();
        }

        private void databind_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            //DataGridViewRow dgr = databind.Rows[e.RowIndex];
            //try
            //{
            //    if (dgr.Cells["TestResult"].Value.ToString() == "NG")
            //    {
            //        dgr.DefaultCellStyle.BackColor = Color.Red;
            //    }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}
        }


    }
}