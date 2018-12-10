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
//using Oracle.DataAccess.Client;
using DevExpress.XtraBars;
using DX_QMS.Common;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using System.Data.OleDb;
using DevExpress.XtraEditors.Controls;

namespace DX_QMS
{
    public partial class IQCTestListcs : DevExpress.XtraBars.Ribbon.RibbonForm
    {
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
        string StrSql = "";
         IQC ic = new IQC();
        string MaterialState = "";
        string sBadcategory, sBaddescribe, stextmanufacturer;
        public IQCTestListcs()
        {
            InitializeComponent();
            string path = System.Configuration.ConfigurationManager.AppSettings["ServerFilePath"].ToString();
            this.serverFilePath = @path;
            dtalready.Columns.Add("id", typeof(string));
            dtalready.Columns.Add("qty", typeof(int));           
            setRule();
            GetRSNO();
        }
        DataTable BadSituation()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("不良类别", System.Type.GetType("System.String"));
            dt.Columns.Add("不良现象", System.Type.GetType("System.String"));
            string[] bad = { "包装", "标签", "数量", "错料", "外观", "尺寸", "试装不良", "功能", "环保", "其它" };
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

        private Decimal ChangeDataToD(string strData)
        {
            Decimal dData = 0.0M;
            if (strData.Contains("E"))
            {
                dData = Convert.ToDecimal(Decimal.Parse(strData.ToString(), System.Globalization.NumberStyles.Float));
            }
            return dData;
        }
        private void KillMessageBox()
        {
            //按照MessageBox的标题，找到MessageBox的窗口   
            IntPtr ptr = FindWindow(null, "MessageBox");
            if (ptr != IntPtr.Zero)
            {
                //找到则关闭MessageBox窗口   
                PostMessage(ptr, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
            }

        }
        private void Timer_Tick(object sender, EventArgs e)
        {
            KillMessageBox();
            //停止Timer   
            ((Timer)sender).Stop();
        }
        private void StartKiller()
        {
            Timer timer = new Timer();
            timer.Interval = 3000; //3秒启动   
            timer.Tick += new EventHandler(Timer_Tick);
            timer.Start();
        }

        public void sspDataReceived(object sender, System.IO.Ports.SerialDataReceivedEventArgs e)
        {
            byte[] readBuffer = new byte[ssp.BytesToRead];
            ssp.Read(readBuffer, 0, readBuffer.Length);
            this.Invoke(updateSmallBox, new object[] { Encoding.UTF8.GetString(readBuffer) });
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
            Dictionary<string, bool> dic = GroupPermission.QMS_SelectRulesForForm(post, "测试");
            btninstock.Enabled = bool.Parse(dic["hasInsert"].ToString());
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
        private DataTable alreadyinstockqty(string deliveryid, string productcode)
        {
            string ssql = "select ISNULL(SUM(OKqty),0) OKqty,ISNULL(SUM(NGqty),0) NGqty from IQC_OtherInstock where  receptid='" + deliveryid + "' and productcode='" + productcode + "'";
            DataSet ds = Common.DbAccess.SelectBySql(ssql);
            return ds.Tables[0];
        }
        private void GetERPInstockqty(string sDeliveryID, string sMaterialCode)
        {

            string sOraSql;
            int m = 0;
            if (sDeliveryID.Contains("REC"))
            {
                DataTable dt = alreadyinstockqty(sDeliveryID, sMaterialCode);
                //m = int.Parse(txtlotqty.Text) - int.Parse(dt.Rows[0]["OKqty"].ToString()) - int.Parse(dt.Rows[0]["NGqty"].ToString());
                m = int.Parse(txttotalqty.Text == "" ? txtlotqty.Text : txttotalqty.Text) - int.Parse(dt.Rows[0]["OKqty"].ToString()) - int.Parse(dt.Rows[0]["NGqty"].ToString());
                txtstockqty.Text = m.ToString();
            }
            else
            {
                if (dtpub != null && dtpub.Rows.Count > 0)
                {
                    for (int i = 0; i < dtpub.Rows.Count; i++)
                    {
                        m = m + int.Parse(dtpub.Rows[i]["primary_quantity"].ToString());
                    }
                    txtstockqty.Text = m.ToString();
                }
            }
        }
        private string getdeliveryinfo(string receipt, string materialcode)
        {
            int receiveqty = 0, acceptqty = 0;
            string sOraSql;
            //(状态为:RECEIVE,ACCEPT)
            sOraSql = " select transaction_type,ship_to_org_id,INVENTORY_ITEM_ID,transaction_id ,receipt_num, round(sum(quantity),0) primary_quantity, item_number,max(item_desc) item_desc,max(transaction_date) transaction_date,ORG_ID ";
            sOraSql += " from cux_inv_rec4_v where receipt_num='" + receipt + "' and item_number='" + materialcode + "'";
            sOraSql += " group by ship_to_org_id,transaction_id,INVENTORY_ITEM_ID,receipt_num,item_number,transaction_type,ORG_ID";

            DataSet dsERPDelivery = Common.DbAccess.SelectByOracle(sOraSql);
            if (dsERPDelivery != null && dsERPDelivery.Tables.Count > 0 && dsERPDelivery.Tables[0].Rows.Count > 0)
            {
                DataRow[] drr;
                //if(dsERPDelivery.Tables[0].Rows[0]["ORG_ID"].ToString()!="82")
                if (dsERPDelivery.Tables[0].Rows[0]["ORG_ID"].ToString() == "81")
                    drr = dsERPDelivery.Tables[0].Select("transaction_type='RECEIVE'");
                else if (dsERPDelivery.Tables[0].Rows[0]["ORG_ID"].ToString() == "203")
                    drr = dsERPDelivery.Tables[0].Select("transaction_type='TRANSFER'");
                else
                    drr = dsERPDelivery.Tables[0].Select("transaction_type='TRANSFER'");
                if (drr.Length > 0)
                {
                    for (int i = 0; i < drr.Length; i++)
                    {
                        receiveqty = receiveqty + int.Parse(drr[i]["primary_quantity"].ToString());
                    }
                }
                DataRow[] dra = dsERPDelivery.Tables[0].Select("transaction_type='ACCEPT'");
                if (dra.Length > 0)
                {
                    for (int j = 0; j < dra.Length; j++)
                    {
                        acceptqty = acceptqty + int.Parse(dra[j]["primary_quantity"].ToString());
                    }
                }
            }
            return "剩余量:" + receiveqty.ToString() + ",已检量：" + acceptqty.ToString();
        }
        private void TestListbind(string testtype, string testitem, string testsubitem, string packtype, string lotno, string sampleqty)
        {
            DataSet ds = ic.SelectTestListNew("查询", testtype, testitem, testsubitem, "", "", packtype, "", 0, 0, 0, lotno, 0, Login.username, 0, 1, "", "", "", "", "", 0, "", "", "", "", "", "", "","","","",0,"");

            databind.DataSource = ds.Tables[0];

            //if (testitem == "规格尺寸")
            DataTable dtscop = bindTypeSet("", "");
            DataRow[] rw = dtscop.Select("TestType='" + testitem + "'");

            if (rw.Length > 0)
                txtsamplefactqty.Text = sampleqty;
            else
                txtsamplefactqty.Text = ds.Tables[0].Rows.Count.ToString();
            if (int.Parse(txtsampleqty.Text) <= int.Parse(txtsamplefactqty.Text == "" ? "0" : txtsamplefactqty.Text) && testsubitem.Contains("外观"))
            {
                txtremarks.Enabled = false;
                btnOK.Enabled = false;

                txtlotno.SelectAll();
                txtlotno.Focus();

                //getInfo(lotno);
                GetERPInstockqty(ds.Tables[0].Rows[0]["receptid"].ToString(), ds.Tables[0].Rows[0]["Productcode"].ToString());
                this.getdeliveryinfo(ds.Tables[0].Rows[0]["receptid"].ToString(), ds.Tables[0].Rows[0]["Productcode"].ToString());
            }
            else if (int.Parse(txtsampleqty.Text) <= int.Parse(txtsamplefactqty.Text == "" ? "0" : txtsamplefactqty.Text))
            {
              
                txtremarks.Enabled = false;
                btnOK.Enabled = false;

                txtlotno.SelectAll();
                txtlotno.Focus();
            }
            else
            {
                txtstockqty.Text = "";
                txtremarks.Enabled = true;
                btnOK.Enabled = true;
            }
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
        private void bindSubitemInfo()
        {
            if (txttestitem.Text == null || txttestsubitem.Text == null) return;
            DataTable dt = dttitem.Clone();
            DataRow[] arrDr = dttitem.Select("TestItem='" + txttestitem.Text.Trim() + "' and TestSubItem='" + txttestsubitem.Text.Trim() + "'");
            for (int i = 0; i < arrDr.Length; i++)
            {
                dt.Rows.Add(arrDr[i].ItemArray);
            }
            testtype = dt.Rows[0]["TestType"].ToString();
            lblAQLValue.Text = dt.Rows[0]["AQLValue"].ToString();
            AQLValue = dt.Rows[0]["AQLValue"].ToString();

            lblAQLRe.Text = showAQLRcvalue(AQLValue);

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

            //checkcycle = int.Parse(dt.Rows[0]["checkcycle"].ToString());

            int testvalueqty = int.Parse(dt.Rows[0]["testvalueqty"].ToString());

            //if (txtsampletype.Text == "MIL-STD-105E")
            if (txtsampletype.Text == "ISO2859-1")
            {
                if (MaterialState == "加严检验")
                {
                    if (lotqty > 1)
                    {
                        string ssampleqty = @"Select case when Sampleqty>lotfactqty then lotfactqty else Sampleqty end as Sampleqty,s.Code from IQC_TestSTD105ECode c ";
                        ssampleqty += "  inner join ";
                        ssampleqty += " (Select " + lotqty + " lotfactqty,AQLValue,AQL,CheckLevel,Ac,Re,s.Code from IQC_TestSTD105ECheckSet i inner join IHPS_QUALITY_SPC_AQLIS02859 s on i.Code=s.Code  where LotSizemin<=" + lotqty.ToString() + " and LotSizemax>=" + lotqty.ToString() + " and CheckLevel='II') s on c.Code=s.Code";
                        DataSet dssampleqty = DbAccess.SelectBySql(ssampleqty);
                        txtsampleqty.Text = dssampleqty.Tables[0].Rows[0]["Sampleqty"].ToString();

                        string txtCode = " Select a.Code from IQC_TestSTD105ECheckSet i inner join IHPS_QUALITY_SPC_AQLIS02859 a on i.Code=a.Code  where ( LotSizemin<=" + lotqty + " and LotSizemax>=" + lotqty + " and CheckLevel='II')";
                        DataTable dds = DbAccess.SelectBySql(txtCode).Tables[0];
                        string Code = dds.Rows[0]["Code"].ToString();

                        string sql = @"select Ac,Re from IHPS_QUALITY_SPC_AQLIS02859 where Type ='加严' and AQLValue = " + float.Parse(AQLValue) + " and Code = '" + Code + "'";

                        DataTable djt = DbAccess.SelectBySql(sql).Tables[0];

                        lblAQLAc.Text = "Ac=" + djt.Rows[0]["Ac"].ToString();
                        lblAQLRe.Text = "Re=" + djt.Rows[0]["Re"].ToString();
                    }
                    else
                    {
                        txtsampleqty.Text = "1";
                        lblAQLAc.Text = "Ac= 1";
                        lblAQLRe.Text = "Re= 1";
                    }
                }
                if (MaterialState == "放宽检验")
                {
                    if (lotqty > 1)
                    {
                        string ssampleqty = @"Select case when Sampleqty>lotfactqty then lotfactqty else Sampleqty end as Sampleqty,s.Code from IQC_TestSTD105ECode c ";
                        ssampleqty += "  inner join ";
                        ssampleqty += " (Select " + lotqty + " lotfactqty,AQLValue,AQL,CheckLevel,Ac,Re,s.Code from IQC_TestSTD105ECheckSet i inner join IHPS_QUALITY_SPC_AQLIS02859 s on i.Code1=s.Code  where LotSizemin<=" + lotqty.ToString() + " and LotSizemax>=" + lotqty.ToString() + " and CheckLevel='II') s on c.Code=s.Code";
                        DataSet dssampleqty = DbAccess.SelectBySql(ssampleqty);
                        txtsampleqty.Text = dssampleqty.Tables[0].Rows[0]["Sampleqty"].ToString();


                        string txtCode = " select s.Code from IQC_TestSTD105ECheckSet i inner join IHPS_QUALITY_SPC_AQLIS02859 s on i.Code1=s.Code  where ( LotSizemin<=" + lotqty + " and LotSizemax>=" + lotqty + " and CheckLevel='II' )";
                        DataTable dds = DbAccess.SelectBySql(txtCode).Tables[0];
                        string Code = dds.Rows[0]["Code"].ToString();
                        string Value = AQLValue;
                        string sql = @"select Ac,Re from IHPS_QUALITY_SPC_AQLIS02859 where ( Type ='放宽' and  AQLValue = " + float.Parse(AQLValue) + " and Code = '" + Code + "')";
                        DataTable dft = DbAccess.SelectBySql(sql).Tables[0];

                        lblAQLAc.Text = "Ac=" + dft.Rows[0]["Ac"].ToString();
                        lblAQLRe.Text = "Re=" + dft.Rows[0]["Re"].ToString();
                    }
                    else
                    {
                        txtsampleqty.Text = "1";
                        lblAQLAc.Text = "Ac= 1";
                        lblAQLRe.Text = "Re= 1";
                    }
                }
                if (MaterialState == "正常检验")
                {
                    if (lotqty > 1)
                    {
                        string ssampleqty = @"Select case when Sampleqty>lotfactqty then lotfactqty else Sampleqty end as Sampleqty,s.Code from IQC_TestSTD105ECode c ";
                        ssampleqty += "  inner join ";
                        ssampleqty += " (Select " + lotqty + " lotfactqty,AQLValue,AQL,CheckLevel,Ac,Re,s.Code from IQC_TestSTD105ECheckSet i inner join IHPS_QUALITY_SPC_AQLIS02859 s on i.Code=s.Code  where LotSizemin<=" + lotqty.ToString() + " and LotSizemax>=" + lotqty.ToString() + " and CheckLevel='II') s on c.Code=s.Code";
                        DataSet dssampleqty = DbAccess.SelectBySql(ssampleqty);
                        txtsampleqty.Text = dssampleqty.Tables[0].Rows[0]["Sampleqty"].ToString();

                        string txtCode = " Select a.Code from IQC_TestSTD105ECheckSet i inner join IHPS_QUALITY_SPC_AQLIS02859 a on i.Code=a.Code  where ( LotSizemin<=" + lotqty + " and LotSizemax>=" + lotqty + " and CheckLevel='II')";
                        DataTable dds = DbAccess.SelectBySql(txtCode).Tables[0];
                        string Code = dds.Rows[0]["Code"].ToString();

                        string sql = @"select Ac,Re from IHPS_QUALITY_SPC_AQLIS02859 where Type ='正常检验' and AQLValue = " + float.Parse(AQLValue) + " and Code = '" + Code + "'";

                        DataTable djt = DbAccess.SelectBySql(sql).Tables[0];

                        lblAQLAc.Text = "Ac=" + djt.Rows[0]["Ac"].ToString();
                        lblAQLRe.Text = "Re=" + djt.Rows[0]["Re"].ToString();
                    }
                    else
                    {
                        txtsampleqty.Text = "1";
                        lblAQLAc.Text = "Ac= 1";
                        lblAQLRe.Text = "Re= 1";
                    }
                }

            }
            else if (txtsampletype.Text == "C=0")
            {
                if (lotqty >= 500000)
                {
                    lotqty = 500000;
                }
                string cosample = @"select case when sampleqty = '*' then '*' else sampleqty end as Sampleqty from IHPS_QUALITY_SPC_AQLC0 ";
                cosample += " where ( Lowervalue <=" + lotqty + "and Uppervalue >=" + lotqty + " and AQLValue =" + float.Parse(AQLValue) + ")";
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

                lblAQLAc.Text = "Ac=0";
                lblAQLRe.Text = "Re=1";


            }
            else if (txtsampletype.Text == "全检")
            {
                string scheckqty = "SELECT count(lotno) qty,SUM(qty) totalqty from delivery where deliveryid='" + dsinfo.Tables[0].Rows[0]["deliveryid"].ToString() + "' and materialcode='" + dsinfo.Tables[0].Rows[0]["materialcode"].ToString() + "'";

                DataTable dtqty = DbAccess.SelectBySql(scheckqty).Tables[0];
                txtsampleqty.Text = dtqty.Rows[0][0].ToString();
                ///////txttotalqty.Text = dtqty.Rows[0][1].ToString();
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
            //20140922修改,不用每次都去绑定图片及路径
            //showpic(sproductcode);
            //showsampledirectory(sproductcode);
            //20140922修改,不用每次都去绑定图片及路径



            string sFlag = IfNoCheck(testtype, txttestitem.Text, txttestsubitem.Text, sproductcode);
            if (sFlag.IndexOf("无需重测") >= 0)
            {
                int j = sFlag.IndexOf("无需重测");
                int i = sFlag.IndexOf(",");

              
                txtremarks.Text = "上次测试时间为:" + sFlag.Substring(i + 1);
                txtremarks.BackColor = Color.Yellow;
                txtremarks.Enabled = false;
                //return;
            }
            else if (sFlag.IndexOf("需重测") >= 0)
            {
                MessageBox.Show(sFlag, "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtremarks.Text = "";
                txtremarks.Enabled = true;

                //20151105新增
                txtlotno.Text = "";
                txtlotno.Focus();
                txtlotno.SelectAll();
                return;
                //20151105新增
            }



            //if (txttestitem.SelectedValue.ToString()=="规格尺寸" && int.Parse(txtsampleqty.Text) > int.Parse(txtsamplefactqty.Text==""?"0":txtsamplefactqty.Text))

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
        private void txttestsubitem_SelectedIndexChanged(object sender, EventArgs e)
        {
            bindSubitemInfo();
        }
        private void txttestitem_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable dt = dttitem.Clone();
            DataRow[] arrDr = dttitem.Select("TestItem='" + txttestitem.Text + "'");
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

                lblAQLRe.Text = showAQLRcvalue(AQLValue);

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
                    if (MaterialState == "加严检验")
                    {
                        if (lotqty > 1)
                        {
                            string ssampleqty = @"Select case when Sampleqty>lotfactqty then lotfactqty else Sampleqty end as Sampleqty,s.Code from IQC_TestSTD105ECode c ";
                            ssampleqty += "  inner join ";
                            ssampleqty += " (Select " + lotqty + " lotfactqty,AQLValue,AQL,CheckLevel,Ac,Re,s.Code from IQC_TestSTD105ECheckSet i inner join IHPS_QUALITY_SPC_AQLIS02859 s on i.Code=s.Code  where LotSizemin<=" + lotqty.ToString() + " and LotSizemax>=" + lotqty.ToString() + " and CheckLevel='II') s on c.Code=s.Code";
                            DataSet dssampleqty = DbAccess.SelectBySql(ssampleqty);
                            txtsampleqty.Text = dssampleqty.Tables[0].Rows[0]["Sampleqty"].ToString();

                            string txtCode = " Select a.Code from IQC_TestSTD105ECheckSet i inner join IHPS_QUALITY_SPC_AQLIS02859 a on i.Code=a.Code  where ( LotSizemin<=" + lotqty + " and LotSizemax>=" + lotqty + " and CheckLevel='II')";
                            DataTable dds = DbAccess.SelectBySql(txtCode).Tables[0];
                            string Code = dds.Rows[0]["Code"].ToString();

                            string sql = @"select Ac,Re from IHPS_QUALITY_SPC_AQLIS02859 where Type ='加严' and AQLValue = " + float.Parse(AQLValue) + " and Code = '" + Code + "'";

                            DataTable djt = DbAccess.SelectBySql(sql).Tables[0];

                            lblAQLAc.Text = "Ac=" + djt.Rows[0]["Ac"].ToString();
                            lblAQLRe.Text = "Re=" + djt.Rows[0]["Re"].ToString();
                        }
                        else
                        {
                            txtsampleqty.Text = "1";
                            lblAQLAc.Text = "Ac= 1";
                            lblAQLRe.Text = "Re= 1";
                        }
                    }
                    if (MaterialState == "放宽检验")
                    {
                        if (lotqty > 1)
                        {
                            string ssampleqty = @"Select case when Sampleqty>lotfactqty then lotfactqty else Sampleqty end as Sampleqty,s.Code from IQC_TestSTD105ECode c ";
                            ssampleqty += "  inner join ";
                            ssampleqty += " (Select " + lotqty + " lotfactqty,AQLValue,AQL,CheckLevel,Ac,Re,s.Code from IQC_TestSTD105ECheckSet i inner join IHPS_QUALITY_SPC_AQLIS02859 s on i.Code1=s.Code  where LotSizemin<=" + lotqty.ToString() + " and LotSizemax>=" + lotqty.ToString() + " and CheckLevel='II') s on c.Code=s.Code";
                            DataSet dssampleqty = DbAccess.SelectBySql(ssampleqty);
                            txtsampleqty.Text = dssampleqty.Tables[0].Rows[0]["Sampleqty"].ToString();

                            string txtCode = " select s.Code from IQC_TestSTD105ECheckSet i inner join IHPS_QUALITY_SPC_AQLIS02859 s on i.Code1=s.Code  where ( LotSizemin<=" + lotqty + " and LotSizemax>=" + lotqty + " and CheckLevel='II' )";
                            DataTable dds = DbAccess.SelectBySql(txtCode).Tables[0];
                            string Code = dds.Rows[0]["Code"].ToString();
                            string Value = AQLValue;
                            string sql = @"select Ac,Re from IHPS_QUALITY_SPC_AQLIS02859 where ( Type ='放宽' and  AQLValue = " + float.Parse(AQLValue) + " and Code = '" + Code + "')";
                            DataTable dft = DbAccess.SelectBySql(sql).Tables[0];

                            lblAQLAc.Text = "Ac=" + dft.Rows[0]["Ac"].ToString();
                            lblAQLRe.Text = "Re=" + dft.Rows[0]["Re"].ToString();
                        }
                        else
                        {
                            txtsampleqty.Text = "1";
                            lblAQLAc.Text = "Ac= 1";
                            lblAQLRe.Text = "Re= 1";
                        }
                    }
                    if (MaterialState == "正常检验")
                    {
                        if (lotqty > 1)
                        {
                            string ssampleqty = @"Select case when Sampleqty>lotfactqty then lotfactqty else Sampleqty end as Sampleqty,s.Code from IQC_TestSTD105ECode c ";
                            ssampleqty += "  inner join ";
                            ssampleqty += " (Select " + lotqty + " lotfactqty,AQLValue,AQL,CheckLevel,Ac,Re,s.Code from IQC_TestSTD105ECheckSet i inner join IHPS_QUALITY_SPC_AQLIS02859 s on i.Code=s.Code  where LotSizemin<=" + lotqty.ToString() + " and LotSizemax>=" + lotqty.ToString() + " and CheckLevel='II') s on c.Code=s.Code";
                            DataSet dssampleqty = DbAccess.SelectBySql(ssampleqty);
                            txtsampleqty.Text = dssampleqty.Tables[0].Rows[0]["Sampleqty"].ToString();

                            string txtCode = " Select a.Code from IQC_TestSTD105ECheckSet i inner join IHPS_QUALITY_SPC_AQLIS02859 a on i.Code=a.Code  where ( LotSizemin<=" + lotqty + " and LotSizemax>=" + lotqty + " and CheckLevel='II')";
                            DataTable dds = DbAccess.SelectBySql(txtCode).Tables[0];
                            string Code = dds.Rows[0]["Code"].ToString();

                            string sql = @"select Ac,Re from IHPS_QUALITY_SPC_AQLIS02859 where Type ='正常检验' and AQLValue = " + float.Parse(AQLValue) + " and Code = '" + Code + "'";

                            DataTable djt = DbAccess.SelectBySql(sql).Tables[0];

                            lblAQLAc.Text = "Ac=" + djt.Rows[0]["Ac"].ToString();
                            lblAQLRe.Text = "Re=" + djt.Rows[0]["Re"].ToString();
                        }
                        else
                        {
                            txtsampleqty.Text = "1";
                            lblAQLAc.Text = "Ac= 1";
                            lblAQLRe.Text = "Re= 1";
                        }
                    }

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
                    lblAQLAc.Text = "Ac=0";
                    lblAQLRe.Text = "Re=1";


                }
                else if (txtsampletype.Text == "全检")
                {
                    string scheckqty = "SELECT count(lotno) qty,SUM(qty) totalqty from delivery where  deliveryid='" + dsinfo.Tables[0].Rows[0]["deliveryid"].ToString() + "' and materialcode='" + dsinfo.Tables[0].Rows[0]["materialcode"].ToString() + "'";
                    DataTable dtqty = DbAccess.SelectBySql(scheckqty).Tables[0];
                    txtsampleqty.Text = dtqty.Rows[0][0].ToString();
                    /////txttotalqty.Text = dtqty.Rows[0][1].ToString();
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

                    txtremarks.Text = "上次测试时间为:" + sFlag.Substring(i + 1);
                    txtremarks.BackColor = Color.Yellow;
                   txtremarks.Enabled = false;
                    //return;
                }
                else if (sFlag.IndexOf("需重测") >= 0)
                {
                    MessageBox.Show(sFlag, "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    txtremarks.Text = "";
                    txtremarks.Enabled = true;

                    //20151105新增
                    txtlotno.Text = "";
                    txtlotno.Focus();
                    txtlotno.SelectAll();
                    return;
                    //20151105新增
                }
            }
            else
            {
                txttestsubitem.Focus();
                bindSubitemInfo();
                //return;
            }
        }


        private void txttestitem_SelectedValueChanged(object sender, EventArgs e)
        {
           /////////// txttestitem_SelectedIndexChanged(sender, e);
        }

        public DataTable dtpub = null;
        public DataTable dtother = null;
        public DataTable dtEMS = null;
        public DataSet dsinfo = null;
        int m = 0;
        string sreceptid = "", spcode = "";
        private void getInfo(string lotno)
        {
            this.dsinfo = null;
            //20150210新增,清空表,以免下一检验批次为有接收单号的批次,而没有清空杂项检验的批次信息
            this.dtother = null;
            m = 0;
            string sSql = "select org_id,deliveryid,d.materialcode,materialname,qty,vendorname,lot_number from delivery d left join MaterialSpec m on d.materialcode=m.materialcode where lotno='" + lotno + "'";
            //string sSql = "select org_id,deliveryid,materialcode,qty,vendorname,lot_number from delivery  where lotno='" + lotno + "'";
            DataSet ds = DbAccess.SelectBySql(sSql);
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                dsinfo = ds;
                //Int32 x=0;
                Int64 x = 0;
                //客供料下面的视图中没有
                if (Int64.TryParse(ds.Tables[0].Rows[0]["deliveryid"].ToString(), out x))
                {
                    string sDeliveryID = ds.Tables[0].Rows[0]["deliveryid"].ToString();
                    string sMaterialCode = ds.Tables[0].Rows[0]["materialcode"].ToString();
                    string sOraSql;
                    sreceptid = sDeliveryID;
                    spcode = sMaterialCode;

                    //取接收数量(状态为:RECEIVE)

                    sOraSql = " select ship_to_org_id,INVENTORY_ITEM_ID,transaction_id ,receipt_num, round(sum(quantity),0) primary_quantity, item_number,max(item_desc) item_desc,max(transaction_date) transaction_date,ORG_ID,transaction_type,ORGANIZATION_CODE,SHIP_TO_ORG_CODE,PARENT_TRANSACTION_TYPE,VENDOR_TYPE_LOOKUP_CODE ";
                    sOraSql += " from cux_inv_rec4_v where (transaction_type<>'ACCEPT' and transaction_type<>'REJECT') and receipt_num='" + sDeliveryID + "' and item_number='" + sMaterialCode + "'";
                    sOraSql += " group by ship_to_org_id,transaction_id,INVENTORY_ITEM_ID,receipt_num,item_number,ORG_ID,transaction_type,ORGANIZATION_CODE,SHIP_TO_ORG_CODE,PARENT_TRANSACTION_TYPE,VENDOR_TYPE_LOOKUP_CODE ";
                    if (ds.Tables[0].Rows[0]["org_id"].ToString() != "HCL")
                    {
                        DataSet dsERPDelivery = DbAccess.SelectByOracle(sOraSql);
                        if (dsERPDelivery != null && dsERPDelivery.Tables.Count > 0 && dsERPDelivery.Tables[0].Rows.Count > 0)
                        {
                            string sOraCheckType = "select TRANSACTION_TYPE_CODE ,FIRM_SHORT_CODE FROM APPS.CUX_INV_PROCESSING_V WHERE RECEIPT_NUM='" + sDeliveryID + "' and ITEM_NUM='" + sMaterialCode + "' and TRANSACTION_TYPE_CODE<>'REJECT' and TRANSACTION_TYPE_CODE<>'ACCEPT'";
                            DataSet dsCheckType =DbAccess.SelectByOracle(sOraCheckType);
                            if (dsCheckType != null && dsCheckType.Tables[0].Rows.Count > 0)
                            {  
                                if (dsCheckType.Tables[0].Rows[0]["TRANSACTION_TYPE_CODE"].ToString() == "RECEIVE" && dsCheckType.Tables[0].Rows[0]["FIRM_SHORT_CODE"].ToString() != "400")
                                {
                                    DataRow[] rw = dsERPDelivery.Tables[0].Select("transaction_type='RECEIVE' ");
                                    dtpub = dsERPDelivery.Tables[0].Clone();
                                    for (int i = 0; i < rw.Length; i++)
                                    {
                                        dtpub.Rows.Add(rw[i].ItemArray);
                                    }
                                    for (int i = 0; i < dtpub.Rows.Count; i++)
                                    {
                                        m = m + int.Parse(dtpub.Rows[i]["primary_quantity"].ToString());
                                    }
                                    txtstockqty.Text = m.ToString();
                                    //20170919新增
                                    txttotalqty.Text = m.ToString();
                                }
                                else if (dsCheckType.Tables[0].Rows[0]["TRANSACTION_TYPE_CODE"].ToString() == "TRANSFER" && dsCheckType.Tables[0].Rows[0]["FIRM_SHORT_CODE"].ToString() == "400")
                                {
                                    DataRow[] rw = dsERPDelivery.Tables[0].Select("transaction_type='TRANSFER' ");
                                    dtpub = dsERPDelivery.Tables[0].Clone();
                                    for (int i = 0; i < rw.Length; i++)
                                    {
                                        dtpub.Rows.Add(rw[i].ItemArray);
                                    }
                                    for (int i = 0; i < dtpub.Rows.Count; i++)
                                    {
                                        m = m + int.Parse(dtpub.Rows[i]["primary_quantity"].ToString());
                                    }
                                    txtstockqty.Text = m.ToString();
                                    txttotalqty.Text = m.ToString();
                                }
                                else
                                {
                                    lblinfo.Text = "单号" + sDeliveryID + ",料号:" + sMaterialCode + ",还没有做接收或转移,请先接收或转移后检验!";
                                    lblinfo.ForeColor = Color.Red;
                                    MessageBox.Show(lblinfo.Text, "提示", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                                    txtlotno.Text = "";
                                    txtlotno.Focus();
                                    return;
                                }
                            }
                            return;
                        }
                    }
                    else if (ds.Tables[0].Rows[0]["org_id"].ToString() == "HCL" && ds.Tables[0].Rows[0]["materialname"].ToString().Contains("HBTG"))
                    {
                        DataSet dsERPDelivery = DbAccess.SelectByOracle(sOraSql);
                        if (dsERPDelivery != null && dsERPDelivery.Tables.Count > 0 && dsERPDelivery.Tables[0].Rows.Count > 0)
                        {
                            string sOraCheckType = "select TRANSACTION_TYPE_CODE ,FIRM_SHORT_CODE FROM APPS.CUX_INV_PROCESSING_V WHERE RECEIPT_NUM='" + sDeliveryID + "' and ITEM_NUM='" + sMaterialCode + "' and TRANSACTION_TYPE_CODE<>'REJECT' and TRANSACTION_TYPE_CODE<>'ACCEPT'";
                            DataSet dsCheckType = DbAccess.SelectByOracle(sOraCheckType);
                            if (dsCheckType != null && dsCheckType.Tables[0].Rows.Count > 0)
                            {
                                if (dsCheckType.Tables[0].Rows[0]["TRANSACTION_TYPE_CODE"].ToString() == "RECEIVE" && dsCheckType.Tables[0].Rows[0]["FIRM_SHORT_CODE"].ToString() != "400")
                                {
                                    DataRow[] rw = dsERPDelivery.Tables[0].Select("transaction_type='RECEIVE' ");
                                    dtpub = dsERPDelivery.Tables[0].Clone();
                                    for (int i = 0; i < rw.Length; i++)
                                    {
                                        dtpub.Rows.Add(rw[i].ItemArray);
                                    }
                                    for (int i = 0; i < dtpub.Rows.Count; i++)
                                    {
                                        m = m + int.Parse(dtpub.Rows[i]["primary_quantity"].ToString());
                                    }
                                    txtstockqty.Text = m.ToString();
                                    //20170919新增
                                    txttotalqty.Text = m.ToString();
                                }
                                else if (dsCheckType.Tables[0].Rows[0]["TRANSACTION_TYPE_CODE"].ToString() == "TRANSFER" && dsCheckType.Tables[0].Rows[0]["FIRM_SHORT_CODE"].ToString() == "400")
                                {
                                    DataRow[] rw = dsERPDelivery.Tables[0].Select("transaction_type='TRANSFER' ");
                                    dtpub = dsERPDelivery.Tables[0].Clone();
                                    for (int i = 0; i < rw.Length; i++)
                                    {
                                        dtpub.Rows.Add(rw[i].ItemArray);
                                    }
                                    for (int i = 0; i < dtpub.Rows.Count; i++)
                                    {
                                        m = m + int.Parse(dtpub.Rows[i]["primary_quantity"].ToString());
                                    }
                                    txtstockqty.Text = m.ToString();
                                    txttotalqty.Text = m.ToString();
                                }
                                else
                                {
                                    lblinfo.Text = "单号" + sDeliveryID + ",料号:" + sMaterialCode + ",还没有做接收或转移,请先接收或转移后检验!";
                                    lblinfo.ForeColor = Color.Red;
                                    MessageBox.Show(lblinfo.Text, "提示", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                                    txtlotno.Text = "";
                                    txtlotno.Focus();
                                    return;
                                }
                            }
                            return;
                        }
                    }
                    else
                    {
                        m = 50000000;
                        return;
                    }
                    //20150325新增EMS客供料测试
                    #region
                    string EMSSql = " select deliveryid, max(item) transactionid, round(sum(qty),0) qty,m.materialcode, max(m.materialname) materialname,";
                    EMSSql += " max(unit) unit, max(vendorname) vendorname, max(vendorcode) vendorcode,max(INVENTORY_ITEM_ID) INVENTORY_ITEM_ID,max(org_id) org_id,'' lot_number ";
                    EMSSql += " from deliveryEMSOtherRec d left join OEM_EMSHYTCusRelation o on d.vendorcode=o.cuscode left join  materialspec m on o.hytcode=m.materialcode  where  deliveryid='" + sDeliveryID + "' and m.materialcode='" + sMaterialCode + "'";
                    EMSSql += " group by deliveryid,m.materialcode";
                    DataSet dsEMSOther = DbAccess.SelectBySql(EMSSql);
                    if (dsEMSOther != null && dsEMSOther.Tables.Count > 0 && dsEMSOther.Tables[0].Rows.Count > 0)
                    {
                        dtEMS = dsEMSOther.Tables[0];
                        m = int.Parse(dsEMSOther.Tables[0].Rows[0]["qty"].ToString());
                        txtstockqty.Text = dsEMSOther.Tables[0].Rows[0]["qty"].ToString();
                        txttotalqty.Text = m.ToString();
                    }
                    #endregion

                    else
                    {
                        lblinfo.Text = "单号" + sDeliveryID + ",料号:" + sMaterialCode + ",还没有生成批次,请先生成后检验!";
                        lblinfo.ForeColor = Color.Red;
                        MessageBox.Show(lblinfo.Text, "提示", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        txtlotno.Text = "";
                        txtlotno.Focus();
                        return;
                    }
                }
                else
                {
                    string ssql = "select max(deliveryid) deliveryid ,d.materialcode,sum(qty) qty,max(vendorname) vendorname,max(INVENTORY_ITEM_ID) INVENTORY_ITEM_ID,max(unit) unit,max(org_id) org_id,'' lot_number from delivery d inner join materialspec m on d.materialcode=m.materialcode where deliveryid='" + ds.Tables[0].Rows[0]["deliveryid"].ToString() + "' and d.materialcode='" + ds.Tables[0].Rows[0]["materialcode"].ToString() + "'  group by d.materialcode ";
                    DataSet dsREC = DbAccess.SelectBySql(ssql);
                    dsinfo = dsREC;
                    dtother = dsREC.Tables[0];
                    if (dtother.Rows.Count > 0)
                    m = int.Parse(dtother.Rows[0]["qty"].ToString());
                }
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

        private void getInfoNew(string lotno)
        {
            this.dsinfo = null;
            //20150210新增,清空表,以免下一检验批次为有接收单号的批次,而没有清空杂项检验的批次信息
            this.dtother = null;
            m = 0;
            //string sSql = "select deliveryid,d.materialcode,materialname,qty from delivery d left join MaterialSpec m on d.materialcode=m.materialcode where lotno='" + lotno + "'";
            string sSql = "select deliveryid,materialcode,qty,vendorname,lot_number from delivery  where lotno='" + lotno + "'";
            DataSet ds = DbAccess.SelectBySql(sSql);
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                dsinfo = ds;
                //Int32 x=0;
                Int64 x = 0;
                //客供料下面的视图中没有
                if (Int64.TryParse(ds.Tables[0].Rows[0]["deliveryid"].ToString(), out x))
                {
                    string sDeliveryID = ds.Tables[0].Rows[0]["deliveryid"].ToString();
                    string sMaterialCode = ds.Tables[0].Rows[0]["materialcode"].ToString();
                    string sOraSql;
                    //取接收数量(状态为:RECEIVE)

                    sOraSql = " select ship_to_org_id,INVENTORY_ITEM_ID,transaction_id ,receipt_num, round(sum(quantity),0) primary_quantity, item_number,max(item_desc) item_desc,max(transaction_date) transaction_date,ORG_ID,transaction_type,ORGANIZATION_CODE,SHIP_TO_ORG_CODE,PARENT_TRANSACTION_TYPE,VENDOR_TYPE_LOOKUP_CODE ";
                    sOraSql += " from cux_inv_rec4_v where (transaction_type<>'ACCEPT' and transaction_type<>'REJECT') and receipt_num='" + sDeliveryID + "' and item_number='" + sMaterialCode + "'";
                    //sOraSql += " from cux_inv_rec4_v where (transaction_type<>'ACCEPT') and receipt_num='" + sDeliveryID + "' and item_number='" + sMaterialCode + "'";
                    sOraSql += " group by ship_to_org_id,transaction_id,INVENTORY_ITEM_ID,receipt_num,item_number,ORG_ID,transaction_type,ORGANIZATION_CODE,SHIP_TO_ORG_CODE,PARENT_TRANSACTION_TYPE,VENDOR_TYPE_LOOKUP_CODE ";

                    DataSet dsERPDelivery = DbAccess.SelectByOracle(sOraSql);
                    if (dsERPDelivery != null && dsERPDelivery.Tables.Count > 0 && dsERPDelivery.Tables[0].Rows.Count > 0)
                    {
                        string sOraCheckType = "select TRANSACTION_TYPE_CODE ,FIRM_SHORT_CODE FROM APPS.CUX_INV_PROCESSING_V WHERE RECEIPT_NUM='" + sDeliveryID + "' and ITEM_NUM='" + sMaterialCode + "' and TRANSACTION_TYPE_CODE<>'REJECT' and TRANSACTION_TYPE_CODE<>'ACCEPT'";
                        DataSet dsCheckType = DbAccess.SelectByOracle(sOraCheckType);
                        if (dsCheckType != null && dsCheckType.Tables[0].Rows.Count > 0)
                        {
                            if (dsCheckType.Tables[0].Rows[0]["TRANSACTION_TYPE_CODE"].ToString() == "RECEIVE" && dsCheckType.Tables[0].Rows[0]["FIRM_SHORT_CODE"].ToString() != "400")
                            {
                                DataRow[] rw = dsERPDelivery.Tables[0].Select("transaction_type='RECEIVE' ");
                                dtpub = dsERPDelivery.Tables[0].Clone();
                                for (int i = 0; i < rw.Length; i++)
                                {
                                    dtpub.Rows.Add(rw[i].ItemArray);
                                }
                                for (int i = 0; i < dtpub.Rows.Count; i++)
                                {
                                    m = m + int.Parse(dtpub.Rows[i]["primary_quantity"].ToString());
                                }
                                txtstockqty.Text = m.ToString();
                                //20170919新增
                                txttotalqty.Text = m.ToString();
                            }
                            else if (dsCheckType.Tables[0].Rows[0]["TRANSACTION_TYPE_CODE"].ToString() == "TRANSFER" && dsCheckType.Tables[0].Rows[0]["FIRM_SHORT_CODE"].ToString() == "400")
                            {
                                DataRow[] rw = dsERPDelivery.Tables[0].Select("transaction_type='TRANSFER' ");
                                dtpub = dsERPDelivery.Tables[0].Clone();
                                for (int i = 0; i < rw.Length; i++)
                                {
                                    dtpub.Rows.Add(rw[i].ItemArray);
                                }
                                for (int i = 0; i < dtpub.Rows.Count; i++)
                                {
                                    m = m + int.Parse(dtpub.Rows[i]["primary_quantity"].ToString());
                                }
                                txtstockqty.Text = m.ToString();
                                txttotalqty.Text = m.ToString();
                            }
                            /*//增加判拒后,可以重新检验
                            else if (dsCheckType.Tables[0].Rows[0]["TRANSACTION_TYPE_CODE"].ToString() == "REJECT")
                            {
                                if (dsERPDelivery.Tables[0].Rows[0]["PARENT_TRANSACTION_TYPE"].ToString() == "TRANSFER" && dsCheckType.Tables[0].Rows[0]["FIRM_SHORT_CODE"].ToString() == "400")
                                {
                                    DataRow[] rw = dsERPDelivery.Tables[0].Select("PARENT_TRANSACTION_TYPE='TRANSFER' ");
                                    dtpub = dsERPDelivery.Tables[0].Clone();
                                    for (int i = 0; i < rw.Length; i++)
                                    {
                                        dtpub.Rows.Add(rw[i].ItemArray);
                                    }
                                    for (int i = 0; i < dtpub.Rows.Count; i++)
                                    {
                                        m = m + int.Parse(dtpub.Rows[i]["primary_quantity"].ToString());
                                    }
                                    txtstockqty.Text = m.ToString();
                                }
                                else if (dsERPDelivery.Tables[0].Rows[0]["PARENT_TRANSACTION_TYPE"].ToString() == "RECEIVE" && dsCheckType.Tables[0].Rows[0]["FIRM_SHORT_CODE"].ToString() != "400")
                                {
                                    DataRow[] rw = dsERPDelivery.Tables[0].Select("PARENT_TRANSACTION_TYPE='RECEIVE' ");
                                    dtpub = dsERPDelivery.Tables[0].Clone();
                                    for (int i = 0; i < rw.Length; i++)
                                    {
                                        dtpub.Rows.Add(rw[i].ItemArray);
                                    }
                                    for (int i = 0; i < dtpub.Rows.Count; i++)
                                    {
                                        m = m + int.Parse(dtpub.Rows[i]["primary_quantity"].ToString());
                                    }
                                    txtstockqty.Text = m.ToString();
                                }
                            }
                            //增加判拒后,可以重新检验
                            */
                            else
                            {
                                lblinfo.Text = "单号" + sDeliveryID + ",料号:" + sMaterialCode + ",还没有做接收或转移,请先接收或转移后检验!";
                                lblinfo.ForeColor = Color.Red;
                                MessageBox.Show(lblinfo.Text, "提示", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                                txtlotno.Text = "";
                                txtlotno.Focus();
                                return;
                            }
                        }
                        return;
                    }
                    //20150325新增EMS客供料测试
                    #region
                    string EMSSql = " select deliveryid, max(item) transactionid, round(sum(qty),0) qty,m.materialcode, max(m.materialname) materialname,";
                    EMSSql += " max(unit) unit, max(vendorname) vendorname, max(vendorcode) vendorcode,max(INVENTORY_ITEM_ID) INVENTORY_ITEM_ID,max(org_id) org_id,'' lot_number ";
                    EMSSql += " from deliveryEMSOtherRec d left join OEM_EMSHYTCusRelation o on d.vendorcode=o.cuscode left join  materialspec m on o.hytcode=m.materialcode  where  deliveryid='" + sDeliveryID + "' and m.materialcode='" + sMaterialCode + "'";
                    EMSSql += " group by deliveryid,m.materialcode";
                    DataSet dsEMSOther = DbAccess.SelectBySql(EMSSql);
                    if (dsEMSOther != null && dsEMSOther.Tables.Count > 0 && dsEMSOther.Tables[0].Rows.Count > 0)
                    {
                        dtEMS = dsEMSOther.Tables[0];
                        m = int.Parse(dsEMSOther.Tables[0].Rows[0]["qty"].ToString());
                        txtstockqty.Text = dsEMSOther.Tables[0].Rows[0]["qty"].ToString();
                        txttotalqty.Text = m.ToString();
                    }
                    #endregion

                    else
                    {
                        lblinfo.Text = "单号" + sDeliveryID + ",料号:" + sMaterialCode + ",还没有生成批次,请先生成后检验!";
                        lblinfo.ForeColor = Color.Red;
                        MessageBox.Show(lblinfo.Text, "提示", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        txtlotno.Text = "";
                        txtlotno.Focus();
                        return;
                    }
                }
                else
                {
                    string ssql = "select max(deliveryid) deliveryid ,d.materialcode,sum(qty) qty,max(vendorname) vendorname,max(INVENTORY_ITEM_ID) INVENTORY_ITEM_ID,max(unit) unit,max(org_id) org_id,'' lot_number from delivery d inner join materialspec m on d.materialcode=m.materialcode where deliveryid='" + ds.Tables[0].Rows[0]["deliveryid"].ToString() + "' and d.materialcode='" + ds.Tables[0].Rows[0]["materialcode"].ToString() + "'  group by d.materialcode ";
                    DataSet dsREC = DbAccess.SelectBySql(ssql);
                    dsinfo = dsREC;
                    dtother = dsREC.Tables[0];
                    if (dtother.Rows.Count > 0)
                        m = int.Parse(dtother.Rows[0]["qty"].ToString());
                }
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

        private string CheckState(string pcode)
        {
            string sstate = "";
            string ssql = "select States from  IQC_TestMaterialState where Materialcode='" + pcode + "'";
            DataTable dt = DbAccess.SelectBySql(ssql).Tables[0];
            if (dt.Rows.Count > 0)
                sstate = dt.Rows[0]["States"].ToString();
            return sstate;
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
                            p1.Properties.SizeMode = PictureSizeMode.Stretch ;

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
                            p1.Properties.SizeMode = PictureSizeMode.Stretch;

                            FileStream fs2 = new FileStream(pt[1].ToString(), FileMode.Open);
                            Image bt2 = Image.FromStream(fs2);
                            fs2.Close();
                            fs2.Dispose();
                            p2.Image = bt2;
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
                            p1.Properties.SizeMode = PictureSizeMode.Stretch;

                            FileStream fs2 = new FileStream(pt[1].ToString(), FileMode.Open);
                            Image bt2 = Image.FromStream(fs2);
                            fs2.Close();
                            fs2.Dispose();
                            p2.Image = bt2;
                            p2.Properties.SizeMode = PictureSizeMode.Zoom;

                            FileStream fs3 = new FileStream(pt[2].ToString(), FileMode.Open);
                            Image bt3 = Image.FromStream(fs3);
                            fs3.Close();
                            fs3.Dispose();
                            p3.Image = bt3;
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
                                    p1.Properties.SizeMode = PictureSizeMode.Stretch;

                                    FileStream fs2 = new FileStream(pt[1].ToString(), FileMode.Open);
                                    Image bt2 = Image.FromStream(fs2);
                                    fs2.Close();
                                    fs2.Dispose();
                                    p2.Image = bt2;
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
                                    p1.Properties.SizeMode = PictureSizeMode.Stretch;

                                    FileStream fs2 = new FileStream(pt[1].ToString(), FileMode.Open);
                                    Image bt2 = Image.FromStream(fs2);
                                    fs2.Close();
                                    fs2.Dispose();
                                    p2.Image = bt2;
                                    p2.Properties.SizeMode = PictureSizeMode.Zoom;

                                    FileStream fs3 = new FileStream(pt[2].ToString(), FileMode.Open);
                                    Image bt3 = Image.FromStream(fs3);
                                    fs3.Close();
                                    fs3.Dispose();
                                    p3.Image = bt3;
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
                string ssql = "select PNo, PName, userid, eventtime, sup, checklist, Remarks from MaterialBitPIC where PNo='" + pcode + "' order  by eventtime asc";
                DataSet ds = Common.DbAccess.SelectBySql(ssql);
                if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                {
                    int i = ds.Tables[0].Rows.Count;
                    string show = "";
                    lblpicdes.Text = "制造商:" + ds.Tables[0].Rows[0]["sup"].ToString() + ",检验项目:" + ds.Tables[0].Rows[0]["checklist"].ToString() + ",备注:" + ds.Tables[0].Rows[0]["Remarks"].ToString();
                    for (int j = 1; j < i; j++)
                    {
                        show =show +" 【"+ds.Tables[0].Rows[j]["Remarks"].ToString() +"】 ";
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
            DataSet ds = ic.SelectTestSamplePosition("查询", pcode, "", "", "", "", "", "", "", "", "", "", "", Login.username, "55");
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

        string MaterialSource( string Materialcode)
        {
            string materialsource = "";
            string sql = " select materialsource from IQC_ProxyCardConfig  where productcode = '" + Materialcode + "' ";
            DataTable  dt = DbAccess.SelectBySql(sql).Tables[0];

            int i = dt.Rows.Count;

            if (dt == null || dt.Rows.Count < 1)
            {
                materialsource = "供应商";
            }
            else
            {
                materialsource = dt.Rows[0]["materialsource"].ToString();

            }
            return materialsource;
        } 

        private void txtlotno_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13)
            {

                if (txtlotno.Text.Trim() == "" || txtmaterialcode.Text.Trim() == "")
                    return;

                if (txtlotno.Text.Trim() != "" || txtmaterialcode.Text.Trim() != "")
                    return;

                getInfo(txtlotno.Text.Trim());
               
                if (txtlotno.Text.Trim() == "")
                {
                    databind.DataSource = null;
                    DbAccess.SetControlEmpty(this);
                    return;
                }
                DataSet ds = new DataSet();
                DataSet dsother = dsinfo;

                if (dsother != null && dsother.Tables.Count > 0 && dsother.Tables[0].Rows.Count > 0)
                {
                    string ssql = "";
                    if (dsother.Tables[0].Rows[0]["deliveryid"].ToString().Contains("REC"))
                    {
                        ds = dsother;
                    }
                    else if (dtEMS != null && this.dtEMS.Rows[0]["deliveryid"].ToString().Length >= 10)
                    {
                        ds.Tables.Add(dtEMS.Copy());
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
                string pre = "";
                if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                {
                    lotqty = int.Parse(ds.Tables[0].Rows[0]["qty"].ToString());
                    txtlotqty.Text = ds.Tables[0].Rows[0]["qty"].ToString();

                    string spcode = ds.Tables[0].Rows[0]["materialcode"].ToString();
                    string s = CheckState(spcode);
                    MaterialState = s;
                    show_Checktype.Visible = true;
                    show_Checktype.Text = s;
                    //20140922修改,不用每次都去绑定图片及路径

                    if (ds.Tables[0].Rows[0]["materialcode"].ToString() != sproductcode)
                    {
                        //showpic(sproductcode);
                        //showsampledirectory(sproductcode);
                        showpic(ds.Tables[0].Rows[0]["materialcode"].ToString());
                        showsampledirectory(ds.Tables[0].Rows[0]["materialcode"].ToString());
                        showRevExcept(ds.Tables[0].Rows[0]["materialcode"].ToString());

                        //20170518新增频率 
                        string ssql = "select code from OQC_TypeDefine where Definetype='测试频率' and Definevalue='" + ds.Tables[0].Rows[0]["materialcode"].ToString() + "'";
                        DataTable dtpre = DbAccess.SelectBySql(ssql).Tables[0];
                        if (dtpre.Rows.Count > 0)
                            pre = dtpre.Rows[0]["code"].ToString();


                        if (s == "暂停检验")
                        {
                            MessageBox.Show(spcode + "物料处于:" + s + ",请SQE去审核,并改进质量", s, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            txtlotno.Focus();
                            txtlotno.SelectAll();
                            txtlotno.Text = "";
                            return;
                        }

                        else if (s == "加严检验")
                        {
         

                        }
                        else if (s == "放宽检验")
                        {
 
                        }

                    }                  
                    string sprog = "select Productcode,productname from IQC_TestProgSet where productcode='" + ds.Tables[0].Rows[0]["materialcode"].ToString() + "'";
                    DataSet dsprog = Common.DbAccess.SelectBySql(sprog);
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
                        if (!dsprog.Tables[0].Rows[0]["Productcode"].ToString().StartsWith("7"))
                        {

                            string sqlproxycard = "declare @pname varchar(max) set @pname='" + dsprog.Tables[0].Rows[0]["productname"].ToString() + "'";
                            sqlproxycard += "  select CONVERT(varchar(11), expirydate, 121) expirydate,CONVERT(varchar(11),GETDATE(), 121) systime,(case when materialsource='制造商' then '不需审核' else ISNULL(states,'未审') end) states,supplier,materialsource,brand,@pname pname from IQC_ProxyCardConfig ";
                            //sqlproxycard += "  where productcode='" + ds.Tables[0].Rows[0]["materialcode"].ToString() + "' and supplier like '%" + ds.Tables[0].Rows[0]["vendorname"].ToString() + "%'";
                            sqlproxycard += "  where productcode='" + ds.Tables[0].Rows[0]["materialcode"].ToString() + "' and supplier like '%" + ds.Tables[0].Rows[0]["vendorname"].ToString() + "%'";
                            DataTable dtproxycard = DbAccess.SelectBySql(sqlproxycard).Tables[0];

                            if (dtproxycard.Rows.Count > 0)
                            {
                                string sbrand = "OK";
                                for (int kk = 0; kk < dtproxycard.Rows.Count; kk++)
                                {
                                    DataRow[] rw = dtproxycard.Select("pname like '%" + dtproxycard.Rows[kk]["brand"].ToString() + "%'");
                                    if (rw.Length > 0)
                                    {
                                        for (int mm = 0; mm < rw.Length; mm++)
                                        {
                                            if (rw[mm]["states"].ToString() == "不需审核")
                                            {
                                                sbrand = "OK";
                                                continue;
                                            }
                                            else if (rw[mm]["states"].ToString() == "未审")
                                            {
                                                lblinfo.Text = dsprog.Tables[0].Rows[0]["Productcode"].ToString() + "代理证未审核!";
                                                lblinfo.ForeColor = Color.Red;
                                                txtlotno.Focus();
                                                txtlotno.SelectAll();
                                                sbrand = "NG";
                                                break;
                                            }
                                            else if (Convert.ToDateTime(rw[mm]["expirydate"].ToString()) < Convert.ToDateTime(rw[mm]["systime"].ToString()))
                                            {
                                                lblinfo.Text = dsprog.Tables[0].Rows[0]["Productcode"].ToString() + "代理证有效期:" + rw[mm]["expirydate"].ToString() + ",代理证已超期!";
                                                lblinfo.ForeColor = Color.Red;
                                                txtlotno.Focus();
                                                txtlotno.SelectAll();
                                                sbrand = "NG";
                                                break;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        continue;
                                    }
                                }

                                if (sbrand == "NG")
                                {
                                    lblinfo.Text = dsprog.Tables[0].Rows[0]["Productcode"].ToString() + "品牌不正确!";
                                    lblinfo.ForeColor = Color.Red;
                                    txtlotno.Focus();
                                    txtlotno.SelectAll();
                                    return;
                                }                              
                            }
                            else
                            {
                                string sqlproxycard1 = "declare @pname varchar(max) set @pname='" + dsprog.Tables[0].Rows[0]["productname"].ToString() + "'";
                                sqlproxycard1 += "  select CONVERT(varchar(11), expirydate, 121) expirydate,CONVERT(varchar(11),GETDATE(), 121) systime,(case when materialsource='制造商' then '不需审核' else ISNULL(states,'未审') end) states,supplier,brand,@pname pname from IQC_ProxyCardConfig ";
                                sqlproxycard1 += "  where productcode='" + ds.Tables[0].Rows[0]["materialcode"].ToString() + "' and materialsource='制造商'";
                                DataTable dtproxycard1 = DbAccess.SelectBySql(sqlproxycard1).Tables[0];
                                if (dtproxycard1.Rows.Count < 1)
                                {
                                    lblinfo.Text = dsprog.Tables[0].Rows[0]["Productcode"].ToString() + "还没有维护相应的代理证!";
                                    lblinfo.ForeColor = Color.Red;
                                    txtlotno.Focus();
                                    txtlotno.SelectAll();
                                    return;
                                }                            
                            }
                        }
                       
                        string materialsource = MaterialSource(dsprog.Tables[0].Rows[0]["Productcode"].ToString());
                        if (materialsource == "现货商")
                        {
                            lblsup.Text = materialsource+":" + ds.Tables[0].Rows[0]["vendorname"].ToString() + ",【频点:" + ds.Tables[0].Rows[0]["lot_number"].ToString() + "】" + "频率:" + pre;
                            lblsup.ForeColor = Color.Red;

                        }
                        else if (materialsource == "制造商")
                        {
                            lblsup.Text = materialsource + ":" + ds.Tables[0].Rows[0]["vendorname"].ToString() + ",【频点:" + ds.Tables[0].Rows[0]["lot_number"].ToString() + "】" + "频率:" + pre;
                            lblsup.ForeColor = Color.Blue;
                        }
                        else if (materialsource == "代理商")
                        {
                            lblsup.Text = materialsource + ":" + ds.Tables[0].Rows[0]["vendorname"].ToString() + ",【频点:" + ds.Tables[0].Rows[0]["lot_number"].ToString() + "】" + "频率:" + pre;
                            lblsup.ForeColor = Color.Blue;
                        }
                        else
                        {
                            lblsup.Text = "供应商:" + ds.Tables[0].Rows[0]["vendorname"].ToString() + ",【频点:" + ds.Tables[0].Rows[0]["lot_number"].ToString() + "】" + "频率:" + pre;
                            lblsup.ForeColor = Color.Blue;
                        }

                        if (dsprog.Tables[0].Rows[0]["productname"].ToString().Contains("F1专用") || dsprog.Tables[0].Rows[0]["productname"].ToString().Contains("F1客供"))
                        {
                            lblprodinfo.Text = "编码:" + dsprog.Tables[0].Rows[0]["Productcode"].ToString() + ",描述:" + dsprog.Tables[0].Rows[0]["productname"].ToString() + GetF1Supp(dsprog.Tables[0].Rows[0]["Productcode"].ToString());
                        }
                        else
                            lblprodinfo.Text = "编码:" + dsprog.Tables[0].Rows[0]["Productcode"].ToString() + ",描述:" + dsprog.Tables[0].Rows[0]["productname"].ToString();
                        try
                        {
                            if (dtpub.Rows[0]["VENDOR_TYPE_LOOKUP_CODE"].ToString().Contains("TEMPORARY"))
                                lblprodinfo.ForeColor = Color.Red;
                            else
                                lblprodinfo.ForeColor = Color.Blue;
                        }
                        catch
                        {
                            lblprodinfo.ForeColor = Color.Blue;
                        }
 
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
                                    bindSubitemInfo();
                                }
                            }
                            else
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
                                txttestitem.Focus();
                                lblinfo.Text = "";
                                //显示第一项目、第一子项目的信息
                                bindSubitemInfo();
                                //txtlotno.Leave += txtlotno_Leave;
                            }
                            lblinfo.Text = "";
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

        private void txtmaterialcode_Leave(object sender, EventArgs e)
        {



        }

        private void DelayMs(int Millisecond) //延迟系统时间，但系统又能同时能执行其它任务；//thread.sleep(int);程序会HOLD住;
        {
            System.Threading.Thread.Sleep(Millisecond);
        }

        private int dev = 0;
        public bool OpenInstrument(int addr)
        {
            dev = GPIB.ibdev(0, addr, 0, (int)GPIB.gpib_timeout.T1s, 1, 0);
            GPIB.ibclr(dev);
            return true;
        }
        public bool write(int addr, string strWrite)
        {
            GPIB.ibwrt(dev, strWrite, strWrite.Length);

            return true;
        }
        public bool read(int addr, ref string strRead)
        {
            StringBuilder str = new StringBuilder(100);

            GPIB.ibrd(dev, str, 100);
            strRead = str.ToString();
            return true;
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
        private void testresult_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (testresult.SelectedIndex == 0)
            {
                Badcategory.Text = "";
                Baddescribe.Text = "";
                Badcategory.Visible = false;
                Badcategory.Enabled = false;
                Baddescribe.Enabled = false;
                Baddescribe.Visible = false;
            }
            else if (testresult.SelectedIndex == 1)
            {
                Badcategory.Visible = true;
                Badcategory.Enabled = true;
                Baddescribe.Enabled = true;
                Baddescribe.Visible = true;
                Badcategory.Properties.DataSource = BadSituation();
                Badcategory.Properties.DisplayMember = "不良类别";
                Badcategory.Properties.ValueMember = "不良类别";
            }
            else if (testresult.SelectedIndex == 2)
            {
                Badcategory.Text = "";
                Baddescribe.Text = "";
                Badcategory.Enabled = false;
                Badcategory.Visible = false;
                Baddescribe.Enabled = false;
                Baddescribe.Visible = false;
            }
        }


        bool  Checkiffirst(string produc)
        {
            bool flagg = false;
            string ssql = "select 1 from  IQC_TestMaterialState where Materialcode='" + produc + "'";
            DataTable dt = DbAccess.SelectBySql(ssql).Tables[0];
            if (dt.Rows.Count > 0)
            {
                return true;
            }
                return flagg;
        }

        string ifDateCode = "否";
        private void txtExpiryDate_KeyUp(object sender, KeyEventArgs e)
        {

            if (e.KeyValue == 13 && txtExpiryDate.Text.Trim() != "")
            {
                DateTime dtDate;
                if (DateTime.TryParse(txtExpiryDate.Text, out dtDate))
                {
                    txtExpiryDate.Text = dtDate.ToString("yyyy-MM-dd");
                    ifDateCode = "是";
                }
                else
                {
                    MessageBox.Show("时间格式不正确", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    ifDateCode = "否";
                    txtExpiryDate.Text = "";
                    return;
                }
            }
        }


        private void btnOK_Click(object sender, EventArgs e)
        {
      
            if ("电阻类,电感类,电容类".Contains(testtype) && txttestsubitem.Text.Contains("测试"))
            {
                return;
            }

            if (txtlotno.Text.Trim() == "" || txtmaterialcode.Text.Trim() == "")
                return;

            if (txttestitem.Text.Contains("包装确认") && txtDateCode.Text == "")
            {
                if (txtExpiryDate.Text == "")
                {
                    lblinfo.Text = "请输入D/C相关的时间";
                    lblinfo.ForeColor = Color.Red;
                    txtExpiryDate.ReadOnly = false;
                    return;
                }
                if (ifDateCode == "否")
                {
                    lblinfo.Text = "输入的失效日期不正确";
                    lblinfo.ForeColor = Color.Red;
                    txtExpiryDate.ReadOnly = false;
                    return;
                }
            }

            if (txtExpiryDate.Text != "")
            {
                if (DateTime.Parse(DateTime.Now.ToString("yyyy-MM-dd")) > DateTime.Parse(txtExpiryDate.Text))
                {
                    DialogResult Rt = MessageBox.Show("物料超期", "警告", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                    if (DialogResult.Cancel == Rt)
                    {
                        return;
                    }
                }
            }

            if (MaterialState != "")
            {
                sCheckType = MaterialState;

            }

            string NGtype = Badcategory.Text.Trim();
            string baddescribe = Baddescribe.Text.Trim();
            string manufacturer = textmanufacturer.Text.Trim();

            string sstate = "NG";
            if (testresult.SelectedIndex == -1)
            {
                lblinfo.Text = "请选择一个结果";
                lblinfo.ForeColor = Color.Red;
                return;
            }
            if (testresult.SelectedIndex == 0)
            {
                sstate = "OK";
                NGtype = "";
                baddescribe = "";
            }
            else if (testresult.SelectedIndex == 1)
            {
                sstate = "NG";
                if ( baddescribe == "" || NGtype == "")
                {
                    MessageBox.Show("请选择不良类别和输入不良现象！", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
            }
            else if (testresult.SelectedIndex == 2)
                sstate = "NA";
            if (txtsampleqty.Text == txtsamplefactqty.Text) return;
            if (txtlotno.Text == "") return;

            string msg = "";

            try
            {
                //msg = ic.AddNewTestList("新增", testtype, txttestitem.SelectedValue.ToString(), txttestsubitem.SelectedValue.ToString(), txttestdes.Text, txttesttools.Text, txtPacktype.Text, txtsampletype.Text, float.Parse(txtscope2.Text == "" ? "0" : txtscope2.Text),
                //    float.Parse(txtscope1.Text == "" ? "0" : txtscope1.Text), float.Parse(AQLValue == "" ? "0" : AQLValue), txtlotno.Text, int.Parse(txtlotqty.Text), Login.username, int.Parse(txtsampleqty.Text), 1, sproductcode, txtAQL.Text, sstate, "", txtremarks.Text, 0, txttestvalue.Text,
                //    lblunit.Text == "" ? lblsetunit.Text : lblunit.Text, sCheckType, txtDateCode.Text, txtMdate.Text, txtExpiryDate.Text, txtrsno.SelectedValue == null ? "" : txtrsno.SelectedValue.ToString());

                msg = ic.AddNewTestListNew("新增", testtype, txttestitem.Text, txttestsubitem.Text, txttestdes.Text, txttesttools.Text, txtPacktype.Text, txtsampletype.Text, float.Parse(txtscope2.Text == "" ? "0" : txtscope2.Text),
                    float.Parse(txtscope1.Text == "" ? "0" : txtscope1.Text), float.Parse(AQLValue == "" ? "0" : AQLValue), txtlotno.Text, int.Parse(txtlotqty.Text), Login.username, int.Parse(txtsampleqty.Text), 1, sproductcode, txtAQL.Text, sstate, "", txtremarks.Text, 0,"",
                    lblunit.Text == "" ? lblsetunit.Text : lblunit.Text, sCheckType, txtDateCode.Text, txtMdate.Text, txtExpiryDate.Text, txtrsno.Text == null ? "" : txtrsno.Text, NGtype,baddescribe,manufacturer, 0,"");

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
                txtremarks.Text = "";
                txtDateCode.Text = "";
                testresult.SelectedIndex = -1;
                this.txtDateCode.Text = "";
                txtMdate.Text = "";
                txtExpiryDate.Text = "";
                textmanufacturer.Text = "";
                GetRSNO();
                txtExpiryDate.ReadOnly = true;
                ifDateCode = "否";
            }
            else
            {
                MessageBox.Show(msg, "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                DbAccess.SetControlEmpty(this);
                txtlotno.Focus();
            }

        }

        private void btndel_Click(object sender, EventArgs e)
        {
            DataTable dss = databind.DataSource as DataTable;
            if (  dss==null || dss.Rows.Count < 0)
            {
                return;
            }
            if (gridView.FocusedRowHandle < 0)
                return;
            if (gridView.GetSelectedRows().Length < 0)
                return;

            for (int i = gridView.GetSelectedRows().Length; i > 0; i--)
            {
                DataRow dr = gridView.GetDataRow(gridView.GetSelectedRows()[i - 1]);
                string k = ic.AddNewTestListNew("删除", dr["TestType"].ToString(), dr["TestItem"].ToString(), dr["TestSubItem"].ToString(), "", "", dr["PackType"].ToString(),
                                  "", 0, 0, 0, dr["LotNo"].ToString(), 0, Login.username, 0, 0, "", "", "", "", "", int.Parse(dr["Items"].ToString()), "", "", "", "", "", "", "","","","",0,"");

                if (k.IndexOf("OK") >= 0)
                {
                    //this.databind.Rows.RemoveAt(databind.SelectedRows[i - 1].Index);
                    gridView.DeleteRow(gridView.GetSelectedRows()[i - 1]);
                }
                    
            }

            txtsamplefactqty.Text = dss.Rows.Count.ToString();
            if (txtsampleqty.Text == txtsamplefactqty.Text)
            {      
                txtremarks.Enabled = false;
                btnOK.Enabled = false;
            }
            else
            {
                txtremarks.Enabled = true;
                btnOK.Enabled = true;
            }
        }

        private void btnsearch_Click(object sender, EventArgs e)
        {
            TestListbind(testtype, txttestitem.Text, txttestsubitem.Text, txtPacktype.Text, txtlotno.Text, "");
        }

        private void instockresult_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (instockresult.SelectedIndex == 0)
            {
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
            else if (instockresult.SelectedIndex == 1)
            {
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
        private DataTable Unfinished()
        {

            string sql = "select distinct TestSubItem from IQC_TestProgSet i where i.Productcode='" + dsinfo.Tables[0].Rows[0]["materialcode"].ToString() + "'";
            sql += " and TestSubItem not in(select distinct TestSubItem from IQC_TestList where Productcode='" + dsinfo.Tables[0].Rows[0]["materialcode"].ToString() + "' and receptid='" + dsinfo.Tables[0].Rows[0]["deliveryid"].ToString() + "' )";
            DataTable dt = Common.DbAccess.SelectBySql(sql).Tables[0];
            return dt;
        }
        protected int OtherReceiveInstock(string org_id, string sDelivery, string productcode)
        {
            int qty = 0;
            string ssql = "select isnull(sum(qty),0) qty from Warehouse_IQCCheck where org_id='" + org_id + "' and receptid='" + sDelivery + "' and productcode='" + productcode + "'";
            DataTable dt = Common.DbAccess.SelectBySql(ssql).Tables[0];
            if (dt.Rows.Count > 0)
                qty = int.Parse(dt.Rows[0]["qty"].ToString());
            return qty;
        }
        public string insertBarCode(string org_id, string item_id, string tran_id, Int32 pqty, Int32 enableqty, string baruser, string receptid, string productcode, string lotno, string states, string remarks, SqlConnection cn, SqlTransaction tran)
        {

            SqlCommand sqlCmd;
            SqlParameter sqlParm;

            sqlCmd = new SqlCommand();
            sqlCmd.CommandType = CommandType.StoredProcedure;
            sqlCmd.CommandText = "Warehouse_IQCCheckInputAdd";
            sqlCmd.Connection = cn;
            sqlCmd.Transaction = tran;
            sqlParm = new SqlParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "@org_id";
            sqlParm.SqlDbType = SqlDbType.VarChar;
            sqlParm.Value = org_id;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new SqlParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "@item_id";
            sqlParm.SqlDbType = SqlDbType.VarChar;
            sqlParm.Value = item_id;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new SqlParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "@tran_id";
            sqlParm.SqlDbType = SqlDbType.VarChar;
            sqlParm.Value = tran_id;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new SqlParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "@pqty";
            sqlParm.SqlDbType = SqlDbType.Int;
            sqlParm.Value = pqty;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new SqlParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "@enableqty";
            sqlParm.SqlDbType = SqlDbType.Int;
            sqlParm.Value = enableqty;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new SqlParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "@userid";
            sqlParm.SqlDbType = SqlDbType.VarChar;
            sqlParm.Value = baruser;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new SqlParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "@receptid";
            sqlParm.SqlDbType = SqlDbType.VarChar;
            sqlParm.Value = receptid;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new SqlParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "@productcode";
            sqlParm.SqlDbType = SqlDbType.VarChar;
            sqlParm.Value = productcode;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new SqlParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "@lotno";
            sqlParm.SqlDbType = SqlDbType.VarChar;
            sqlParm.Value = lotno;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new SqlParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "@states";
            sqlParm.SqlDbType = SqlDbType.VarChar;
            sqlParm.Value = states;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new SqlParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "@remarks";
            sqlParm.SqlDbType = SqlDbType.VarChar;
            sqlParm.Value = remarks;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new SqlParameter("@msg", SqlDbType.VarChar, 100);
            sqlParm.Direction = ParameterDirection.Output;
            sqlCmd.Parameters.Add(sqlParm);

            sqlCmd.ExecuteNonQuery();
            return sqlCmd.Parameters["@msg"].Value.ToString();

        }
       
        public string insertOracle(Int32 tran_id, Int32 pqty, string baruser, string p_comments, OracleConnection cn, OracleTransaction tran)
        {
            OracleCommand sqlCmd;
            OracleParameter sqlParm;

            sqlCmd = new OracleCommand();
            sqlCmd.CommandType = CommandType.StoredProcedure;
            sqlCmd.CommandText = "CUX_RCV_TRS_PKG.rcv_check";
            sqlCmd.Connection = cn;
            sqlCmd.Transaction = tran;


            sqlParm = new OracleParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "p_rcv_transaction_id";
            sqlParm.DbType = DbType.Int32;
            sqlParm.Value = tran_id;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new OracleParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "p_qty";
            sqlParm.DbType = DbType.Int32;
            sqlParm.Value = pqty;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new OracleParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "p_barcode_user";
            sqlParm.DbType = DbType.String;
            sqlParm.Value = baruser;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new OracleParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "p_comments";
            sqlParm.DbType = DbType.String;
            sqlParm.Value = p_comments;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new OracleParameter("X_ERR_MSG", OracleType.VarChar, 50);
            sqlParm.Direction = ParameterDirection.Output;
            sqlCmd.Parameters.Add(sqlParm);

            sqlCmd.ExecuteNonQuery();
            return sqlCmd.Parameters["X_ERR_MSG"].Value.ToString();

        }

        

        /*
        public string insertOracle(Int32 tran_id, Int32 pqty, string baruser, string p_comments, Oracle.ManagedDataAccess.Client.OracleConnection cn, Oracle.ManagedDataAccess.Client.OracleTransaction tran)
        {

            Oracle.ManagedDataAccess.Client.OracleCommand sqlCmd;
            Oracle.ManagedDataAccess.Client.OracleParameter sqlParm;

            sqlCmd = new Oracle.ManagedDataAccess.Client.OracleCommand();
            sqlCmd.CommandType = CommandType.StoredProcedure;
            sqlCmd.CommandText = "CUX_RCV_TRS_PKG.rcv_check";
            sqlCmd.Connection = cn;
            sqlCmd.Transaction = tran;


            sqlParm = new Oracle.ManagedDataAccess.Client.OracleParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "p_rcv_transaction_id";
            sqlParm.DbType = DbType.Int32;
            sqlParm.Value = tran_id;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new Oracle.ManagedDataAccess.Client.OracleParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "p_qty";
            sqlParm.DbType = DbType.Int32;
            sqlParm.Value = pqty;
            sqlCmd.Parameters.Add(sqlParm);


            sqlParm = new Oracle.ManagedDataAccess.Client.OracleParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "p_barcode_user";
            sqlParm.DbType = DbType.String;
            sqlParm.Value = baruser;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new Oracle.ManagedDataAccess.Client.OracleParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "p_comments";
            sqlParm.DbType = DbType.String;
            sqlParm.Value = p_comments;
            sqlCmd.Parameters.Add(sqlParm);
            sqlParm = new Oracle.ManagedDataAccess.Client.OracleParameter("X_ERR_MSG", Oracle.ManagedDataAccess.Client.OracleDbType.NVarchar2, 50);
            sqlParm.Direction = ParameterDirection.Output;
            sqlCmd.Parameters.Add(sqlParm);

            sqlCmd.ExecuteNonQuery();
            return sqlCmd.Parameters["X_ERR_MSG"].Value.ToString();

        }
        
        */
        
        public string insertOracleReject(Int32 tran_id, Int32 pqty, string p_inspection_code, Int32 p_reason_id, string p_comments, string baruser, OracleConnection cn, OracleTransaction tran)
        {

            OracleCommand sqlCmd;
            OracleParameter sqlParm;

            sqlCmd = new OracleCommand();
            sqlCmd.CommandType = CommandType.StoredProcedure;
            sqlCmd.CommandText = "CUX_RCV_TRS_PKG.rcv_reject";
            sqlCmd.Connection = cn;
            sqlCmd.Transaction = tran;


            sqlParm = new OracleParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "p_rcv_transaction_id";
            sqlParm.DbType = DbType.Int32;
            sqlParm.Value = tran_id;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new OracleParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "p_qty";
            sqlParm.DbType = DbType.Int32;
            sqlParm.Value = pqty;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new OracleParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "p_inspection_code";
            sqlParm.DbType = DbType.String;
            sqlParm.Value = p_inspection_code;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new OracleParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "p_reason_id";
            sqlParm.DbType = DbType.String;
            sqlParm.Value = p_reason_id;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new OracleParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "p_comments";
            sqlParm.DbType = DbType.String;
            sqlParm.Value = p_comments;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new OracleParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "p_barcode_user";
            sqlParm.DbType = DbType.String;
            sqlParm.Value = baruser;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new OracleParameter("x_err_msg", OracleType.VarChar, 50);
            sqlParm.Direction = ParameterDirection.Output;
            sqlCmd.Parameters.Add(sqlParm);

            sqlCmd.ExecuteNonQuery();
            return sqlCmd.Parameters["x_err_msg"].Value.ToString();

        }
        
        /*
        public string insertOracleReject(Int32 tran_id, Int32 pqty, string p_inspection_code, Int32 p_reason_id, string p_comments, string baruser, Oracle.ManagedDataAccess.Client.OracleConnection cn, Oracle.ManagedDataAccess.Client.OracleTransaction tran)
        {

            Oracle.ManagedDataAccess.Client.OracleCommand sqlCmd;
            Oracle.ManagedDataAccess.Client.OracleParameter sqlParm;

            sqlCmd = new Oracle.ManagedDataAccess.Client.OracleCommand();
            sqlCmd.CommandType = CommandType.StoredProcedure;
            sqlCmd.CommandText = "CUX_RCV_TRS_PKG.rcv_reject";
            sqlCmd.Connection = cn;
            sqlCmd.Transaction = tran;


            sqlParm = new Oracle.ManagedDataAccess.Client.OracleParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "p_rcv_transaction_id";
            sqlParm.DbType = DbType.Int32;
            sqlParm.Value = tran_id;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new Oracle.ManagedDataAccess.Client.OracleParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "p_qty";
            sqlParm.DbType = DbType.Int32;
            sqlParm.Value = pqty;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new Oracle.ManagedDataAccess.Client.OracleParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "p_inspection_code";
            sqlParm.DbType = DbType.String;
            sqlParm.Value = p_inspection_code;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new Oracle.ManagedDataAccess.Client.OracleParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "p_reason_id";
            sqlParm.DbType = DbType.String;
            sqlParm.Value = p_reason_id;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new Oracle.ManagedDataAccess.Client.OracleParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "p_comments";
            sqlParm.DbType = DbType.String;
            sqlParm.Value = p_comments;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new Oracle.ManagedDataAccess.Client.OracleParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "p_barcode_user";
            sqlParm.DbType = DbType.String;
            sqlParm.Value = baruser;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new Oracle.ManagedDataAccess.Client.OracleParameter("x_err_msg", Oracle.ManagedDataAccess.Client.OracleDbType.Varchar2, 50);
            sqlParm.Direction = ParameterDirection.Output;
            sqlCmd.Parameters.Add(sqlParm);

            sqlCmd.ExecuteNonQuery();
            return sqlCmd.Parameters["x_err_msg"].Value.ToString();

        }

        */

        private void gridView_CustomDrawCell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
        {
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

        private void p1_MouseClick(object sender, MouseEventArgs e)
        {
        
            if (this.p1.Width == 350)
            {
                this.p1.SetBounds(this.p1.Location.X, this.p1.Location.Y, 244, 180);
            }
            else
            {
                this.p1.SetBounds(this.p1.Location.X, this.p1.Location.Y, 350, 350);
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


        private void p2_DoubleClick(object sender, EventArgs e)
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

                            if (pt.Length == 2)
                            {
                                string filename = Path.GetFileName(pt[1].ToString());
                                OpenFile(floerPathold, filename);
                            }
                        }
                    }
                }
            }
        }

        private void p3_DoubleClick(object sender, EventArgs e)
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
                            if (pt.Length == 3)
                            {
                                string filename = Path.GetFileName(pt[2].ToString());
                                OpenFile(floerPathold, filename);
                            }
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
        private DataTable IQCCheckResultToERP(string lotno)
        {
            string sql = " declare @recept varchar(30),@p varchar(30)";
            sql += "select @recept=receptid,@p=Productcode from IQC_TestList where LotNo='" + lotno + "'";
            sql += " SELECT deliveryid,item, materialcode,d.lotno,(d.qty-isnull(finOK.qty,0)-isnull(finNG.qty,0)) qty,finOK.qty OKqty,finNG.qty NGqty,org_id,iqcTransactionid,'' TestResult,(d.qty-isnull(finOK.qty,0)-isnull(finNG.qty,0)) checkqty  from delivery d left join";
            sql += "(select distinct lotno,productcode,receptid,states,isnull(sum(qty),0) qty from deliveryCheck t where receptid=@recept and productcode=@p and states='OK' group by lotno,productcode,receptid,states";
            sql += ") finOK on d.lotno=finOK.lotno and d.materialcode=finOK.productcode and d.deliveryid=finOK.receptid  left join ";
            sql += "(select lotno,productcode,receptid,states,isnull(sum(qty),0) qty from deliveryCheck t where receptid=@recept and productcode=@p and states='NG' group by lotno,productcode,receptid,states";
            sql += ") finNG on d.lotno=finNG.lotno and d.materialcode=finNG.productcode and d.deliveryid=finNG.receptid  ";
            sql += " where deliveryid=@recept and materialcode=@p ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            return dt;
        }
        private void btninstock_Click(object sender, EventArgs e)
        {
            string states = "OK";
            if (instockresult.SelectedIndex == -1)
            {
                this.lblinfo.ForeColor = Color.Red;
                lblinfo.Text = "请选择一个结果项";
                return;
            }
            if (instockresult.SelectedIndex == 0)
                states = "OK";
            else if (instockresult.SelectedIndex == 1)
                states = "NG";

            int output = 0;
            if (!int.TryParse(txtstockqty.Text.Trim(), out output))
            {
                this.lblinfo.ForeColor = Color.Red;
                lblinfo.Text = "批次数量输入错误，请重新输入";
                txtstockqty.Text = "";
                txtstockqty.Focus();
                return;
            }

            if (int.Parse(txtstockqty.Text) > m)
            {
                MessageBox.Show("此次检查数:" + txtstockqty.Text + ",超过还可检验数:" + m.ToString(), "提示", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                txtstockqty.Text = "";
                txtstockqty.Focus();
                return;
            }
            if (int.Parse(txtstockqty.Text) == 0)
            {
                txtstockqty.Text = "";
                MessageBox.Show(lblinfo.Text, "提示", MessageBoxButtons.OK, MessageBoxIcon.Stop);

                return;
            }
            if (lblprodinfo.Text.ToUpper().Contains("D1专用") || lblprodinfo.Text.ToUpper().Contains("D1客供") || lblprodinfo.Text.ToUpper().Contains("D1R专用") || lblprodinfo.Text.ToUpper().Contains("D1R客供"))
            {
                if (txtmadecode.Text.Trim() == "")
                {
                    MessageBox.Show("原厂料号不能为空!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    txtmadecode.Text = "";
                    txtmadecode.Focus();
                    txtremarks.Text = "";
                    return;
                }
                else if (!lblprodinfo.Text.ToUpper().Contains(txtmadecode.Text.Trim().ToUpper()))
                {
                    MessageBox.Show(txtmadecode.Text.Trim() + ",原厂料号不正确", "提示", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    txtmadecode.Text = "";
                    txtmadecode.Focus();
                    txtremarks.Text = "";
                    return;
                }
                else
                {
                    txtremarks.Text = txtremarks.Text + "原厂料号:" + txtmadecode.Text;
                }
            }
            //20160928新增加by husanbao
            else if (sproductcode.StartsWith("5") && (!lblprodinfo.Text.ToUpper().Contains("【F1")) && "电阻类,电感类,电容类,发光二极管,三极管,IC/BGA".Contains(testtype))
            {
                //只对对讲机产品进行管控
                if (txtmadecode.Text.Trim().ToUpper().StartsWith("P"))
                    txtmadecode.Text = txtmadecode.Text.Trim().Substring(1);
                else if (txtmadecode.Text.Trim().ToUpper().StartsWith("1P"))
                    txtmadecode.Text = txtmadecode.Text.Trim().Substring(2);
                if (txtmadecode.Text.Trim() == "")
                {
                    MessageBox.Show("原厂料号不能为空!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    txtmadecode.Text = "";
                    txtmadecode.Focus();
                    return;
                }
                else if (!lblprodinfo.Text.ToUpper().Contains(txtmadecode.Text.Trim().ToUpper()))
                {
                    DialogResult Rt = MessageBox.Show(txtmadecode.Text.Trim() + ",原厂料号不正确,是否继续?", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                    if (DialogResult.Cancel == Rt)
                    {
                        txtmadecode.Text = "";
                        txtmadecode.Focus();
                        txtremarks.Text = "";
                        return;
                    }
                    else
                    {
                        txtremarks.Text = txtremarks.Text + "原厂料号:" + txtmadecode.Text;
                    }
                }
            }
            else if (lblprodinfo.Text.ToUpper().Contains("专用") && (!lblprodinfo.Text.ToUpper().Contains("【F1")) && "电阻类,电感类,电容类,发光二极管,三极管,IC/BGA".Contains(testtype))
            {
                //只对对讲机产品进行管控
                if (txtmadecode.Text.Trim().ToUpper().StartsWith("P"))
                    txtmadecode.Text = txtmadecode.Text.Trim().Substring(1);
                else if (txtmadecode.Text.Trim().ToUpper().StartsWith("1P"))
                    txtmadecode.Text = txtmadecode.Text.Trim().Substring(2);
                if (txtmadecode.Text.Trim() == "")
                {
                    MessageBox.Show("原厂料号不能为空!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    txtmadecode.Text = "";
                    txtmadecode.Focus();
                    return;
                }
                else if (!lblprodinfo.Text.ToUpper().Contains(txtmadecode.Text.Trim().ToUpper()))
                {
                    DialogResult Rt = MessageBox.Show(txtmadecode.Text.Trim() + ",原厂料号不正确,是否继续?", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                    if (DialogResult.Cancel == Rt)
                    {
                        txtmadecode.Text = "";
                        txtmadecode.Focus();
                        txtremarks.Text = "";
                        return;
                    }
                    else
                    {
                        txtremarks.Text = txtremarks.Text + "原厂料号:" + txtmadecode.Text;
                    }
                }
            }
            //20160928新增加by husanbao

            //20150114新增加的,当所有检验项目检完后才能入库
            //if ((lblprodinfo.Text.Contains("D1专用") || lblprodinfo.Text.Contains("D1客供")) && "电阻类,电感类,电容类".Contains(testtype))
            //{
            //    //不执行操作(20160318新增加的)
            //    ;
            //}
            //else
            //{
            DataTable dtunfinish = Unfinished();
            if (dtunfinish.Rows.Count > 0)
            {
                string s = "";
                for (int i = 0; i < dtunfinish.Rows.Count; i++)
                {
                    s += dtunfinish.Rows[i][0].ToString() + ";";
                }
                MessageBox.Show(s.TrimEnd(';'), "提示,框中的项目还没有检验完,不能入库!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            string sqlorg = "select org_id from delivery where lotno='" + txtlotno.Text + "'";
            DataTable dtorg = DbAccess.SelectBySql(sqlorg).Tables[0];
            if (dtorg.Rows[0][0].ToString() == "HCL")
            {
                DataTable dtToERP = IQCCheckResultToERP(txtlotno.Text);
                TestResultCheck TRC = new TestResultCheck(dtToERP);
                TRC.ShowDialog();
                txtmadecode.Text = "";
                txtstockqty.Text = "";
                return;
            }
            //}
            //新增加的杂项入库数量20141223

            if (dtother != null && dtother.Rows.Count > 0)
            {
                //SqlConnection conn1 = new SqlConnection(DbAccess.connSql);
                //if (conn1.State == ConnectionState.Closed)
                //conn1.Open();

                //SqlTransaction tran2 = conn1.BeginTransaction();
                int stockqty = int.Parse(txtstockqty.Text);
                string org_id = dtother.Rows[0]["org_id"].ToString();
                string srecept = dtother.Rows[0]["deliveryid"].ToString();
                string productcode = dtother.Rows[0]["materialcode"].ToString();
                string item_id = dtother.Rows[0]["INVENTORY_ITEM_ID"].ToString();
                int alreadyqty = OtherReceiveInstock(org_id, srecept, productcode);
                string enableqty = dtother.Rows[0]["qty"].ToString();

                //还可入库数量(大于将要入库数量)
                if (int.Parse(enableqty == "" ? "0" : enableqty) - alreadyqty >= int.Parse(txtstockqty.Text))
                //if (int.Parse(txttotalqty.Text==""?txtlotqty.Text:txttotalqty.Text) - alreadyqty >= int.Parse(txtstockqty.Text))
                {
                    IQC ic = new IQC();
                    //string barcodemsg = insertBarCode(org_id, item_id, productcode, stockqty, stockqty,
                    //Login.username, srecept, productcode, txtlotno.Text, states, txtremarks.Text, conn1, tran2);
                    string barcodemsg = ic.IQC_TestOtherInstock(org_id, item_id, productcode, srecept, productcode, stockqty.ToString(), enableqty/*txtlotqty.Text*/, Login.username, txtlotno.Text, states, txtremarks.Text);
                    if (barcodemsg.IndexOf("成功") >= 0)
                    {
                        lblinfoinstock.ForeColor = Color.Blue;
                        lblinfoinstock.Text = barcodemsg;
                        txtstockqty.Text = "";
                        this.txtremarks.Text = "";
                    }
                    else
                    {
                        return;
                    }
                    return;
                }
                else
                {
                    lblinfoinstock.ForeColor = Color.Red;
                    //lblinfoinstock.Text = "超过了最大可入库数量:" + (int.Parse(txttotalqty.Text) - alreadyqty).ToString();
                    lblinfoinstock.Text = "超过了最大可入库数量:" + (int.Parse(enableqty == "" ? "0" : enableqty) - alreadyqty).ToString();
                    return;
                }
            }


            //新增加的EMS客供料入库数量20150325

            else if (dtEMS != null && dtEMS.Rows.Count > 0)
            {
                //SqlConnection conn1 = new SqlConnection(DbAccess.connSql);
                //if (conn1.State == ConnectionState.Closed)
                //conn1.Open();

                //SqlTransaction tran2 = conn1.BeginTransaction();
                int stockqty = int.Parse(txtstockqty.Text);
                string org_id = dtEMS.Rows[0]["org_id"].ToString();
                string srecept = dtEMS.Rows[0]["deliveryid"].ToString();
                string productcode = dtEMS.Rows[0]["materialcode"].ToString();
                string item_id = dtEMS.Rows[0]["INVENTORY_ITEM_ID"].ToString();
                int alreadyqty = OtherReceiveInstock(org_id, srecept, productcode);
                string enableqty = dtEMS.Rows[0]["qty"].ToString();

                //还可入库数量(大于将要入库数量)
                if (int.Parse(enableqty == "" ? "0" : enableqty) - alreadyqty >= int.Parse(txtstockqty.Text))
                //if (int.Parse(txttotalqty.Text == "" ? txtlotqty.Text : txttotalqty.Text) - alreadyqty >= int.Parse(txtstockqty.Text))
                {
                    IQC ic = new IQC();
                    string barcodemsg = ic.IQC_TestOtherInstock(org_id, item_id, productcode, srecept, productcode, stockqty.ToString(), enableqty/*txtlotqty.Text*/, Login.username, txtlotno.Text, states, txtremarks.Text);
                    if (barcodemsg.IndexOf("成功") >= 0)
                    {
                        lblinfoinstock.ForeColor = Color.Blue;
                        lblinfoinstock.Text = barcodemsg;
                        txtstockqty.Text = "";
                        this.txtremarks.Text = "";
                    }
                    else
                    {
                        return;
                    }
                    return;
                }
                else
                {
                    lblinfoinstock.ForeColor = Color.Red;
                    //lblinfoinstock.Text = "超过了最大可入库数量:" + (int.Parse(txttotalqty.Text == "" ? txtlotqty.Text : txttotalqty.Text) - alreadyqty).ToString();
                    lblinfoinstock.Text = "超过了最大可入库数量:" + (int.Parse(enableqty == "" ? "0" : enableqty) - alreadyqty).ToString();
                    return;
                }
            }


            DataTable dt = dtpub;

            SqlConnection conn = new SqlConnection(DbAccess.connSql);
           // Oracle.ManagedDataAccess.Client.OracleConnection orac = new Oracle.ManagedDataAccess.Client.OracleConnection(DbAccess.connOral);
           OracleConnection orac = new OracleConnection(DbAccess.connOral);
            if (conn.State == ConnectionState.Closed)
                conn.Open();
            if (orac.State == ConnectionState.Closed)
                orac.Open();

            SqlTransaction tran1 = conn.BeginTransaction();
            OracleTransaction oratran = orac.BeginTransaction();
           // Oracle.ManagedDataAccess.Client.OracleTransaction oratran = orac.BeginTransaction();

            try 
            {

                Int32 org_id = 0, item_id = 0, tran_id = 0;
                Int64 recepid = 0;
                string org_idnew = "";
                int instockqty = 0;
                int restenableqty = int.Parse(txtstockqty.Text);
                string msg = "", barcodemsg = "", msgcurrent = "", barcodemsgcurrent = "";
                for (int j = 0; j < dt.Rows.Count; j++)
                {
                    int alreadyqty = 0;
                    org_idnew = dt.Rows[j]["ORGANIZATION_CODE"].ToString();
                    org_id = Int32.Parse(dt.Rows[j]["ship_to_org_id"].ToString());
                    item_id = Int32.Parse(dt.Rows[j]["INVENTORY_ITEM_ID"].ToString());
                    tran_id = Int32.Parse(dt.Rows[j]["transaction_id"].ToString());
                    //recepid = Int32.Parse(dt.Rows[0]["receipt_num"].ToString());
                    recepid = Int64.Parse(dt.Rows[0]["receipt_num"].ToString());

                    //判断当前行可检验数量大于界面录入数量,则把所有数量录入到第一笔.
                    if (int.Parse(dt.Rows[j]["primary_quantity"].ToString()) >= restenableqty)
                    {
                        //则把用户输入的检验数量作为此行的检验数
                        instockqty = restenableqty;
                        //20150312新增
                        //x表示还在接口表中的数量
                        //int x = ERPinterface(tran_id.ToString());
                        //用临时变量记录此次操作数量
                        int x = alreadyqty;
                        //y表示还可处理的数量
                        int y = int.Parse(dt.Rows[j]["primary_quantity"].ToString()) - x;
                        if (y >= restenableqty)
                            instockqty = restenableqty;
                        else if (y > 0 && y < restenableqty)
                            instockqty = y;
                        else
                            continue;
                        //20150312新增

                        barcodemsg = barcodemsg + insertBarCode(org_idnew.ToString(), item_id.ToString(), tran_id.ToString(), instockqty, Int32.Parse(dt.Rows[j]["primary_quantity"].ToString()),
                                                                 Login.username, recepid.ToString(), dt.Rows[j]["item_number"].ToString(), txtlotno.Text, states, txtremarks.Text, conn, tran1);
                        if (barcodemsg.IndexOf("成功") >= 0)
                        {
                            if (states == "OK")
                            {
                                msgcurrent = insertOracle(tran_id, instockqty, Login.username, txtremarks.Text, orac, oratran);
                                if (msgcurrent.IndexOf("Error") >= 0)
                                {
                                    msg = msg + msgcurrent;

                                    break;
                                }
                                else
                                {
                                    msg = msg + msgcurrent;
                                    txtstockqty.Text = "0";
                                    break;
                                }
                            }
                            else if (states == "NG")
                            {
                                msgcurrent = insertOracleReject(tran_id, instockqty, "BAD", 31, txtremarks.Text, Login.username, orac, oratran);
                                if (msgcurrent.IndexOf("Error") >= 0)
                                {
                                    msg = msg + msgcurrent;
                                    break;
                                }
                                else
                                {
                                    msg = msg + msgcurrent;
                                    txtstockqty.Text = "0";
                                    break;
                                }
                            }
                        }
                        else
                        {
                            //防止第一笔已入库(一批次分多次入库问题),要继续循环下一笔
                            continue;
                        }
                    }
                    //如果当前行的数量小于录入数量,则把当前行先进行检验通过该行的数量,同时把总共要检验的数量-第一笔已录的=作为剩下的可录入数量，再进行下一次循环
                    else
                    {
                        instockqty = int.Parse(dt.Rows[j]["primary_quantity"].ToString());
                        //20150312新增
                        //x表示还在接口表中的数量
                        //int x = ERPinterface(tran_id.ToString());
                        //用临时变量记录此次操作数量

                        DataRow[] rwalready = dtalready.Select(" id='" + tran_id + "'");
                        if (rwalready.Length > 0)
                        {
                            for (int z = 0; z < rwalready.Length; z++)
                            {
                                alreadyqty = alreadyqty + int.Parse(rwalready[z]["qty"].ToString());
                            }
                        }
                        int x = alreadyqty;
                        //y表示还可处理的数量
                        int y = int.Parse(dt.Rows[j]["primary_quantity"].ToString()) - x;
                        if (y >= restenableqty)
                            instockqty = restenableqty;
                        else if (y > 0 && y < restenableqty)
                            instockqty = y;
                        else
                            continue;
                        //20150312新增

                        barcodemsgcurrent = insertBarCode(org_idnew.ToString(), item_id.ToString(), tran_id.ToString(), instockqty, instockqty,
                                                                Login.username, recepid.ToString(), dt.Rows[j]["item_number"].ToString(), txtlotno.Text, states, txtremarks.Text, conn, tran1);
                        barcodemsg = barcodemsg + barcodemsgcurrent;
                        if (barcodemsgcurrent.IndexOf("成功") >= 0)
                        {
                            if (states == "OK")
                            {
                                msgcurrent = insertOracle(tran_id, instockqty, Login.username, txtremarks.Text, orac, oratran);
                                if (msgcurrent.IndexOf("Error") >= 0)
                                {
                                    msg = msg + msgcurrent;
                                    break;
                                }
                                msg = msg + msgcurrent;
                                restenableqty = restenableqty - instockqty;
                                txtstockqty.Text = restenableqty.ToString();

                                DataRow newRow;
                                newRow = dtalready.NewRow();
                                newRow["id"] = tran_id;
                                newRow["qty"] = instockqty.ToString();
                                dtalready.Rows.Add(newRow);

                                continue;
                            }
                            else if (states == "NG")
                            {
                                msgcurrent = insertOracleReject(tran_id, instockqty, "BAD", 31, txtremarks.Text, Login.username, orac, oratran);
                                if (msgcurrent.IndexOf("Error") >= 0)
                                {
                                    msg = msg + msgcurrent;
                                    break;
                                }
                                msg = msg + msgcurrent;
                                restenableqty = restenableqty - instockqty;
                                txtstockqty.Text = restenableqty.ToString();

                                DataRow newRow;
                                newRow = dtalready.NewRow();
                                newRow["id"] = tran_id;
                                newRow["qty"] = instockqty.ToString();
                                dtalready.Rows.Add(newRow);
                                continue;
                            }
                        }
                        else
                        {
                            //防止第一笔已入库(一批次分多次入库问题),要继续循环下一笔
                            continue;
                        }
                    }
                }

                if (msg.IndexOf("Error") >= 0)
                {
                    this.lblinfoinstock.ForeColor = Color.Red;
                    lblinfoinstock.Text = msg + ";ERP写入失败!";
                    tran1.Rollback();
                    oratran.Rollback();
                }
                else if (barcodemsg.IndexOf("失败") >= 0)
                {
                    lblinfoinstock.ForeColor = Color.Red;
                    lblinfoinstock.Text = barcodemsg + ";ERP写入失败!";
                    tran1.Rollback();
                    oratran.Rollback();
                }
                else
                {
                    lblinfoinstock.ForeColor = Color.Blue;
                    txtstockqty.Text = "";
                    this.txtremarks.Text = "";
                    tran1.Commit();
                    oratran.Commit();
                    //lblinfoinstock.Text = barcodemsg + ";" + getdeliveryinfo(recepid.ToString(), dt.Rows[0]["item_number"].ToString());
                    lblinfoinstock.Text = states + ";" + getdeliveryinfo(recepid.ToString(), dt.Rows[0]["item_number"].ToString());

                    //dtpub = null;
                    txtmadecode.Text = "";
                    txtstockqty.Text = "";
                }
            }
            catch (Exception ex)
            {
                lblinfoinstock.Text = ex.ToString();
                tran1.Rollback();
                oratran.Rollback();
            }
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


        private void TestListcs_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (ssp.IsOpen)
                ssp.Close();
            this.Dispose();
            this.Close();
        }
        void TestList_MouseWheel(object sender, MouseEventArgs e)
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
        private void TestListcs_Load(object sender, EventArgs e)
        {
            lblsup.Text = "";
            lblprodinfo.Text = "";
            lblinfo.Text = "";
           // label3.Text = "";
            this.MouseWheel += new MouseEventHandler(TestList_MouseWheel);
            //labelControl1.Enabled = false;
           // labelControl1.Visible = false;
            Badcategory.Enabled = false;
            Badcategory.Visible = false;
           // labelControl2.Enabled = false;
           // labelControl2.Visible = false;
            Baddescribe.Enabled = false;
            Baddescribe.Visible = false;
            testresult.SelectedIndex = -1;
        }
    }
}