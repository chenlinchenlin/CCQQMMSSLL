using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using DevExpress.XtraBars.Helpers;
using DevExpress.XtraEditors;
using DevExpress.XtraBars.Ribbon;
using DX_QMS.SystemConfig;
using DevExpress.XtraNavBar;
using DX_QMS.SystemConfig;
using System.Net.Sockets;
using System.Net;
using IQCReport;
using DX_QMS.Common;
using DX_QMS.Suggestions;
using DX_QMS.KPI;
using DevExpress.XtraBars;
using DevExpress.XtraGrid.Demos;
using DX_QMS.IPQC;
using DX_QMS.IQCFilePosition;
using DX_QMS.SMTFolder;

namespace DX_QMS
{
    public partial class Form_QMS : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        [DllImport("user32", EntryPoint = "SetParent", ExactSpelling = true, CharSet = CharSet.Ansi, SetLastError = true)]
        public static extern int SetParent(IntPtr hWndChild, IntPtr hWndNewParent);
        private const int GWL_STYLE = -16;
        private const int WS_CHILD = 0x40000000;  //设置窗口属性为child
        [DllImport("user32.dll", EntryPoint = "GetWindowLong")]
        public static extern int GetWindowLong(IntPtr hwnd, int nIndex);
        [DllImport("user32.dll", EntryPoint = "SetWindowLong")]
        public static extern int SetWindowLong(IntPtr hwnd, int nIndex, int dwNewLong);
        string groupname = Login.groupId;
        IQC ic = new IQC();

        public Form_QMS()
        {
            InitializeComponent();
            //searchAllmenu();
        }
        void showform(object sender, DevExpress.XtraBars.ItemClickEventArgs e,string strTitle, string strTag, RibbonForm frmBase, Image image)
        {
            bool resetWindowStyle = false;
            if (resetWindowStyle)
                LayoutMdi();
            if (string.IsNullOrEmpty(e.Item.Tag as string)) return;
            if (!ShowOpenedPage(e.Item.Tag as string))
            {
                CreateMDIControl( strTitle,  strTag, frmBase, image);
            }
        }

        void frm_FormClosing(object sender, FormClosingEventArgs e)
        {
        }

        System.Windows.Forms.MdiLayout MdiLayout;
        bool IsTabbedMdi
        {
            get
            {
                return barTabbedMdi.Down;
            }
        }
        public void LayoutMdi()
        {
                LayoutMdi(MdiLayout);
        }
        public void CreateMDIControl(string strTitle, string strTag, RibbonForm frmBase, Image image)
        {
            try
            {
                DevExpress.XtraTabbedMdi.XtraMdiTabPage currentPage = null;
                if (currentPage == null)
                {
                    loading.ShowWaitForm();
                    frmBase.Dock = DockStyle.Fill;
                    frmBase.FormClosing += new FormClosingEventHandler(frm_FormClosing);
                    frmBase.Text = strTitle;
                    frmBase.Tag = strTag;
                    if (!IsTabbedMdi)
                    frmBase.ClientSize = new Size(this.Width, this.Height);
                    frmBase.MdiParent = this;
                    if (image != null)
                    {
                        frmBase.ShowIcon = true;
                    }
                    else
                    image = base.Icon.ToBitmap();
                    frmBase.Show();
                    currentPage = xtraTabbedMdiMgr.SelectedPage;
                    xtraTabbedMdiMgr.SelectedPage.Image = image;
                }
                else
                {
                    xtraTabbedMdiMgr.SelectedPage = currentPage;
                }
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
            finally
            {
                loading.CloseWaitForm();
                LayoutMdi();
            }
        }

        void InitTabbedMDI()
        {
            xtraTabbedMdiMgr.MdiParent = IsTabbedMdi ? this : null;
            barCascade.Enabled = barHorizontal.Enabled = barVertical.Enabled = IsTabbedMdi ? false : true;
        }

        public bool ShowOpenedPage(string fName)
        {
            foreach (DevExpress.XtraTabbedMdi.XtraMdiTabPage page in xtraTabbedMdiMgr.Pages)
            {
                if (page.Text.Trim().ToUpper() == fName.Trim().ToUpper())
                {
                    xtraTabbedMdiMgr.SelectedPage = page;
                    return true;
                }
            }
            return false;
        }

        private string GetIpAddress()
        {
            string hostName = Dns.GetHostName();   //获取本机名
            IPHostEntry localhost = Dns.GetHostByName(hostName);    //方法已过期，可以获取IPv4的地址
            //IPHostEntry localhost = Dns.GetHostEntry(hostName);   //获取IPv6地址
            IPAddress localaddr = localhost.AddressList[0];

            return localaddr.ToString();
        }
        private void Form_QMS_Load(object sender, EventArgs e)
        {
            barStaticItem1.Caption = "海能达通信股份有限公司";
            barStaticItem2.Caption = "用户名:"+Login.username;  
            barStaticItem3.Caption = "用户工号:" + Login.userId;
            string ipaddr = "";
            ipaddr = GetIpAddress();
            barStaticItem4.Caption = "IP地址:"+ipaddr;

            if (Login.manager != "IT管理员" && Login.manager != "")
                initialmenu(Login.manager);             
            if (Login.manager != "IT管理员" && Login.manager == "")
                initialmenu(Login.post);
        }
        public void ShowForm(Type typeForm, Form formMdiParent)
        {
            Form form = Activator.CreateInstance(typeForm) as Form;
            form.MdiParent = formMdiParent;
            form.StartPosition = FormStartPosition.CenterScreen;
            form.Show();
        }



        bool permissions(string menuname,string post)
        {
            bool ifpermissions = false;
            string permissions = "";
                 
            string sql = @" select permissions from QMS_PermissionMenu  where menuname='" + menuname+ "' and post = '" +post+ "'";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            if (dt != null && dt.Rows.Count > 0)
            {
                permissions = dt.Rows[0]["permissions"].ToString();
                if (permissions == "是")
                {
                    ifpermissions = true;
                }
            }
            return ifpermissions;
        }


        bool supermanagerpermission()
        {
            bool ifpermission = false;
            string manager = "";
            string sql = @"select manager from QMS_UserInfo where userId = '" + Login.userId + "'";
            DataTable dts = DbAccess.SelectBySql(sql).Tables[0];
            if (dts != null && dts.Rows.Count > 0)
            {
                manager = dts.Rows[0]["manager"].ToString();
                if (manager == "IT管理员")
                {
                    ifpermission = true;
                    return ifpermission;
                }
            }
            return ifpermission;
        }

        bool managerpermission()
        {
            bool ifpermission = false;
            string manager = "";
            string sql = @"select manager from QMS_UserInfo where userId = '" + Login.userId + "'";
            DataTable dts = DbAccess.SelectBySql(sql).Tables[0];
            if (dts != null && dts.Rows.Count > 0)
            {
                ifpermission = true;
                return ifpermission;               
            }
            return ifpermission;
        }




        public void searchAllmenu()
        {
            foreach (RibbonPage module in this.ribbonControl.Pages)
            {
                foreach (RibbonPageGroup moduleGroup in module.Groups)
                {

                    foreach (BarItemLink menu  in moduleGroup.ItemLinks)
                    {

                        BarItem baritem = menu.Item;
                         if (baritem is BarSubItem)
                        {

                            BarSubItem subitem = baritem as BarSubItem;
                            foreach (BarItemLink ssubitem in subitem.ItemLinks)
                            {
                                ic.QMS_MenuList("新增", module.Text.Trim(), moduleGroup.Text.Trim(), ssubitem.Caption.Trim(),"","","");
                            }


                        }
                         else  if (baritem is BarButtonItem)
                        {
                            ic.QMS_MenuList("新增",module.Text.Trim(),moduleGroup.Text.Trim(),menu.Caption.Trim(),"","","");
                        }
                    }
                }
            }
        }

        void initialmenu(string post)
        {
            barButtonItem_jianyanweihu.Enabled = permissions("通用设置", post);
            barButtonItem_chengxu.Enabled = permissions("程序", post);
            BtnRepeatTestPro.Enabled = barButtonItem_chengxu.Enabled;
            barButtonItem_gongcha.Enabled = permissions("公差", post);
            barButtonItem_biaozhun.Enabled = permissions("标准", post);
            barButtonItem_siyintu.Enabled = permissions("丝印图", post);
            barButtonItem_tupiaoyang.Enabled = permissions("图片样", post);
            barButtonItem_yangpinuanli.Enabled = permissions("样品管理", post);
            barButtonItem_ceshi.Enabled = permissions("测试", post);
            Btn_FTest.Enabled = barButtonItem_ceshi.Enabled;
            barButtonItem_ROSH.Enabled = permissions("RoHS", post);
            barButtonItem_yicheng.Enabled = permissions("异常", post);
            BtnItem_shixiao.Enabled = permissions("失效分析", post);
            BtnItem_IQCCheckAgain.Enabled = permissions("IQC改判", post);
            BtnCommonreturn.Enabled = permissions("退料检验", post);
            BtnMaterialLot.Enabled = permissions("来料批次", post);
            barButtonItem_jianyanjilu.Enabled = permissions("检验记录", post);
            barButtonItem_jianyanbaobiao.Enabled = permissions("检验报表", post);
            BtnItem_yichangmingxi.Enabled = permissions("异常明细", post);
            BtnItemExceptionshow.Enabled = permissions("异常报表", post);
            BtnItem_tuiliaojilu.Enabled = permissions("退料记录", post);
            BtnItem_gongyingshangziliao.Enabled = permissions("代理证维护", post);

           ///// BtnIt_dataguidang.Enabled = permissions("资料档库", post);

            BtnIPQCExceptionProgSet.Enabled = permissions("制程信息", post);
            BtnIPQCException.Enabled = permissions("制程异常", post);
            BtnItem_jiliangMSA.Enabled = permissions("计量型", post);
            BtonItem_jishuMSA.Enabled = permissions("计数型", post);
            BtnItem_ESDItemInfo.Enabled = permissions("项目维护", post);
            BtnESDTestLabel.Enabled = permissions("标签打印", post);
            BtnItem_ESDTestNew.Enabled = permissions("测试设置", post);
            BtnItem_ESDjianyangaipan.Enabled = permissions("测试记录", post);

            BtnItem_zuoyeshezhi.Enabled = permissions("作业设置", post);
           // BtnItem_zuoyeshezhi.Enabled = permissions("信息维护", post);

            BtnItem_chuhuopingshen.Enabled = permissions("出货评审", post);
            BtnItem_jianyangaipan.Enabled = permissions("检验改判", post);
            baBtn_QAPackCheck.Enabled = permissions("QA核对", post);
            //BtnItem_chuhuojianyan.Enabled = permissions("出货检验", post);

            BtnItem_chuhuojianyan.Enabled = permissions("通用", post);
            BtnItem_chuhuojianyan.Enabled = permissions("专用", post);
            BtnItem_jianyanjilu.Enabled = permissions("出检记录", post);
            Btn_MEShistory.Enabled = permissions("历史数据", post);

            BtnItem_jianyanbaobiao.Enabled = permissions("出检报表", post);
            Btn_OQCexception.Enabled = permissions("出检异常", post);
            BtnItem_OQCyichangdan.Enabled = permissions("出检红单", post);
            BtnItem_mokuaishezhi.Enabled = permissions("模块设计", post);
            BtnItem_chuhuobaogao.Enabled = permissions("出货报告", post);

            //BtnassignManagerdPower.Enabled = supermanagerpermission();
            //itemPostProSet.Enabled = supermanagerpermission();
            //BtnemployeePower.Enabled = managerpermission();
            BtnassignManagerdPower.Enabled = permissions("管理员权限", post);
            itemPostProSet.Enabled = permissions("岗位维护", post);
            BtnemployeePower.Enabled = permissions("员工权限", post);
            BtnItem_Userregister.Enabled = permissions("用户注册", post);
            BtnItem_passworkupdate.Enabled = permissions("密码修改", post);
        }
        private void barButtonItem_jianyanweihu_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            showform(sender, e, "通用设置", "", new tongyongjianyan(), e.Item.Glyph); //LargeGlyph
        }
        private void BtnItem_chuhuojianyan_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            showform(sender, e, "出货检验", "", new OQCTestList(), e.Item.Glyph);//OQCTestList, chuhuojianyan
        }

        private void barButtonItem_yangpinuanli_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            showform(sender, e, "样品管理", "", new TestSamplePosition(), e.Item.Glyph);
        }

        private void BtnCommonreturn_ItemClick(object sender, ItemClickEventArgs e)
        {
            showform(sender, e, "退料检验", "", new IQCTestReturn(), e.Item.Glyph);
        }

        private void barButtonItem_jianyanjilu_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
 
            showform(sender, e, "检验记录", "", new TestListSearch(), e.Item.Glyph);
        }

        private void BtnItem_ESDTestNew_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            showform(sender, e, "测试设置", "", new ESDProgSet(), e.Item.Glyph);
        }

        private void BtnItem_zuoyeshezhi_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            showform(sender, e, "作业设置", "", new OQCTestProgSet(), e.Item.Glyph);
        }

        private void Form_QMS_FormClosing(object sender, FormClosingEventArgs e)
        {
            foreach (Form frm in MdiChildren)
            {
                frm.Dispose();
                frm.Close();
            }
            if (Common.COMScanDrive.sPrtScan != null && Common.COMScanDrive.sPrtScan.IsOpen)
                Common.COMScanDrive.EndComRead();
            Application.Exit();
        }


        private void BtnItem_ESDItemInfo_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            showform(sender, e, "项目维护", "", new ESDItemInfo(), e.Item.Glyph);
        }

        private void BtnItem_ESDjianyangaipan_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            showform(sender, e, "测试记录", "", new ESDTestNew(), e.Item.Glyph);
        }

        private void BtnItem_jianyangaipan_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            showform(sender, e, "检验改判", "", new OQCCheckAgain(), e.Item.Glyph);
        }

        private void barButtonItem_ceshi_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            showform(sender, e, "测试", "", new TestListcs(), e.Item.Glyph);

        }

        private void barButtonItem_ROSH_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            showform(sender, e, "RoHS", "", new TestROHS(), e.Item.Glyph);
        }

        private void barButtonItem_yicheng_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {

            showform(sender, e, "异常", "", new IQCRevExcept(), e.Item.Glyph);
        }

        private void barButtonItem_chengxu_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            showform(sender, e, "程序", "", new TestProgSet(), e.Item.Glyph);
        }

        private void barButtonItem_gongcha_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            showform(sender, e, "公差", "", new TestValue(), e.Item.Glyph);
        }

        private void barButtonItem_biaozhun_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            showform(sender, e, "标准", "", new TestMaterialCheckBySQE(), e.Item.Glyph);
        }

        private void barButtonItem_jianyanbaobiao_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
           //// showform(sender, e, "IQC检验报表", "", new IQCTestShow(), e.Item.Glyph);
             showform(sender, e, "检验报表", "", new DX_QMS.IQCFilePosition.IQCReport(), e.Item.Glyph);
        }

        private void barButtonItem_siyintu_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            showform(sender, e, "丝印图", "", new IQCProdPic(), e.Item.Glyph);
        }

        private void barButtonItem_tupiaoyang_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            showform(sender, e, "图片样", "", new IQCPurchaseMulPcode(), e.Item.Glyph);
        }

        private void BtnItem_jianyanjilu_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            showform(sender, e, "出检记录", "", new OQCTestListSearch(), e.Item.Glyph);
        }

        private void BtnItem_yichangmingxi_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            showform(sender, e, "异常明细", "", new ExceptionDetails(), e.Item.Glyph);
        }

        private void BtnItem_tuiliaojilu_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            showform(sender, e, "退料记录", "", new ReturnSearsh(), e.Item.Glyph);
        }

        private void baBtn_QAPackCheck_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            showform(sender, e, "QA核对", "", new QAPackCheck(), e.Item.Glyph);

        }

        private void BtnItem_Menu_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            ///// showform(sender, e, "菜单信息", "", new QMSMenu(), e.Item.Glyph);
           searchAllmenu();
           showform(sender, e, "管理员权限", "", new QMSassignManagerPower(), e.Item.Glyph);
        }

        private void BtnItem_permissions_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            showform(sender, e, "员工权限", "", new QMSEmployeePower(), e.Item.Glyph);
        }

        private void BtnItem_Userregister_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            showform(sender, e, "用户注册", "", new QMSRegister(), e.Item.Glyph);
        }

        private void BtnItem_passworkupdate_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            showform(sender, e, "密码修改", "", new SetPassword(), e.Item.Glyph);
        }

        private void BtnItem_jianyanbaobiao_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            showform(sender, e, "出检报表", "", new OQCTestShow(), e.Item.Glyph);
        }

        private void itemPostProSet_ItemClick(object sender, ItemClickEventArgs e)
        {
            showform(sender, e, "岗位维护", "", new PostMaintain(), e.Item.Glyph);

        }

        private void BtnItem_shixiao_ItemClick(object sender, ItemClickEventArgs e)
        {
            showform(sender, e, "失效分析", "", new FailureAnalysisDataInput(), e.Item.Glyph);
        }

        private void BtnItem_gongyingshangziliao_ItemClick(object sender, ItemClickEventArgs e)
        {
            showform(sender, e, "代理证维护", "", new overtimecheck(), e.Item.Glyph);
        }

        private void BtnItem_IQCCheckAgain_ItemClick(object sender, ItemClickEventArgs e)
        {
            showform(sender, e, "IQC改判", "", new TestIQCCheckAgain(), e.Item.Glyph);

        }

        private void Btn_MEShistory_ItemClick(object sender, ItemClickEventArgs e)
        {
            showform(sender, e, "历史数据", "", new OQCMESHistory(), e.Item.Glyph);
        }

        private void BtnOQCgeneraltest_ItemClick(object sender, ItemClickEventArgs e)
        {
            showform(sender, e, "通用", "", new OQCTestList(), e.Item.Glyph);
        }

        private void BtnOQCspecialtest_ItemClick(object sender, ItemClickEventArgs e)
        {
            showform(sender, e, "专用", "", new OQCF1TestList(), e.Item.Glyph);
        }

        private void BtnItem_chuhuobaogao_ItemClick(object sender, ItemClickEventArgs e)
        {
            showform(sender, e, "出货报告", "", new OQCReport(), e.Item.Glyph);
        }

        private void BtnItemExceptionshow_ItemClick(object sender, ItemClickEventArgs e)
        {
            showform(sender, e, "异常报表", "", new IQCExceptTestShow(), e.Item.Glyph);
        }

        private void bBtnOQCTestProgSet_ItemClick(object sender, ItemClickEventArgs e)
        {
            showform(sender, e, "作业维护", "", new OQCTestProgSet(), e.Item.Glyph);
        }

        private void bBtnOQCInformationSet_ItemClick(object sender, ItemClickEventArgs e)
        {
            showform(sender, e, "信息维护", "", new OQCinformation(), e.Item.Glyph);
        }

        private void Btn_OQCexception_ItemClick(object sender, ItemClickEventArgs e)
        {
            showform(sender, e, "出检异常", "", new OQCExceptTestShow(), e.Item.Glyph);
        }

        private void BtnQualityItem_ItemClick(object sender, ItemClickEventArgs e)
        {
            System.Diagnostics.Process ie = new System.Diagnostics.Process();

            try
            {
                ie.StartInfo.FileName = "IEXPLORE.EXE";
                ie.StartInfo.Arguments = "http://webapp.hytera.com/weboa/App/ES/isoprocess.nsf/2f139d6067a84a934825779600147bb1/17c7bc63c5ffa2e2482582310023318b?OpenDocument&Highlight=0,%E8%B4%A8%E9%87%8F%E6%89%8B%E5%86%8C";
                ie.Start();
            }
            catch 
            {
                MessageBox.Show("无法打开链接","提醒",MessageBoxButtons.OK ,MessageBoxIcon.Information);
            }
        }

        private void BtnMaterialLot_ItemClick(object sender, ItemClickEventArgs e)
        {
            showform(sender, e, "来料批次", "", new MaterialLot(), e.Item.Glyph);
        }

        private void BtnESDTestLabel_ItemClick(object sender, ItemClickEventArgs e)
        {
            showform(sender, e, "标签打印", "", new ESDTestLabel(), e.Item.Glyph);
        }

        private void BtnItemEMSReceive_ItemClick(object sender, ItemClickEventArgs e)
        {
            showform(sender, e, "客供报表", "", new EMSOtherReceiveRpt(), e.Item.Glyph);
        }

        private void BtnItem_jianyi_ItemClick(object sender, ItemClickEventArgs e)
        {
            showform(sender, e, "建议或反馈", "", new QMS_FeedBack(), e.Item.Glyph);
        }

        private void BtnItem_KPIindicator_ItemClick(object sender, ItemClickEventArgs e)
        {
            ///// showform(sender, e, "KPI指标", "", new DataSummaries(), e.Item.Glyph);
        }

        private void BtnItemKPI_ItemClick(object sender, ItemClickEventArgs e)
        {
            showform(sender, e, "KPI", "", new KPIitem(), e.Item.Glyph);
        }

        private void BtnItem_OQCyichangdan_ItemClick(object sender, ItemClickEventArgs e)
        {
           showform(sender, e, "出检红单", "", new QARedOrder(), e.Item.Glyph);
        }

        private void BtnIPQCException_ItemClick(object sender, ItemClickEventArgs e)
        {
            showform(sender, e, "制程异常","", new IPQCExceptionList(), e.Item.Glyph);
        }

        private void BtnIPQCExceptionProgSet_ItemClick(object sender, ItemClickEventArgs e)
        {
            showform(sender, e, "制程信息", "", new IPQCExceptionProgSet(), e.Item.Glyph);
        }

        private void BtnItemIPQCExcepReport_ItemClick(object sender, ItemClickEventArgs e)
        {
            showform(sender, e, "异常报表", "", new IPQCExceptionReportIPQC(), e.Item.Glyph);

        }

        private void BtnSMTFrontReport_ItemClick(object sender, ItemClickEventArgs e)
        {
            showform(sender, e, "SMT良率", "", new SMTFrontReport(), e.Item.Glyph);
        }

        private void BtnIt_dataguidang_ItemClick(object sender, ItemClickEventArgs e)
        {
            showform(sender, e, "资料档库", "", new IQCUploadFile(), e.Item.Glyph);
        }

        private void BtnMouldLife_ItemClick(object sender, ItemClickEventArgs e)
        {
            showform(sender, e, "寿命周期", "", new IQC_MouldLife(), e.Item.Glyph);
        }

        private void Btn_FTest_ItemClick(object sender, ItemClickEventArgs e)
        {
          showform(sender, e, "辅料检验", "", new IQCTestListcs(), e.Item.Glyph);
        }

        private void BtnC1report_ItemClick(object sender, ItemClickEventArgs e)
        {
            showform(sender, e, "测试良率", "", new SMTC1Report(), e.Item.Glyph);
        }

        private void Btn_SupperFile_ItemClick(object sender, ItemClickEventArgs e)
        {
             showform(sender, e, "环保资料", "", new IQC_SupperFile(), e.Item.Glyph);
        }

        private void BtnTestExpDate_ItemClick(object sender, ItemClickEventArgs e)
        {
            showform(sender, e, "超期重检", "", new IQC_MaterialExpiryDate(), e.Item.Glyph);
        }

        private void BtnRepeatTestPro_ItemClick(object sender, ItemClickEventArgs e)
        {
            showform(sender, e, "重检程序", "", new IQC_RepeatTest(), e.Item.Glyph);
        }


        //////  showform(sender, e, "作业设置", "", new OQCTestProgSet(), e.Item.Glyph);
    }
}
