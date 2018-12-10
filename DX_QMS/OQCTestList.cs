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
using System.Collections;
using DX_QMS.Common;
using System.Data;
using System.Data.SqlClient;
using System.Collections;

namespace DX_QMS
{
    public partial class OQCTestList : DevExpress.XtraBars.Ribbon.RibbonForm
    {

        string ifflat = "否";
        public OQCTestList()
        {
            InitializeComponent();
        }
        private void bindline()
        {
            string sql = "select Definetype,Definevalue from OQC_TypeDefine where Definetype='线体' order by sort ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            txtline.Properties.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
                txtline.Properties.Items.Add(row["Definevalue"]);
            }
            txtline.SelectedIndex = 0;
        }
        private void bindmaster()
        {
            string sql = "select Definetype,Definevalue from OQC_TypeDefine where Definetype='责任主管' order by sort ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            txtmaster.Properties.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
                txtmaster.Properties.Items.Add(row["Definevalue"]);
            }
            txtmaster.SelectedIndex = 0;
        }
        private void bindQE()
        {
            string sql = "select Definetype,Definevalue from OQC_TypeDefine where Definetype='责任QE' order by sort ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            txtQE.Properties.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
                txtQE.Properties.Items.Add(row["Definevalue"]);
            }
            txtQE.SelectedIndex = 0;
        }

        private void bindstate()
        {
            string sql = "select Definetype,Definevalue from OQC_TypeDefine where Definetype='产品状态' order by sort ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            txtstate.Properties.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
                txtstate.Properties.Items.Add(row["Definevalue"]);
            }
            txtstate.SelectedIndex = 0;
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

        void bindbadclass()
        {
            string sql = @"  select distinct badclass from OQC_baditem  ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            txtbadclass.Properties.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
                txtbadclass.Properties.Items.Add(row["badclass"]);
            }

        }
        private void txtbadclass_SelectedIndexChanged(object sender, EventArgs e)
        {
            string sql = @"  select badphenomenon,defects from OQC_baditem where badclass= '"+ txtbadclass.Text+ "'";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            txtbaddescribe.Properties.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
                txtbaddescribe.Properties.Items.Add(row["badphenomenon"]);
            }
            txtbaddescribe.SelectedIndex = 0;

            if (txtbaddescribe.Text == "需手动输入")
            {
                txtbaddescribe.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard;
                txtbaddescribe.Focus();
            }
            else
            {
                txtbaddescribe.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor;
            }
        }

        private void OQCTestList_Load(object sender, EventArgs e)
        {
            bindline();
            bindmaster();
            bindQE();
            bindstate();
            bindorg_id();
            bindbadclass();
            txtfactsampleqty.ReadOnly = true;
            txtNGnumber.ReadOnly = true;
            txtMAqty.Text = "0";
            txtMIqty.Text = "0";
            txtNGnumber.Text = "0";
            txtsendqty.ReadOnly = true;
            string ifflat = "否";
            // txtsendqty.Text = "0";
            //  JudgeTestItem(txtcustomer.Text.Trim(), txtPN.Text.Trim());
            // txtChecktype.SelectedIndex = 0;    
        }

        private bool JudgeWorknoEverTest(string workno, string boxsno)
        {
            ArrayList List = new ArrayList();
            string sql = @"select CartonNo from OQC_TestListNew where workno = '" + workno + "' and testresult = 'OK' ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            if (dt == null && dt.Rows.Count < 0)
            {
                return false;
            }
            else
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string Boxnumber = dt.Rows[i]["CartonNo"].ToString();
                    string[] values = Boxnumber.Split(',');
                    for (int n = 0; n < values.Length ; n++ )
                    {
                        if (values[n] == boxsno)
                        {
                            return true;
                        }
                    }
                }
                            return false;
            }
        }

        private bool JudgeWorknoOrRepete(string myworkno, ArrayList List)
        {
            string[] values = (string[])List.ToArray(typeof(string));
            for (int i = 0; i < values.Length; i++)
            {
                if (values[i] == myworkno)
                {
                    return true;
                }
            }
            return false;
        }


        ArrayList cartonlist = new ArrayList();
        private static int TestNum = 0;
        private static int CartonNum = 0;
        private void txtboxsno_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13 && txtboxsno.Text != "" )
            {
                if (txtboxnumber.Text == "")
                {
                    MessageBox.Show("请输入箱数！", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                string sqlstr = "select workno from";
                sqlstr += "(select org_id, workno, bigCartonNo AS CartonNo from OEM_FinisarRelation where org_id='" + txtorg_id.Text + "' and bigCartonNo='" + txtboxsno.Text + "' ";
                sqlstr += "union all select orgid AS org_id, workno, CartonNo from OEM_PackingCheck where CartonNo='" + txtboxsno.Text + "' ";
                sqlstr += "union all select 'SHL' AS org_id, workno, CartonNo from OEM_JSRSnRelation where CartonNo='" + txtboxsno.Text + "') AS a ";
                DataSet ds = DbAccess.SelectBySql(sqlstr);
                if (ds.Tables[0].Rows.Count <= 0)
                {                  
                    sqlstr = "select org_id, workno, lotno CartonNo,qty from OEM_EMSLotNo where lotno='" + txtboxsno.Text + "' and org_id='" + txtorg_id.Text + "'";
                    DataTable dt = DbAccess.SelectBySql(sqlstr).Tables[0];
                    if (dt == null || dt.Rows.Count <= 0)
                    {
                        
                       string zhusql = @"  select org_id ,workno,CartonNo,qty from  ";
                       zhusql = zhusql + "( select min(org_id) org_id ,min(workno) workno,min(BoxNo) CartonNo,count(SN ) qty  from SMT_TestPackCheck  where BoxNo = '" + txtboxsno.Text.Trim() + "' and org_id = '" + txtorg_id.Text + "'  ";
                       zhusql = zhusql + " union ";
                       zhusql = zhusql + "  select  org_id ,workno,flotno CartonNo,qty from SMT_Serialbuild  where flotno = '"+ txtboxsno.Text.Trim()+ "'  and org_id= '"+ txtorg_id.Text + "' ) t  where workno is not null and  CartonNo is not null  and qty is not null  ";

                        dt = DbAccess.SelectBySql(zhusql).Tables[0];

                        if (dt == null || dt.Rows.Count < 1)
                        {

                            string x1sql = @"  select org_id ,workno ,lotno CartonNo, qty   from SMT_SONGBAN  where lotno = '"+ txtboxsno.Text.Trim() + "' and org_id = '"+ txtorg_id.Text + "' ";
                            dt = DbAccess.SelectBySql(x1sql).Tables[0];
                            if (dt == null || dt.Rows.Count < 1)
                            {
                                MessageBox.Show("输入的箱号：" + txtboxsno.Text + " 错误");
                                txtboxsno.Text = "";
                                txtboxsno.Focus();
                                return;
                            }
                        }

                    }            
                        string workno = dt.Rows[0]["workno"].ToString();
                        if (JudgeWorknoEverTest(workno, txtboxsno.Text.Trim()))
                        {
                            MessageBox.Show("箱号：" + txtboxsno.Text + " 已检查");
                            txtboxsno.Text = "";
                            return;
                        }
                        if (JudgeWorknoOrRepete(txtboxsno.Text, cartonlist))
                        {
                            MessageBox.Show("箱号：" + txtboxsno.Text + " 重复输入");
                            txtboxsno.Text = "";
                            return;
                        }
                        else
                        {
                            string[] testvalues = (string[])cartonlist.ToArray(typeof(string));
                            if (testvalues.Length == 0)
                            {
                                txtworkno.Text = dt.Rows[0]["workno"].ToString();
                                txtworkno_Leave(sender, e);
                                cartonlist.Add(txtboxsno.Text);
                                TestNum += int.Parse(dt.Rows[0]["qty"].ToString());
                                CartonNum++;
                            }
                            else
                            {                                
                                    cartonlist.Add(txtboxsno.Text);
                                    TestNum += int.Parse(dt.Rows[0]["qty"].ToString());
                                    CartonNum++;                                
                            }

                            string cartonstr = "";
                            string[] values = (string[])cartonlist.ToArray(typeof(string));
                            for (int i = 0; i < values.Length - 1; i++)
                            {
                                cartonstr += values[i] + "；";
                            }
                                cartonstr += values[values.Length - 1];
                            if (CartonNum == Convert.ToInt32(txtboxnumber.Text))
                            {
                                txtsendqty.Text = Convert.ToString(TestNum);

                                txtsendqty_Leave(sender, e);
                                txtSerialnumber.Text =  ProductSerialnumber(workno);

                                lblinfo.ForeColor = Color.Green;
                                lblinfo.Text = "箱号 " + cartonstr + " 全部输入完成";
                                txtboxsno.ReadOnly = true;

                               // txtcustomer.Focus();   /////////////////////////////

                                TestNum = 0;
                                CartonNum = 0;
                            }
                            else
                            {
                                lblinfo.ForeColor = Color.Green;
                                lblinfo.Text = "箱号 " + cartonstr + " 已录入。已录入箱数为：" + CartonNum.ToString();
                                txtboxsno.Text = "";
                            }
                        }
                   
                 }
                else
                {
                    string workno = ds.Tables[0].Rows[0]["workno"].ToString();
                    if (JudgeWorknoEverTest(workno, txtboxsno.Text.Trim()))
                    {
                        MessageBox.Show("箱号：" + txtboxsno.Text + " 已检查");
                        txtboxsno.Text = "";
                        return;
                    }
                    if (JudgeWorknoOrRepete(txtboxsno.Text, cartonlist))
                    {
                        MessageBox.Show("箱号：" + txtboxsno.Text + " 重复输入");
                        txtboxsno.Text = "";
                        return;
                    }
                    else
                    {
                        string[] testvalues = (string[])cartonlist.ToArray(typeof(string));
                        if (testvalues.Length == 0)
                        {
                            txtworkno.Text = ds.Tables[0].Rows[0]["workno"].ToString();
                            txtworkno_Leave(sender, e);
                            cartonlist.Add(txtboxsno.Text);
                            TestNum += ds.Tables[0].Rows.Count;
                            CartonNum++;
                        }
                        else
                        {                         
                                cartonlist.Add(txtboxsno.Text);
                                TestNum += ds.Tables[0].Rows.Count;
                                CartonNum++;                            
                        }
                        string cartonstr = "";
                        string[] values = (string[])cartonlist.ToArray(typeof(string));
                        for (int i = 0; i < values.Length - 1; i++)
                        {
                            cartonstr += values[i] + "；";
                        }
                        cartonstr += values[values.Length - 1];
                        if (CartonNum == Convert.ToInt32(txtboxnumber.Text))
                        {
                            txtsendqty.Text = Convert.ToString(TestNum);

                            txtsendqty_Leave(sender, e);
                            txtSerialnumber.Text = ProductSerialnumber(workno);


                            lblinfo.ForeColor = Color.Green;
                            lblinfo.Text = "箱号 " + cartonstr + " 全部输入完成";
                            txtboxsno.ReadOnly = true;
                           // txtcustomer.Focus();  //////////////////////////////////////
                            TestNum = 0;
                            CartonNum = 0;
                        }
                        else
                        {
                            lblinfo.ForeColor = Color.Green;
                            lblinfo.Text = "箱号 " + cartonstr + " 已录入。已录入箱数为：" + CartonNum.ToString();
                            txtboxsno.Text = "";   
                        }
                    }
                 }
               }
            }

         string  MAvalue="", MIvalue="";            
        private void JudgeTestItem(string customer, string PN)
        {
            string sqll = "";
            string sql = @" select testitem 测试项目,checkmethod 检查方法,standardsequence 序号,teststandard 检验标准,checkMA MA主要缺陷,checkMI MI次要缺陷,'' 检验项目结果 from OQCTestProgSet where customer ='" + customer + "' and PN='" + PN + "'";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            if (dt != null && dt.Rows.Count > 0)
            {
                testcontent.DataSource = dt;
                sqll = @" select sampleplan,MA,MAvalue,MI,MIvalue,SNstates from OQCTestProgSet where customer ='" + customer + "' and PN='" + PN + "'";
                DataTable dt1 = DbAccess.SelectBySql(sqll).Tables[0];
                txtsamleplan.Text = dt1.Rows[0]["sampleplan"].ToString();
                txtMA.Text = dt1.Rows[0]["MA"].ToString();
                txtMI.Text = dt1.Rows[0]["MI"].ToString();  /////////////////////////////////
                MAvalue = dt1.Rows[0]["MAvalue"].ToString();
                MIvalue = dt1.Rows[0]["MIvalue"].ToString();
                string SNstates = dt1.Rows[0]["SNstates"].ToString();
                if (SNstates == "是")
                {
                    lblSNstates.Text = "需要扫描SN号";
                    txtfactsampleqty.ReadOnly = true;
                    txtNGnumber.ReadOnly = true;
                }
                else
                {
                    lblSNstates.Text = "不需要扫描SN号";
                    txtfactsampleqty.ReadOnly = false ;
                    txtNGnumber.ReadOnly = false ;
                }
            }
            else
            {
                sql = @" select testitem 测试项目,checkmethod 检查方法,standardsequence 序号,teststandard 检验标准,checkMA MA主要缺陷,checkMI MI次要缺陷,'' 检验项目结果 from OQCTestProgSet where customer ='" + customer + "' and PN='' ";
                DataTable dts = DbAccess.SelectBySql(sql).Tables[0];
                if (dts != null && dts.Rows.Count > 0)
                {
                    testcontent.DataSource = dts;
                    sqll = @" select sampleplan,MA,MAvalue,MI,MIvalue,SNstates from OQCTestProgSet where customer ='" + customer + "' and PN='' ";
                    DataTable dts1 = DbAccess.SelectBySql(sqll).Tables[0];
                    txtsamleplan.Text = dts1.Rows[0]["sampleplan"].ToString();
                    txtMA.Text = dts1.Rows[0]["MA"].ToString();
                    txtMI.Text = dts1.Rows[0]["MI"].ToString();  /////////////////////////////////
                    MAvalue = dts1.Rows[0]["MAvalue"].ToString();
                    MIvalue = dts1.Rows[0]["MIvalue"].ToString();
                    string SNstates = dts1.Rows[0]["SNstates"].ToString();
                    if (SNstates == "是")
                    {
                        lblSNstates.Text = "需要扫描SN号";
                        txtfactsampleqty.ReadOnly = true;
                        txtNGnumber.ReadOnly = true;
                    }
                    else
                    {
                        lblSNstates.Text = "不需要扫描SN号";
                        txtfactsampleqty.ReadOnly = false;
                        txtNGnumber.ReadOnly = false;
                    }
                }
                else
                {
                    sql = @" select testitem 测试项目,checkmethod 检查方法,standardsequence 序号,teststandard 检验标准,checkMA MA主要缺陷,checkMI MI次要缺陷 ,'' 检验项目结果 from OQCTestProgSet where customer ='' and PN='' ";
                    DataTable dss = DbAccess.SelectBySql(sql).Tables[0];
                    if (dss != null && dss.Rows.Count > 0)
                    {
                        testcontent.DataSource = dss;
                        sqll = @" select sampleplan,MA,MAvalue,MI,MIvalue,SNstates from OQCTestProgSet where customer ='' and PN='' ";
                        DataTable dss1 = DbAccess.SelectBySql(sqll).Tables[0];
                        txtsamleplan.Text = dss1.Rows[0]["sampleplan"].ToString();
                        txtMA.Text = dss1.Rows[0]["MA"].ToString();
                        txtMI.Text = dss1.Rows[0]["MI"].ToString();  /////////////////////////////////
                        MAvalue = dss1.Rows[0]["MAvalue"].ToString();
                        MIvalue = dss1.Rows[0]["MIvalue"].ToString();
                        string SNstates = dss1.Rows[0]["SNstates"].ToString();
                        if (SNstates == "是")
                        {
                            lblSNstates.Text = "需要扫描SN号";
                            txtfactsampleqty.ReadOnly = true;
                            txtNGnumber.ReadOnly = true;
                        }
                        else
                        {
                            lblSNstates.Text = "不需要扫描SN号";
                            txtfactsampleqty.ReadOnly = false;
                            txtNGnumber.ReadOnly = false;
                        }
                    }
                    else
                    {
                        MessageBox.Show("没有维护检验项目","提醒",MessageBoxButtons.OK,MessageBoxIcon.Information);
                        return;
                    }
                }
            }

        }

        private string ProductSerialnumber(string workno)
        {
            string Serialnumber = "";
            string sql = " select SubString( replace(replace(replace(replace(convert(varchar(30),getdate(),121),'-',''),' ',''),':',''),'.',''),3,15) ";
            Serialnumber = DbAccess.SelectBySql(sql).Tables[0].Rows[0][0].ToString();
            Serialnumber = workno + Serialnumber;
            return Serialnumber;
        }


        private string checktype( string PN )
        {    
            string checktype ="正常检验";
            int NGConut = 0;
            string sql = @" select checktype,testresult from OQC_TestListNew where hytcode ='" + PN + "' order by checkdate desc ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            if (dt == null && dt.Rows.Count < 0)
            {
                checktype = "正常检验";
            } 
             checktype = dt.Rows[0]["checktype"].ToString();
            if (checktype == "正常检验")
            {
                for (int i =0; i<5;i++ )
                {
                    //NGConut = NGConut+int.Parse(dt.Rows[i]["testresult"].ToString());
                    string testresult = dt.Rows[i]["testresult"].ToString();
                    if (testresult == "NG")
                        NGConut = NGConut + 1;
                }
                if (NGConut < 2)
                {
                    checktype = "正常检验";
                }
                else               
                    checktype = "加严检验";               
            }
            else
                checktype = "加严检验";
            return checktype;
        }


        string MaterialState(string Materialcode)
        {
            string  sql, State = "正常检验";
            sql = @" select States from OQC_TestMaterialState where Materialcode ='"+ Materialcode + "'";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            if (dt!=null && dt.Rows.Count >0)
            {
                State = dt.Rows[0]["States"].ToString();
            }
            return State;
        }

        private void txtworkno_Leave(object sender, EventArgs e)
        {
            gridControl.DataSource = null;
            if (txtworkno.Text == "") return;
            txtworkno.Leave -= txtworkno_Leave;
            //string sql = "select item_num,item_desc,start_quantity,job_desc,WO.DEPARTMENT_ID,completion_subinventory,order_number,gn_mo,V.ORGANIZATION_ID from cux_wip_header_v v,CUX_WIP_OPERATIONS_V WO  where " +
            //          " WO.WIP_ENTITY_ID =v.WIP_ENTITY_ID and wip_entity_name =" + "'" + txtworkno.Text.ToUpper() + "' and v.ORGANIZATION_CODE = '" + txtorg_id + "'";

            string sql = "select item_num,item_desc,start_quantity,job_desc,WO.DEPARTMENT_ID,completion_subinventory,order_number,gn_mo,V.ORGANIZATION_ID from cux_wip_header_v v,CUX_WIP_OPERATIONS_V WO  where " +
                                " WO.WIP_ENTITY_ID =v.WIP_ENTITY_ID and wip_entity_name =" + "'" + txtworkno.Text.Trim().ToUpper() + "' and v.ORGANIZATION_CODE = '" + txtorg_id.Text + "'";
            DataSet ds = DbAccess.SelectByOracle(sql);
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                lblinfo.Text = ds.Tables[0].Rows[0]["item_desc"].ToString();
                lblinfo.ForeColor = Color.Blue;
                txtPN.Text = ds.Tables[0].Rows[0]["item_num"].ToString();
                txtdelivery.Text = ds.Tables[0].Rows[0]["job_desc"].ToString();
                txtworknumber.Text = ds.Tables[0].Rows[0]["start_quantity"].ToString();
                if (txtworkno.Text.ToUpper().EndsWith("WS") || txtworkno.Text.ToUpper().EndsWith("AS") || txtworkno.Text.ToUpper().EndsWith("WN") || txtworkno.Text.ToUpper().EndsWith("AN"))
                {
                    string s = "";
                    string ssql = "select Definevalue,sort from OQC_TypeDefine where Definetype='检验项目' order by sort ";
                    DataTable dt = DbAccess.SelectBySql(ssql).Tables[0];
                    if (dt.Rows.Count > 0)
                    {
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            s += dt.Rows[i]["sort"].ToString() + "." + dt.Rows[i]["Definevalue"].ToString() + " ";
                        }
                    }
                    DialogResult result = MessageBox.Show(s, "请确认是否检验完下面的项目!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (DialogResult.No == result)
                    {
                        txtworkno.Text = "";
                        txtworkno.Focus();
                        txtworkno.Leave += txtworkno_Leave;
                        return;
                    }
                }

                ArrayList ltL = new ArrayList();
                ArrayList customerL = new ArrayList();
                int index = 0, customerindex = 0;
                string str = ds.Tables[0].Rows[0]["item_desc"].ToString();
                foreach (Char ch in str)
                {
                    if (ch == '('|| ch == '（')
                    {
                        ltL.Add(index);
                    }
                    if (ch == '【')
                    {
                        customerL.Add(customerindex);
                    }
                    index++;
                    customerindex++;
                }
                ArrayList ltR = new ArrayList();
                ArrayList customerR = new ArrayList();
                int indexR = 0, customerindexR = 0;
                foreach (Char ch in str)
                {
                    if (ch == ')' || ch == '）')
                    {
                        ltR.Add(indexR);
                    }
                    if (ch == '】')
                    {
                        customerR.Add(customerindexR);
                    }
                    indexR++;
                    customerindexR++;
                }
                try
                {
                    try
                    {
                        txtmodel.Text = str.Substring(int.Parse(ltL[0].ToString()) + 1, int.Parse(ltR[0].ToString()) - int.Parse(ltL[0].ToString()) - 1);
                    }
                    catch
                    {
                        txtmodel.Text = "";
                    }

                    txtcustomer.Text = str.Substring(int.Parse(customerL[0].ToString()) + 1, int.Parse(customerR[0].ToString()) - int.Parse(customerL[0].ToString()) - 1).Replace("专用", "");

                    if (txtcustomer.Text.Contains("申请"))
                    {
                        txtcustomer.Text = str.Substring(int.Parse(customerL[0].ToString()) + 1, int.Parse(customerR[0].ToString()) - int.Parse(customerL[0].ToString()) - 1).Replace("申请", "");
                    }
                    if (txtcustomer.Text == "保税")
                    {
                      txtcustomer.Text = str.Substring(int.Parse(customerL[1].ToString()) + 1, int.Parse(customerR[1].ToString()) - int.Parse(customerL[1].ToString()) - 1).Replace("专用", "");
                    }
                    if (txtcustomer.Text == "返工")
                    {
                        txtcustomer.Text = str.Substring(int.Parse(customerL[1].ToString()) + 1, int.Parse(customerR[1].ToString()) - int.Parse(customerL[1].ToString()) - 1).Replace("专用", "");
                    }
                    if (txtcustomer.Text == "16004")
                    {
                        txtcustomer.Text = "新飞通";
                    }

                    txtproductdate.Text = DateTime.Now.ToString("yyyy-MM-dd");                                   
                }
                catch
                {   
                    ///// 增加主营0309
                    if (txtPN.Text.Trim().StartsWith("1150") || txtPN.Text.Trim().StartsWith("6103") || txtPN.Text.Trim().StartsWith("1151"))
                    {
                        txtcustomer.Text = "主营";
                        txtmodel.Text = lblinfo.Text;
                        txtdelivery.Text = "";
                        txtproductdate.Text = DateTime.Now.ToString("yyyy-MM-dd");
                    }
                }

                if (txtcustomer.Text == "F1")
                {
                    txtsendqty.ReadOnly = false;
                }
                else
                    txtsendqty.ReadOnly = true;

                JudgeTestItem(txtcustomer.Text.Trim(), txtPN.Text.Trim());
                if (txtsamleplan.Text == "ISO2859-1")
                {
                   // txtChecktype.SelectedIndex = 0;
                }
                else if (txtsamleplan.Text == "全检")
                {
                    txtChecktype.Enabled = false;
                    txtfactsampleqty.ReadOnly = false;
                    txtNGnumber.ReadOnly = false;
                }

                txtChecktype.Text = MaterialState(txtPN.Text.Trim());
               ///////////////// txtsendqty_Leave(sender,e);

            }
            else
            {


                ////////txtPN.Text = "70010000000094 ";
                ////////txtworknumber.Text = "5000";
                ////////txtmodel.Text = "C170010000000094";
                ////////txtcustomer.Text = "C1";

                ////////JudgeTestItem(txtcustomer.Text.Trim(), txtPN.Text.Trim());
                ////////if (txtsamleplan.Text == "ISO2859-1")
                ////////{
                ////////    txtChecktype.SelectedIndex = 0;
                ////////}
                ////////else
                ////////{
                ////////    txtChecktype.Enabled = false;
                ////////}
                //////// txtsendqty_Leave(sender, e);

                lblinfo.Text = txtworkno.Text + "该工单号不存在";
                lblinfo.ForeColor = Color.Red;
                txtworkno.Text = "";
                txtworkno.Focus();
            }
            txtworkno.Leave += txtworkno_Leave;
        }

        private void txtboxnumber_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13 && txtboxnumber.Text != "")
            {
                // txtboxnumber_TextChanged(sender, e);
                int boxqty = 0;
                if (!int.TryParse(txtboxnumber.Text, out boxqty))
                {
                    MessageBox.Show("箱数不正确", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txtboxnumber.Text = "";
                    return;
                }
                if (boxqty < 0 || boxqty > 50)
                {
                    MessageBox.Show("箱数不能小于0大于50", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txtboxnumber.Text = "";
                    return;
                }
                txtworkno.ReadOnly = true;
                txtboxnumber.ReadOnly = true;
                txtboxsno.ReadOnly = false;
                txtboxsno.Focus();
            }
        }

        private void txtboxnumber_Leave(object sender, EventArgs e)
        {
            txtboxnumber_TextChanged(sender, e);
            txtboxsno.ReadOnly = false;
            txtboxsno.Focus();
        }

        private void txtsendqty_Leave(object sender, EventArgs e)
        {
            int sendqty = 0;
            if (!int.TryParse(txtsendqty.Text, out sendqty))
            {
               // MessageBox.Show("送检数量不正确", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
               // lalinformation.Text = "送检数量不正确";
                return;
            }
           // lalinformation.Text = "";
            double douMAvalue = 0, douMIvalue = 0;
            string sql = "", checktype = "", Code = "";
            string sqlMA = "",sqlMI="";
            if (!double.TryParse(MAvalue, out douMAvalue))
            {
                MessageBox.Show("MA数值不正确", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (!double.TryParse(MIvalue, out douMIvalue))
            {
                MessageBox.Show("MI数值不正确", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (txtsamleplan.Text == "ISO2859-1")
            {
                checktype = txtChecktype.Text;
                if (checktype == "正常检验")
                {

                    if (sendqty > 1)
                    {
                        sql = @"Select case when Sampleqty>lotfactqty then lotfactqty else Sampleqty end as Sampleqty,s.Code from IQC_TestSTD105ECode c ";
                        sql += "  inner join ";
                        sql += " (Select " + sendqty + " lotfactqty,AQLValue,AQL,CheckLevel,Ac,Re,s.Code from IQC_TestSTD105ECheckSet i inner join IHPS_QUALITY_SPC_AQLIS02859 s on i.Code=s.Code  where LotSizemin<=" + sendqty.ToString() + " and LotSizemax>=" + sendqty.ToString() + " and CheckLevel='II') s on c.Code=s.Code";
                        DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
                        txtsampleqty.Text = dt.Rows[0]["Sampleqty"].ToString();
                        Code = dt.Rows[0]["Code"].ToString();
                        if (douMAvalue != 0)
                        {
                            sqlMA = @"select Ac,Re from IHPS_QUALITY_SPC_AQLIS02859 where Type ='正常检验' and AQLValue = " + douMAvalue + " and Code = '" + Code + "'";
                        }
                        else
                        {
                            sqlMA = @"select Ac,Re from IHPS_QUALITY_SPC_AQLIS02859 where Type ='正常检验' and AQLValue = 1.0 and Code = '" + Code + "'";
                        }
                        if (douMIvalue != 0)
                        {
                            sqlMI = @"select Ac,Re from IHPS_QUALITY_SPC_AQLIS02859 where Type ='正常检验' and AQLValue = " + douMIvalue + " and Code = '" + Code + "'";
                        }
                        else
                        {
                            sqlMI = @"select Ac,Re from IHPS_QUALITY_SPC_AQLIS02859 where Type ='正常检验' and AQLValue = 1.0  and Code = '" + Code + "'";
                        }
                        DataTable dtz = DbAccess.SelectBySql(sqlMA).Tables[0];
                        DataTable dtzMI = DbAccess.SelectBySql(sqlMI).Tables[0];
                        txtMAallowqty.Text = dtz.Rows[0]["Ac"].ToString();
                        txtMIallowqty.Text = dtzMI.Rows[0]["Ac"].ToString();
                    }
                    else
                    {
                        txtsampleqty.Text = "1";
                        txtMAallowqty.Text = "1";
                        txtMIallowqty.Text = "1";
                    }

                }
                else if (checktype == "加严检验")
                {

                    if (sendqty > 1)
                    {
                        sql = @"Select case when Sampleqty>lotfactqty then lotfactqty else Sampleqty end as Sampleqty,s.Code from IQC_TestSTD105ECode c ";
                        sql += "  inner join ";
                        sql += " (Select " + sendqty + " lotfactqty,AQLValue,AQL,CheckLevel,Ac,Re,s.Code from IQC_TestSTD105ECheckSet i inner join IHPS_QUALITY_SPC_AQLIS02859 s on i.Code=s.Code  where LotSizemin<=" + sendqty.ToString() + " and LotSizemax>=" + sendqty.ToString() + " and CheckLevel='II') s on c.Code=s.Code";
                        DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
                        txtsampleqty.Text = dt.Rows[0]["Sampleqty"].ToString();
                        Code = dt.Rows[0]["Code"].ToString();

                        if (douMAvalue != 0)
                        {
                            sqlMA = @"select Ac,Re from IHPS_QUALITY_SPC_AQLIS02859 where Type ='加严' and AQLValue = " + douMAvalue + " and Code = '" + Code + "'";
                        }
                        else
                        {
                            sqlMA = @"select Ac,Re from IHPS_QUALITY_SPC_AQLIS02859 where Type ='加严' and AQLValue = 1.0 and Code = '" + Code + "'";
                        }
                        if (douMIvalue != 0)
                        {
                            sqlMI = @"select Ac,Re from IHPS_QUALITY_SPC_AQLIS02859 where Type ='加严' and AQLValue = " + douMIvalue + " and Code = '" + Code + "'";
                        }
                        else
                        {
                            sqlMI = @"select Ac,Re from IHPS_QUALITY_SPC_AQLIS02859 where Type ='加严' and AQLValue = 1.0  and Code = '" + Code + "'";
                        }
                        DataTable dtj = DbAccess.SelectBySql(sqlMA).Tables[0];
                        DataTable dtjMI = DbAccess.SelectBySql(sqlMI).Tables[0];
                        txtMAallowqty.Text = dtj.Rows[0]["Ac"].ToString();
                        txtMIallowqty.Text = dtjMI.Rows[0]["Ac"].ToString();

                    }
                    else
                    {
                        txtsampleqty.Text = "1";
                        txtMAallowqty.Text = "1";
                        txtMIallowqty.Text = "1";
                    }
                }
                else if (checktype == "放宽检验")
                {

                    if (sendqty > 1)
                    {
                        sql = @"Select distinct case when Sampleqty>lotfactqty then lotfactqty else Sampleqty end as Sampleqty,s.Code from IQC_TestSTD105ECode c ";
                        sql += "  inner join ";
                        sql += " (Select " + sendqty + " lotfactqty,AQLValue,AQL,CheckLevel,Ac,Re,s.Code from IQC_TestSTD105ECheckSet i inner join IHPS_QUALITY_SPC_AQLIS02859 s on i.Code1=s.Code  where LotSizemin<=" + sendqty.ToString() + " and LotSizemax>=" + sendqty.ToString() + " and CheckLevel='II') s on c.Code=s.Code";
                        DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
                        txtsampleqty.Text = dt.Rows[0]["Sampleqty"].ToString();
                        Code = dt.Rows[0]["Code"].ToString();
                        if (douMAvalue != 0)
                        {
                            sqlMA = @"select Ac,Re from IHPS_QUALITY_SPC_AQLIS02859 where Type ='放宽' and AQLValue = " + douMAvalue + " and Code = '" + Code + "'";
                        }
                        else
                        {
                            sqlMA = @"select Ac,Re from IHPS_QUALITY_SPC_AQLIS02859 where Type ='放宽' and AQLValue = 1.0 and Code = '" + Code + "'";
                        }
                        if (douMIvalue != 0)
                        {
                            sqlMI = @"select Ac,Re from IHPS_QUALITY_SPC_AQLIS02859 where Type ='放宽' and AQLValue = " + douMIvalue + " and Code = '" + Code + "'";
                        }
                        else
                        {
                            sqlMI = @"select Ac,Re from IHPS_QUALITY_SPC_AQLIS02859 where Type ='放宽' and AQLValue = 1.0  and Code = '" + Code + "'";
                        }
                        DataTable dtf = DbAccess.SelectBySql(sqlMA).Tables[0];
                        DataTable dtfMI = DbAccess.SelectBySql(sqlMI).Tables[0];
                        txtMAallowqty.Text = dtf.Rows[0]["Ac"].ToString();
                        txtMIallowqty.Text = dtfMI.Rows[0]["Ac"].ToString();
                    }
                    else
                    {
                        txtsampleqty.Text = "1";
                        txtMAallowqty.Text = "1";
                        txtMIallowqty.Text = "1";
                    }
                }
                else if (checktype == "放宽检验")
                {
                    MessageBox.Show("改物料暂停检验，请联系QE重新审核", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
            }
            else if (txtsamleplan.Text == "C=0")
            {

                if (sendqty > 1)
                {
                    double MAAQLValue = 1.0, MIAQLValue = 1.0;
                    if (sendqty >= 500000)
                    {
                        sendqty = 500000;
                    }
                    if (douMAvalue != 0)
                    {
                        MAAQLValue = douMAvalue;
                    }
                    if (douMIvalue != 0)
                    {
                        MIAQLValue = douMIvalue;
                    }
                    string cosampleMA = @"select case when sampleqty = '*' then '*' else sampleqty end as Sampleqty from IHPS_QUALITY_SPC_AQLC0 ";
                    cosampleMA += " where ( Lowervalue <=" + sendqty + "and Uppervalue >=" + sendqty.ToString() + "and AQLValue = " + MAAQLValue + ")";

                    string cosampleMI = @"select case when sampleqty = '*' then '*' else sampleqty end as Sampleqty from IHPS_QUALITY_SPC_AQLC0 ";
                    cosampleMI += " where ( Lowervalue <=" + sendqty + "and Uppervalue >=" + sendqty.ToString() + "and AQLValue = " + MIAQLValue + ")";

                    DataTable dttqtyMA = DbAccess.SelectBySql(cosampleMA).Tables[0];
                    DataTable dttqtyMI = DbAccess.SelectBySql(cosampleMI).Tables[0];

                    string SampleqtyMA = dttqtyMA.Rows[0]["Sampleqty"].ToString();
                    string SampleqtyMI = dttqtyMI.Rows[0]["Sampleqty"].ToString();

                    if (SampleqtyMA.Contains("*"))
                    {
                        txtsampleqty.Text = sendqty.ToString();
                    }
                    else
                    {
                        txtsampleqty.Text = SampleqtyMA;
                    }
                    txtMAallowqty.Text = "0";
                    txtMIallowqty.Text = "0";
                }
                else
                {
                    txtsampleqty.Text = "1";
                    txtMAallowqty.Text = "0";
                    txtMIallowqty.Text = "0";
                }

            }
            else
            {
                txtsampleqty.Text = txtsendqty.Text;
                txtMAallowqty.Text = "0";
                txtMIallowqty.Text = "0";
            }
           
        }

        private void txtChecktype_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtsendqty_Leave(sender, e);
        }

        private void gridtestcontent_Click(object sender, EventArgs e)
        {
            if (testcontent.DataSource == null)
            {
                MessageBox.Show("没有数据","提醒",MessageBoxButtons.OK,MessageBoxIcon.Warning);
                return;
            }

            //txttestitem.Text = gridtestcontent.GetFocusedRowCellValue("测试项目").ToString();
            //txttools.Text = gridtestcontent.GetFocusedRowCellValue("检查方法").ToString();

        }


        ArrayList SNnumberlist = new ArrayList();
        private static int factsampleqty = 0;
        private bool JudgeSNnumber(string SNnumber, ArrayList List)
        {
            string[] values = (string[])List.ToArray(typeof(string));
            for (int i = 0; i < values.Length; i++)
            {
                if (values[i] == SNnumber)
                {
                    return true;
                }
            }
            return false;
        }

        static string sNnumber = "";

        private void txtSNnumber_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13 && txtSNnumber.Text != "")
            {
                if (lblSNstates.Text == "不需要扫描SN号")
                {
                    MessageBox.Show("不需要扫描SN号", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    txtSNnumber.Text = "";
                    return;
                }
                //string sqlstr = "select 1 from SMT_PinLot s inner join lotnoinfo l on s.lotno=l.lotno where s.serialno='" + txtSNnumber.Text + "' and workno='" + txtworkno.Text + "'";
                //DataTable dtbar = DbAccess.SelectBySql(sqlstr).Tables[0];

                string zhusql = @"  select 1  from  SMT_TestPackCheck where SN = '" + txtSNnumber.Text + "' and workno= '" + txtworkno.Text + "' ";         
                DataTable zhudt = DbAccess.SelectBySql(zhusql).Tables[0];
                if (zhudt.Rows.Count <= 0 && this.txtcustomer.Text == "主营")
                {
                    lblinfo.ForeColor = Color.Red;
                    lblinfo.Text = txtSNnumber.Text + " 不是此工单的二维码或不存在!";
                    txtSNnumber.Text = "";
                    return;
                }
                else
                {
                    //DataTable IfRelation = DbAccess.SelectBySql("select checkdate,org_id, workno,serialno from OQC_SampleFactList where serialno='"
                    //                  + this.txtSNnumber.Text.Trim() + "'").Tables[0];
                    DataTable IfRelation = DbAccess.SelectBySql("select checkdate,org_id, workno,SNnumber from OQC_SampleNewList where SNnumber='"
                  + this.txtSNnumber.Text.Trim() + "'  and  workno = '"+txtworkno.Text+ "' ").Tables[0];

                    if (IfRelation.Rows.Count > 0)
                    {
                        if (IfRelation.Rows[0]["checkdate"].ToString() != "")
                        {
                            this.lblinfo.Text = "你输入的条码号:" + txtSNnumber.Text.ToUpper().Trim() + ",已装入工单:" + IfRelation.Rows[0]["workno"].ToString();
                            lblinfo.ForeColor = Color.Red;
                            txtSNnumber.Text = "";
                            txtSNnumber.Focus();
                            return;
                        }
                    }
                    if (JudgeSNnumber(txtSNnumber.Text, SNnumberlist))
                    {
                        MessageBox.Show("SN号：" + txtSNnumber.Text + " 已输入");
                        txtSNnumber.Text = "";
                        return;
                    }
                    if (factsampleqty >= int.Parse(txtsampleqty.Text))  ///////int.Parse(this.txtfactsampleqty.Text.Trim())
                    {
                        this.lblinfo.Text = "已经输入完成,不能超过抽样数量:" + txtsampleqty.Text;
                        lblinfo.ForeColor = Color.Red;
                        this.txtSNnumber.Text = "";
                        return;
                    }

                    SNnumberlist.Add(txtSNnumber.Text);
                    factsampleqty++;
                    txtfactsampleqty.Text = factsampleqty.ToString();
                    lblinfo.ForeColor = Color.Green;
                    lblinfo.Text = txtSNnumber.Text + " 输入完成，已输入 " + factsampleqty.ToString() + " 个";
                    sNnumber = txtSNnumber.Text;
                    txtSNnumber.Text = "";

                    if (int.Parse(txtsendqty.Text) == factsampleqty)
                    {
                        factsampleqty = 0;
                        lblinfo.ForeColor = Color.Green;
                        lblinfo.Text = "送检条码已全部输入完毕";
                        txtSNnumber.ReadOnly = true;
                    }
                    if (txtsampleqty.Text == this.txtfactsampleqty.Text.Trim())
                    {
                        txtQC.Focus();
                    }
                    else
                    {
                        //this.lblinfo.Text = "";
                        txtSNnumber.Enabled = true;
                        txtSNnumber.Focus();
                    }
                }
            }
        }

        public static DataTable CreateTable()
        {
            DataTable namesTable = new DataTable("SNbaddetails");
            DataColumn SNnumber = new DataColumn();
            SNnumber.DataType = System.Type.GetType("System.String");
            SNnumber.ColumnName = "SNnumber";
            namesTable.Columns.Add(SNnumber);

            DataColumn badnumber = new DataColumn();
            badnumber.DataType = System.Type.GetType("System.String");
            badnumber.ColumnName = "badnumber";
            namesTable.Columns.Add(badnumber);

            DataColumn badclass = new DataColumn();
            badclass.DataType = System.Type.GetType("System.String");
            badclass.ColumnName = "badclass";
            namesTable.Columns.Add(badclass);

            DataColumn baddescribe = new DataColumn();
            baddescribe.DataType = System.Type.GetType("System.String");
            baddescribe.ColumnName = "baddescribe";
            namesTable.Columns.Add(baddescribe);
            return namesTable;
        }

        static int NGCount = 0;

        ArrayList badnumberlist = new ArrayList();
        DataTable table = CreateTable();
        private void sBtnadd_Click(object sender, EventArgs e)
        {

            if (lblSNstates.Text == "需要扫描SN号")
            {
                if (txtbadnumber.Text == "" || txtbadclass.Text == "" || txtbaddescribe.Text == "")
                {
                    MessageBox.Show("请选择不良信息！", "停止", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }
                DialogResult result = MessageBox.Show("是否确定有不良？", "询问", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (DialogResult.Yes == result)
                {
                    if (!JudgeSNnumber(sNnumber, badnumberlist))
                    {
                        badnumberlist.Add(sNnumber);
                        NGCount = NGCount + 1;
                        txtNGnumber.Text = NGCount.ToString();
                    }
                }
                else
                    return;

                //if (int.Parse(txtNGnumber.Text) >= int.Parse(txtMAallowqty.Text))
                //{
                //    MessageBox.Show("NG数量已大于允收数量，该批拒收！", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //}
                string sql = @" select defects from OQC_baditem where badclass='" + txtbadclass.Text + "' and badphenomenon='" + txtbaddescribe.Text + "'";
                DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
                string defects = "";
                if (dt != null && dt.Rows.Count > 0)
                {

                    defects = dt.Rows[0]["defects"].ToString();
                    if (defects == "MA")
                    {
                        txtMAqty.Text = (int.Parse(txtMAqty.Text) + 1).ToString();
                    }
                    if (defects == "MI")
                    {
                        txtMIqty.Text = (int.Parse(txtMIqty.Text) + 1).ToString();
                    }
                    DataRow row;
                    row = table.NewRow();
                    row["SNnumber"] = sNnumber;
                    row["badnumber"] = txtbadnumber.Text;
                    row["badclass"] = txtbadclass.Text;
                    row["baddescribe"] = txtbaddescribe.Text;
                    table.Rows.Add(row);
                    txtNGpoint.Text = table.Rows.Count.ToString();
                    baddetails.DataSource = table;
                }
                else
                {
                    if (txtbadclass.Text == "功能不良" || txtbadclass.Text == "其他")
                    {
                        if (txtbaddescribe.Text == "需手动输入")
                        {
                            MessageBox.Show("请手动输入不良现象", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }
                        txtMAqty.Text = (int.Parse(txtMAqty.Text) + 1).ToString();
                        DataRow row;
                        row = table.NewRow();
                        row["SNnumber"] = sNnumber;
                        row["badnumber"] = txtbadnumber.Text;
                        row["badclass"] = txtbadclass.Text;
                        row["baddescribe"] = txtbaddescribe.Text;
                        table.Rows.Add(row);
                        txtNGpoint.Text = table.Rows.Count.ToString();
                        baddetails.DataSource = table;
                    }
                }
                if ((int.Parse(txtMAqty.Text) > int.Parse(txtMAallowqty.Text)) || (int.Parse(txtMIqty.Text) > int.Parse(txtMIallowqty.Text)))
                {
                    MessageBox.Show("缺陷数量已大于允收数量，该批拒收！", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    ifflat = "是";
                }
                txtbadnumber.Text = "";
                txtbadclass.Text = "";
                txtbaddescribe.Text = "";

            }
            else
            {
                //if (txtbadclass.Text == "功能不良")
                //{
                //    if (txtbaddescribe.Text == "需手动输入")
                //    {
                //        MessageBox.Show("请手动输入不良现象", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //        return;
                //    }
                //}
                string sql = @" select defects from OQC_baditem where badclass='" + txtbadclass.Text+ "' and badphenomenon='"+ txtbaddescribe.Text + "'";
                DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
                string defects = "";
                if (dt != null && dt.Rows.Count > 0)
                {
                    defects = dt.Rows[0]["defects"].ToString();
                    if (defects == "MA")
                    {
                        txtMAqty.Text = (int.Parse(txtMAqty.Text) + 1).ToString();
                    }
                    if (defects == "MI")
                    {
                        txtMIqty.Text = (int.Parse(txtMIqty.Text) + 1).ToString();
                    }
                    DataRow row;
                    row = table.NewRow();
                    // row["SNnumber"] = sNnumber;
                    row["badnumber"] = txtbadnumber.Text;
                    row["badclass"] = txtbadclass.Text;
                    row["baddescribe"] = txtbaddescribe.Text;
                    table.Rows.Add(row);
                    txtNGpoint.Text = table.Rows.Count.ToString();
                    baddetails.DataSource = table;
                }
                else
                {
                    if (txtbadclass.Text == "功能不良" || txtbadclass.Text == "其他")
                    {
                        if (txtbaddescribe.Text == "需手动输入")
                        {
                            MessageBox.Show("请手动输入不良现象", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }
                        txtMAqty.Text = (int.Parse(txtMAqty.Text) + 1).ToString();
                        DataRow row;
                        row = table.NewRow();
                        row["SNnumber"] = sNnumber;
                        row["badnumber"] = txtbadnumber.Text;
                        row["badclass"] = txtbadclass.Text;
                        row["baddescribe"] = txtbaddescribe.Text;
                        table.Rows.Add(row);
                        txtNGpoint.Text = table.Rows.Count.ToString();
                        baddetails.DataSource = table;
                    }
                }
                if ((int.Parse(txtMAqty.Text) > int.Parse(txtMAallowqty.Text)) || (int.Parse(txtMIqty.Text) > int.Parse(txtMIallowqty.Text)))
                {
                    MessageBox.Show("缺陷数量已大于允收数量，该批拒收！", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    ifflat = "是";
                }
                txtbadnumber.Text = "";
                txtbadclass.Text = "";
                txtbaddescribe.Text = "";
            }
        }

        private void sBtndelete_Click(object sender, EventArgs e)
        {

            int i = gridbaddetails.FocusedRowHandle;
            if (i < 0)
            {
                MessageBox.Show("请选中需要删除的信息", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            string SNnumber = gridbaddetails.GetFocusedRowCellValue("SNnumber").ToString();
            string badnumber = gridbaddetails.GetFocusedRowCellValue("badnumber").ToString();
            string badclass = gridbaddetails.GetFocusedRowCellValue("badclass").ToString();
            string baddescribe = gridbaddetails.GetFocusedRowCellValue("baddescribe").ToString();

            if (lblSNstates.Text == "需要扫描SN号")
            {
               
                int count = 0;
                for (int m = 0; m <= table.Rows.Count - 1; m++)
                {
                    if (table.Rows[m]["SNnumber"].ToString() == SNnumber)
                    {
                        count = count + 1;
                    }
                }
                if (count == 1)
                {
                    txtNGnumber.Text = (int.Parse(txtNGnumber.Text) - 1).ToString();
                }

                string sql = @" select defects from OQC_baditem where badclass='" +badclass+ "' and badphenomenon='" +baddescribe+ "'";
                DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
                string defects = "";
                if (dt != null && dt.Rows.Count > 0)
                {
                    defects = dt.Rows[0]["defects"].ToString();
                    if (defects == "MA")
                    {
                        txtMAqty.Text = (int.Parse(txtMAqty.Text) - 1).ToString();
                    }
                    if (defects == "MI")
                    {
                        txtMIqty.Text = (int.Parse(txtMIqty.Text) - 1).ToString();
                    }
                }
                else
                {
                    if (badclass == "功能不良" || badclass == "其他")
                    {
                      txtMAqty.Text = (int.Parse(txtMAqty.Text) - 1).ToString();
                    }
                }
                int n;
                for (n = table.Rows.Count - 1; n >= 0; n--)
                {
                    if ((table.Rows[n]["SNnumber"].ToString() == SNnumber) && (table.Rows[n]["badnumber"].ToString() == badnumber) && (table.Rows[n]["badclass"].ToString() == badclass) && (table.Rows[n]["baddescribe"].ToString() == baddescribe))
                    {
                        table.Rows[n].Delete();
                    }
                }
                txtNGpoint.Text = (int.Parse(txtNGpoint.Text) - 1).ToString();
                baddetails.DataSource = table;

                if ((int.Parse(txtMAqty.Text) < int.Parse(txtMAallowqty.Text)) && (int.Parse(txtMIqty.Text) < int.Parse(txtMIallowqty.Text)))
                {
                    ifflat = "否";                   
                }

                if ((txtMAqty.Text == "0" && txtMAallowqty.Text == "0") && (int.Parse(txtMIqty.Text) < int.Parse(txtMIallowqty.Text)))
                {
                    ifflat = "否";
                }
                if ((int.Parse(txtMAqty.Text) < int.Parse(txtMAallowqty.Text)) && (txtMIqty.Text == "0" && txtMIallowqty.Text == "0"))
                {
                    ifflat = "否";
                }
                if ((txtMAqty.Text == "0" && txtMAallowqty.Text == "0") && (txtMIqty.Text == "0" && txtMIallowqty.Text == "0"))
                {
                    ifflat = "否";
                }
            }
            else
            {

                string sql = @" select defects from OQC_baditem where badclass='" + badclass + "' and badphenomenon='" + baddescribe + "'";
                DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
                string defects = "";
                if (dt != null && dt.Rows.Count > 0)
                {
                    defects = dt.Rows[0]["defects"].ToString();
                    if (defects == "MA")
                    {
                        txtMAqty.Text = (int.Parse(txtMAqty.Text) - 1).ToString();
                    }
                    if (defects == "MI")
                    {
                        txtMIqty.Text = (int.Parse(txtMIqty.Text) - 1).ToString();
                    }
                }
                else
                {
                    if (badclass == "功能不良" || badclass == "其他" )
                    {
                        txtMAqty.Text = (int.Parse(txtMAqty.Text) - 1).ToString();
                    }
                }
                int n;
                for (n = table.Rows.Count - 1; n >= 0; n--)
                {
                    if ((table.Rows[n]["SNnumber"].ToString() == SNnumber) && (table.Rows[n]["badnumber"].ToString() == badnumber) && (table.Rows[n]["badclass"].ToString() == badclass) && (table.Rows[n]["baddescribe"].ToString() == baddescribe))
                    {
                        table.Rows[n].Delete();
                    }
                }
                txtNGpoint.Text = (int.Parse(txtNGpoint.Text) - 1).ToString();
                baddetails.DataSource = table;
            }

        }

        //static int NGCount = 0;

        //ArrayList badnumberlist = new ArrayList();

        private void txtbadnumber_Leave(object sender, EventArgs e)
        {
          
            if (txtbadnumber.Text == null)
                return;
            //DialogResult result = MessageBox.Show("是否确定有不良？","询问", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            //if (DialogResult.Yes == result)
            //{
            //    if (!JudgeSNnumber(sNnumber, badnumberlist))
            //    {
            //        badnumberlist.Add(sNnumber);
            //        NGCount = NGCount + 1;
            //        txtNGnumber.Text = NGCount.ToString();
            //    }

            //    else
            //    {

            //    }
            //}
            //else
            //    return;

        }
        string itemresult = "";
        private void sBtnOK_Click(object sender, EventArgs e)
        {
            itemresult = "OK";
            int i = gridtestcontent.FocusedRowHandle;
            if (i < 0)
            {
                MessageBox.Show("请选择检验的内容", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }            
        }

        private void sBtnNG_Click(object sender, EventArgs e)
        {
            itemresult = "NG";
            int i = gridtestcontent.FocusedRowHandle;
            if (i < 0)
            {
                MessageBox.Show("请选择检验的内容", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
        }

        public string AddOQCTestRecord( string opertype, string item , string customer, string cuscode,string model , string deliveryID,string hytcode
                                    ,string workno ,string serialnumber ,string sendqty,string org_id,string sampleqty,string sampleplan,string MA,string MI
                                    ,string checktype,string allowqty,string lineid ,string latyper,string QC,string masters,string QE,string CartonNo,string productionphase
                                    ,string ECNnumber,string factsampleqty,string NGQty,string NGpoint ,string rsno ,string productstate ,string teststandard ,string testremark
                                    ,string testresult ,string testman ,string checkman,string Auditman, string badinformation)
        {
            SqlParameter[] para = new SqlParameter[38];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@item", item);
            para[2] = new SqlParameter("@customer", customer);
            para[3] = new SqlParameter("@cuscode", cuscode);
            para[4] = new SqlParameter("@model", model);
            para[5] = new SqlParameter("@deliveryID", deliveryID);
            para[6] = new SqlParameter("@hytcode", hytcode);
            para[7] = new SqlParameter("@workno", workno);
            para[8] = new SqlParameter("@serialnumber", serialnumber);
            para[9] = new SqlParameter("@sendqty", sendqty);
            para[10] = new SqlParameter("@org_id", org_id);
            para[11] = new SqlParameter("@sampleqty", sampleqty);
            para[12] = new SqlParameter("@sampleplan", sampleplan);
            para[13] = new SqlParameter("@MA", MA);
            para[14] = new SqlParameter("@MI", MI);
            para[15] = new SqlParameter("@checktype", checktype);
            para[16] = new SqlParameter("@allowqty", allowqty);
            para[17] = new SqlParameter("@lineid", lineid);
            para[18] = new SqlParameter("@latyper", latyper);
            para[19] = new SqlParameter("@QC", QC);
            para[20] = new SqlParameter("@masters", masters);
            para[21] = new SqlParameter("@QE", QE);
            para[22] = new SqlParameter("@CartonNo", CartonNo);
            para[23] = new SqlParameter("@productionphase", productionphase);
            para[24] = new SqlParameter("@ECNnumber", ECNnumber);
            para[25] = new SqlParameter("@factsampleqty", factsampleqty);
            para[26] = new SqlParameter("@NGQty", NGQty);
            para[27] = new SqlParameter("@NGpoint", NGpoint);
            para[28] = new SqlParameter("@rsno", rsno);
            para[29] = new SqlParameter("@productstate", productstate);
            para[30] = new SqlParameter("@teststandard", teststandard);
            para[31] = new SqlParameter("@testremark", testremark);
            para[32] = new SqlParameter("@testresult", testresult);
            para[33] = new SqlParameter("@testman", testman);
            para[34] = new SqlParameter("@checkman", checkman);
            para[35] = new SqlParameter("@Auditman", Auditman);
            para[36] = new SqlParameter("@badinformation", badinformation);
            para[37] = new SqlParameter("@msg", SqlDbType.VarChar, 100);
            para[37].Direction = ParameterDirection.Output;
            DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "OQC_AddTestListNew", para);
            return para[37].Value.ToString();
        }

        public DataSet AddOQCTestRecordSearch(string opertype, string item, string customer, string cuscode, string model, string deliveryID, string hytcode
                                    , string workno, string serialnumber, string sendqty, string org_id, string sampleqty, string sampleplan, string MA, string MI
                                    , string checktype, string allowqty, string lineid, string latyper, string QC, string masters, string QE, string CartonNo, string productionphase
                                    , string ECNnumber, string factsampleqty, string NGQty, string NGpoint, string rsno, string productstate, string teststandard, string testremark
                                    , string testresult, string testman, string checkman, string Auditman)
        {
            SqlParameter[] para = new SqlParameter[37];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@item", item);
            para[2] = new SqlParameter("@customer", customer);
            para[3] = new SqlParameter("@cuscode", cuscode);
            para[4] = new SqlParameter("@model", model);
            para[5] = new SqlParameter("@deliveryID", deliveryID);
            para[6] = new SqlParameter("@hytcode", hytcode);
            para[7] = new SqlParameter("@workno", workno);
            para[8] = new SqlParameter("@serialnumber", serialnumber);
            para[9] = new SqlParameter("@sendqty", sendqty);
            para[10] = new SqlParameter("@org_id", org_id);
            para[11] = new SqlParameter("@sampleqty", sampleqty);
            para[12] = new SqlParameter("@sampleplan", sampleplan);
            para[13] = new SqlParameter("@MA", MA);
            para[14] = new SqlParameter("@MI", MI);
            para[15] = new SqlParameter("@checktype", checktype);
            para[16] = new SqlParameter("@allowqty", allowqty);
            para[17] = new SqlParameter("@lineid", lineid);
            para[18] = new SqlParameter("@latyper", latyper);
            para[19] = new SqlParameter("@QC", QC);
            para[20] = new SqlParameter("@masters", masters);
            para[21] = new SqlParameter("@QE", QE);
            para[22] = new SqlParameter("@CartonNo", CartonNo);
            para[23] = new SqlParameter("@productionphase", productionphase);
            para[24] = new SqlParameter("@ECNnumber", ECNnumber);
            para[25] = new SqlParameter("@factsampleqty", factsampleqty);
            para[26] = new SqlParameter("@NGQty", NGQty);
            para[27] = new SqlParameter("@NGpoint", NGpoint);
            para[28] = new SqlParameter("@rsno", rsno);
            para[29] = new SqlParameter("@productstate", productstate);
            para[30] = new SqlParameter("@teststandard", teststandard);
            para[31] = new SqlParameter("@testremark", testremark);
            para[32] = new SqlParameter("@testresult", testresult);
            para[33] = new SqlParameter("@testman", testman);
            para[34] = new SqlParameter("@checkman", checkman);
            para[35] = new SqlParameter("@Auditman", Auditman);
            para[36] = new SqlParameter("@msg", SqlDbType.VarChar, 100);
            para[36].Direction = ParameterDirection.Output;
            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "OQC_AddTestListNew", para);
        }

        public string OQC_SampleAddList(string opertype, string org_id, string workno, string SNnumber, string CartonNo, string VersionNO, string badnumber, string badclass, string baddescribe, string items)
        {
            SqlParameter[] para = new SqlParameter[11];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@org_id", org_id);
            para[2] = new SqlParameter("@workno", workno);
            para[3] = new SqlParameter("@SNnumber", SNnumber);
            para[4] = new SqlParameter("@CartonNo", CartonNo);
            para[5] = new SqlParameter("@VersionNO", VersionNO);
            para[6] = new SqlParameter("@badnumber", badnumber);
            para[7] = new SqlParameter("@badclass", badclass);
            para[8] = new SqlParameter("@baddescribe", baddescribe);
            para[9] = new SqlParameter("@items", items);
            para[10] = new SqlParameter("@msg", SqlDbType.VarChar, 100);
            para[10].Direction = ParameterDirection.Output;
            DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "OQC_SampleAddList", para);
            return para[10].Value.ToString();
        }

        string item = "";

        private void sBtnreport_Click(object sender, EventArgs e)  //保存
        {
            if ((txtboxsno.ReadOnly == false || txtboxnumber.ReadOnly == false) && txtcustomer.Text != "F1")
            {
                MessageBox.Show("请输入箱号", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            if (txtworkno.Text == "") return;
            if (txtsendqty.Text == "") return;

            if (txtline.Text == null || txtQE.Text == null || txtmaster.Text == null || txtstate.Text == null)
            {
                MessageBox.Show("请选择线体、QE、责任主管、产品状态的内容", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            if (txtChecktype.Text == "暂停检验")
            {
                MessageBox.Show("该物料暂停检验，请联系QE重新审核", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            int m = 1, n = 0;
            if (!int.TryParse(this.txtsendqty.Text == "" ? "1" : this.txtsendqty.Text, out m))
            {
                MessageBox.Show("送检数量必须为数字", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                this.txtsendqty.Text = "";
                this.txtsendqty.Focus();
                return;
            }
            else
            {
                if (m <= 0)
                {
                    MessageBox.Show("送检数量必须为大于0", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    this.txtsendqty.Text = "";
                    this.txtsendqty.Focus();
                    return;
                }
            }

            if (txtNGnumber.Text == "")
                txtNGnumber.Text = "0";


            if (!int.TryParse(txtNGnumber.Text, out n))
            {

                MessageBox.Show("NG数量不正确", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
               // txtNGnumber.Text = "0";
                return;

            }

            string testresult = "OK";
            if (lblSNstates.Text == "需要扫描SN号")
            {

                if (txtsamleplan.Text != "全检")
                {
                    if (SNnumberlist.Count < 1)
                    {
                        MessageBox.Show("需要扫描SN号", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                    if (ifflat == "否")
                    {
                        if (int.Parse(txtfactsampleqty.Text == "" ? "0" : txtfactsampleqty.Text) < int.Parse(txtsampleqty.Text))
                        {
                            MessageBox.Show("实抽数量小于应抽数量", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            txtSNnumber.Focus();
                            return;
                        }
                    }
                    if ((int.Parse(txtMAqty.Text) > int.Parse(txtMAallowqty.Text)) || (int.Parse(txtMIqty.Text) > int.Parse(txtMIallowqty.Text)))
                    {
                        testresult = "NG";
                    }

                }
                else
                {
                    if (int.Parse(txtfactsampleqty.Text == "" ? "0" : txtfactsampleqty.Text) < int.Parse(txtsampleqty.Text))
                    {
                        MessageBox.Show("实抽数量小于应抽数量", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        txtfactsampleqty.Focus();
                        return;
                    }
                    else
                    {
                        if (int.Parse(txtfactsampleqty.Text == "" ? "0" : txtfactsampleqty.Text) > int.Parse(txtsendqty.Text))
                        {
                            MessageBox.Show("实抽数量大于送检数量", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            txtfactsampleqty.Focus();
                            return;
                        }
                    }
                    if ((int.Parse(txtMAqty.Text) > int.Parse(txtMAallowqty.Text)) || (int.Parse(txtMIqty.Text) > int.Parse(txtMIallowqty.Text)))
                    {
                        testresult = "NG";
                    }
                }

            }

            //if (SNnumberlist.Count > 0)
            //{
            //    if (int.Parse(txtfactsampleqty.Text == "" ? "0" : txtfactsampleqty.Text) < int.Parse(txtsampleqty.Text))
            //    {
            //        if (int.Parse(txtNGnumber.Text) < int.Parse(txtMAallowqty.Text))
            //        {
            //            MessageBox.Show("实抽数量小于应抽数量", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //            return;
            //        }
            //    }
            //}
            else
            {
                txtfactsampleqty.ReadOnly = false;
                if (int.Parse(txtfactsampleqty.Text == "" ? "0" : txtfactsampleqty.Text) < int.Parse(txtsampleqty.Text))
                {
                    MessageBox.Show("实抽数量小于应抽数量", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    txtfactsampleqty.Focus();
                    return;
                }
                else
                {
                    if (int.Parse(txtfactsampleqty.Text == "" ? "0" : txtfactsampleqty.Text) > int.Parse(txtsendqty.Text))
                    {
                        MessageBox.Show("实抽数量大于送检数量", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        txtfactsampleqty.Focus();
                        return;
                    }
                }
                if ((int.Parse(txtMAqty.Text) > int.Parse(txtMAallowqty.Text)) || (int.Parse(txtMIqty.Text) > int.Parse(txtMIallowqty.Text)))
                {
                    testresult = "NG";
                }
            }
            
            //string itemtest = gridtestcontent.GetFocusedRowCellValue("检验项目结果").ToString();
            string Testcontent = "";
           
            DataTable dt = testcontent.DataSource as DataTable;
            if (dt != null && dt.Rows.Count > 0)
            {
                string itemtest = "OK";
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string testietm = dt.Rows[i]["测试项目"].ToString();
                    string standardsequence = dt.Rows[i]["序号"].ToString();
                    itemtest = dt.Rows[i]["检验项目结果"].ToString();
                    if (itemtest == "")
                    {
                        //MessageBox.Show("还有检验项目没检查完！", "停止", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        //return;
                        itemtest = "OK";
                    }
                    //if (itemtest=="NG")
                    //{
                    //    testresult = "NG";
                    //}
                    string ss = testietm + "，" + standardsequence + "，"+ itemtest + "； ";
                    Testcontent += ss;
                }
            }


            string sql = "select replace(replace(replace(replace(convert(varchar(30),getdate(),121),'-',''),' ',''),':',''),'.','')";
            item = DbAccess.SelectBySql(sql).Tables[0].Rows[0][0].ToString();
            string org_id = txtorg_id.Text.Trim(), workno = txtworkno.Text.Trim();
            string badinformation = "";


            DataTable Baddetail = baddetails.DataSource as DataTable;
           if (Baddetail != null && Baddetail.Rows.Count > 0)
            {
                string SNnumber, badnumber, badclass, baddescribe;
                for (int i = 0; i < Baddetail.Rows.Count; i++)
                {
                    SNnumber = Baddetail.Rows[i]["SNnumber"].ToString();
                    badnumber = Baddetail.Rows[i]["badnumber"].ToString();
                    badclass = Baddetail.Rows[i]["badclass"].ToString();
                    baddescribe = Baddetail.Rows[i]["baddescribe"].ToString();
                    badinformation = badinformation + "[" + (i + 1).ToString() + "]【" + SNnumber + "；" + badnumber + "；" + badclass + "；" + baddescribe + "】";
                    OQC_SampleAddList("新增不良", org_id, workno, SNnumber,"","", badnumber, badclass, baddescribe, item);
                }
            }
            string[] snnumber = (string[])SNnumberlist.ToArray(typeof(string));
            for (int i = 0; i < snnumber.Length; i++)
            {
                OQC_SampleAddList("新增良品", org_id, workno, snnumber[i],"", "", "", "", "", item);
            }


            string serialnumber = ProductSerialnumber(txtworkno.Text.Trim());


            string Cartonstr = "";
            string[] CartonNoList = (string[])cartonlist.ToArray(typeof(string));
            StringBuilder stringbuilder = new StringBuilder();
            for (int i = 0; i < CartonNoList.Length - 1; i++)
            {
                stringbuilder.Append(CartonNoList[i]);
                stringbuilder.Append(",");
            }
            try
            {
                stringbuilder.Append(CartonNoList[CartonNoList.Length - 1]);
                Cartonstr = stringbuilder.ToString();
            }
            catch

            { }

            string flag =  AddOQCTestRecord("新增",item, txtcustomer.Text.Trim(),txtmodel.Text.Trim(), txtmodel.Text, txtdelivery.Text, txtPN.Text.Trim()
                                    , txtworkno.Text.Trim(),serialnumber, txtsendqty.Text, txtorg_id.Text, txtsampleqty.Text, txtsamleplan.Text, txtMA.Text, txtMI.Text
                                    , txtChecktype.Text, txtMAallowqty.Text, txtline.Text , txtlatyper.Text, txtQC.Text, txtmaster.Text, txtQE.Text, Cartonstr, txtProductionphase.Text
                                    , txtECNnumber.Text, txtfactsampleqty.Text, txtNGnumber.Text, txtNGpoint.Text, txtrsno.Text, txtstate.Text, Testcontent, txtremark.Text
                                    , testresult,Login.username,"","", badinformation);
            if (flag.IndexOf("完毕") >= 0)
            {
                MessageBox.Show(flag, "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                badinformation = "";
            }
            else
            {
                MessageBox.Show("保存失败", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            reseach(item);  // search 
            ///// System.Threading.Thread.Sleep(500);
            //// sBtnreset_Click(sender,e);
            Clearcontrol();
        }

        private void reseach(string item)
        {
            /*
            string customer = txtcustomer.Text.Trim();
            string deliveryid = txtdelivery.Text.Trim();
            string PN = txtPN.Text.Trim();
            string workno = txtworkno.Text.Trim();
            string org_id = txtorg_id.Text.Trim();
            string line = txtline.Text.Trim();
            string master = txtmaster.Text.Trim();
            string QE = txtQE.Text.Trim();
            */
            string Serialnumber = txtSerialnumber.Text;
           // string Item = item;

            
            string where = "where 1=1";
            if (!string.IsNullOrEmpty(item))
            {
                where += " and item like '%" + item + "%'";
            }

            /*
            if (!string.IsNullOrEmpty(Serialnumber))
            {
                where += " and serialnumber ='" + Serialnumber + "' ";
            }           
            if (!string.IsNullOrEmpty(customer))
            {
                where += " and customer ='" + customer + "' ";
            }
            if (!string.IsNullOrEmpty(deliveryid))
            {
                where += " and deliveryID ='" + deliveryid + "' ";
            }
            if (!string.IsNullOrEmpty(PN))
            {
                where += " and hytcode ='" + PN + "' ";
            }
            if (!string.IsNullOrEmpty(workno))
            {
                where += " and workno ='" + workno + "' ";
            }
            if (!string.IsNullOrEmpty(org_id))
            {
                where += " and org_id ='" + org_id + "' ";
            }
            if (!string.IsNullOrEmpty(line))
            {
                where += " and lineid ='" + line + "' ";
            }
            if (!string.IsNullOrEmpty(master))
            {
                where += " and masters ='" + master + "' ";
            }
            if (!string.IsNullOrEmpty(QE))
            {
                where += " and QE ='" + QE + "' ";
            }

            */

            string sql = @" select checkdate 检查时间,customer 客户,model 机型,hytcode 编码,workno 工单号,serialnumber 抽样流水号,
                            sendqty 送检批量,org_id 组织,sampleqty 应抽数量,sampleplan 抽样计划,checktype 检验方式,allowqty 允收数量,CartonNo 箱号,productionphase 生产阶段,productstate 产品状态,
                            factsampleqty 实际抽样数量,NGQty NG数量,testresult 检查结果,testman 检验人,testremark 检验备注  from OQC_TestListNew  ";
            sql += where + " order by checkdate desc ";

            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            if (dt != null && dt.Rows.Count > 0)
            {
                gridControl.DataSource = dt;
            }
            else
            {
                MessageBox.Show("没有刷新！","提醒",MessageBoxButtons.OK,MessageBoxIcon.Information);

            }

        }

        void Clearcontrol()
        {
            table.Rows.Clear();
            baddetails.DataSource = null;
            testcontent.DataSource = null;
            cartonlist.Clear();
            TestNum = 0;
            CartonNum = 0;
            SNnumberlist.Clear();
            badnumberlist.Clear();
            factsampleqty = 0;
            txtfactsampleqty.Text = "0";
            txtsendqty.Text = "";
            txtsampleqty.Text = "";
            txtQC.Text = "";
            txtlatyper.Text = "";
            txtremark.Text = "";
            txtworkno.Text = "";
            lblinfo.Text = "";
            txtcustomer.Text = "";
            txtmodel.Text = "";
            txtdelivery.Text = "";
            txtPN.Text = "";
            txtworknumber.Text = "";
            txtproductdate.Text = "";
            txtSerialnumber.Text = "";
            txtsamleplan.Text = "";
            txtMA.Text = "";
            txtMI.Text = "";
            txtMAallowqty.Text = "";
            txtlatyper.Text = "";
            txtQC.Text = "";
            txtECNnumber.Text = "";
            txtNGnumber.Text = "";
            txtNGpoint.Text = "";
            txtrsno.Text = "";
            txtbadnumber.Text = "";
            txtbaddescribe.Text = "";
            txtboxsno.Text = "";
            txtboxnumber.Text = "";
            txtremark.Text = "";
            NGCount = 0;
            sNnumber = "";
            lblSNstates.Text ="";
            txtMAallowqty.Text = "0";
            txtMIallowqty.Text = "0";
            txtMAqty.Text = "0";
            txtMIqty.Text = "0";
            txtboxsno.ReadOnly = false;
            txtboxnumber.ReadOnly = false; 
            txtSNnumber.Text = "";
            txtSNnumber.ReadOnly = false;
            txtworkno.ReadOnly = false;
            txtsendqty.ReadOnly = true;
            ifflat = "否";
            txtbadclass.Text = "";
        }

        private void sBtnreset_Click(object sender, EventArgs e)
        {
            Clearcontrol();
            gridControl.DataSource = null;

            //table.Rows.Clear();
            //baddetails.DataSource = null;
            //testcontent.DataSource = null;
            //gridControl.DataSource = null;
            //cartonlist.Clear();
            //TestNum = 0;
            //CartonNum = 0;
            //SNnumberlist.Clear();
            //badnumberlist.Clear();
            //factsampleqty = 0;
            //txtfactsampleqty.Text = "0";
            //txtsendqty.Text = "";
            //txtsampleqty.Text = "";
            //txtQC.Text = "";
            //txtlatyper.Text = "";
            //txtremark.Text = "";
            //txtworkno.Text = "";
            //lblinfo.Text = "";
            //txtcustomer.Text = "";
            //txtmodel.Text = "";
            //txtdelivery.Text = "";
            //txtPN.Text = "";
            //txtworknumber.Text = "";
            //txtproductdate.Text = "";
            //txtSerialnumber.Text = "";
            //txtsamleplan.Text = "";
            //txtMA.Text = "";
            //txtMI.Text = "";
            //txtallowqty.Text = "";
            //txtlatyper.Text = "";
            //txtQC.Text = "";
            //txtECNnumber.Text = "";
            //txtNGnumber.Text = "";
            //txtNGpoint.Text = "";
            //txtrsno.Text = "";
            //txtbadnumber.Text = "";
            //txtbaddescribe.Text = "";
            //txtboxsno.Text = "";
            //txtboxnumber.Text = "";
            //txtremark.Text = "";
            //NGCount = 0;
            //sNnumber = "";
            //txtboxsno.ReadOnly = false;
            //txtboxnumber.ReadOnly = false;
            //txtSNnumber.Text = "";
            //txtSNnumber.ReadOnly = false;
            //txtworkno.ReadOnly = false;
        }
        private void gridtestcontent_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            /*
            int hand = e.RowHandle;
            if (hand < 0) return;      
            if(testresult =="NG")      
            e.Appearance.BackColor = Color.Red ;
            if(testresult == "OK")
            e.Appearance.BackColor = Color.Yellow;
            */
        }

        private void sBtndeleteresult_Click(object sender, EventArgs e)
        {

            int i = gridView.FocusedRowHandle;
            if (i < 0)
            {
                MessageBox.Show("请选中需要删除的信息", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null && dt.Rows.Count < 0)
            {
                MessageBox.Show("没有数据可删除", "停止", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            string item = gridView.GetFocusedRowCellValue("抽样流水号").ToString();
            string org_id = gridView.GetFocusedRowCellValue("组织").ToString();
            string workno = gridView.GetFocusedRowCellValue("工单号").ToString();

            string flag =  AddOQCTestRecord("删除", item, "", "", "", "", ""  , workno, "", "", org_id, "", "", "", "" , "", "", "", "", "", "", "", "", "" , "", "", "", "", "", "", "", ""
                                    , "", "", "", "","");
            if (flag.Contains("删除OK"))
            {
                MessageBox.Show("删除成功", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                gridControl.DataSource = null;
            }
            else
            {
                MessageBox.Show("删除失败", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }


        }

        private void OQCTestList_FormClosing(object sender, FormClosingEventArgs e)
        {
            sBtnreset_Click(sender,e);
        }

        private void gridView_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
        {

            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 0)
            {
                return;
            }
            if (gridView.GetDataRow(e.RowHandle)["检查结果"].ToString() == "NG")
            {
                e.Appearance.BackColor = Color.Red;
            }

        }

        private void txtbaddescribe_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (txtbaddescribe.Text == "需手动输入")
            {
                txtbaddescribe.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard;
                // txtbaddescribe.Text = "";
                txtbaddescribe.Focus();
            }
            else
            {
               // txtbaddescribe.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor;
            }
        }

        private void txtstate_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lblSNstates.Text == "不需要扫描SN号")
            {
                if (txtstate.Text != "正常品" || txtProductionphase.Text != "正常量产")
                {
                    txtsamleplan.Text = "全检";
                    txtsendqty_Leave(sender, e);
                }
                else
                {
                    JudgeTestItem(txtcustomer.Text.Trim(), txtPN.Text.Trim());
                    txtsendqty_Leave(sender, e);
                }
            }
            else
            {
                if (txtstate.Text != "正常品" || txtProductionphase.Text != "正常量产")
                {
                    txtsamleplan.Text = "全检";
                    txtsendqty_Leave(sender, e);
                    txtfactsampleqty.ReadOnly = false;
                    txtNGnumber.ReadOnly = false;
                }
                else
                {
                    JudgeTestItem(txtcustomer.Text.Trim(), txtPN.Text.Trim());
                    txtsendqty_Leave(sender, e);
                }
            }
        }

        private void txtProductionphase_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lblSNstates.Text == "不需要扫描SN号")
            {
                if (txtProductionphase.Text != "正常量产" || txtstate.Text != "正常品")
                {
                    txtsamleplan.Text = "全检";
                    txtsendqty_Leave(sender, e);
                }
                else
                {
                    JudgeTestItem(txtcustomer.Text.Trim(), txtPN.Text.Trim());
                    txtsendqty_Leave(sender, e);
                }
            }
            else
            {
                if (txtProductionphase.Text != "正常量产" || txtstate.Text != "正常品")
                {
                    txtsamleplan.Text = "全检";
                    txtsendqty_Leave(sender, e);
                    txtfactsampleqty.ReadOnly = false;
                    txtNGnumber.ReadOnly = false;
                }
                else
                {
                    JudgeTestItem(txtcustomer.Text.Trim(), txtPN.Text.Trim());
                    txtsendqty_Leave(sender, e);
                }
            }
        }

        private void gridbaddetails_Click(object sender, EventArgs e)
        {
            DataTable dt = baddetails.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 0)
                return;


        }

        private void txtworkno_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13 && txtworkno.Text != "")
            {
                txtworkno_Leave( sender,e);
                txtSerialnumber.Text = ProductSerialnumber(txtworkno.Text.Trim());
            }

        }

        private void txtsendqty_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13 && txtworkno.Text != "")
            {
                if (int.Parse(txtsendqty.Text) > int.Parse(txtworknumber.Text))
                {
                    MessageBox.Show("送检数量不能大于工单数量","提醒",MessageBoxButtons.OK,MessageBoxIcon.Information);
                    return;
                }
                if (int.Parse(txtsendqty.Text) < 0 )
                {
                    MessageBox.Show("送检数量不能为负数", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                txtsendqty_Leave(sender, e);
                txtsendqty.ReadOnly = true;
            }
        }
        private void txtboxnumber_TextChanged(object sender, EventArgs e)
        {
            if (txtboxnumber.Text != "")
            {
                txtworkno.ReadOnly = true;
                txtboxnumber.ReadOnly = true;
            }
        }
    }
}