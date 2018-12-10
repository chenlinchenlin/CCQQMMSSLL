using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using DevExpress.XtraBars;
using System.IO;
using DX_QMS.Common;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraEditors;

namespace DX_QMS
{
    public partial class TestProgSet : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private Microsoft.Office.Interop.Excel.ApplicationClass appClsExcel = null;
        DataTable dtsubitem = null;
         IQC ic = new IQC();
        public TestProgSet()
        {
            InitializeComponent();
            //bindDeviceType();
            //bindTestTools();
            //bindTestAQL();
            //bindsampleType();
            //bindTestUnit();
            //txtIQCChecktype.SelectedIndex = 1;
            //txtAQL.SelectedIndex = 1;
            //BindReceiver("SQE审核");
            //setRule();
            //gridView.PrintExportProgress += new ProgressChangedEventHandler(gridView_PrintExportProgress);
        }
        private void bindDeviceType()
        {
            DataTable dt = ic.SelectTestTypeRecord("查询", "", "测试类别", "").Tables[0];           
            cbTestType.Properties.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
                cbTestType.Properties.Items.Add(row["TestType"]);
            }
           //////// cbTestType.SelectedIndex = 0;
        }
        private void bindTestTools()
        {
            DataTable dt = ic.SelectTestTypeRecord("查询", "", "测试工具", "").Tables[0];
            txttesttools.Properties.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
                txttesttools.Properties.Items.Add(row["TestType"]);
            }
           /////////////// txttesttools.SelectedIndex = 0;
        }
        private void bindTestAQL()
        {
            DataTable dt = DbAccess.SelectBySql("select AQL,AQLValue from IQC_TestAQL").Tables[0];
            txtAQL.DataSource = dt;
            txtAQL.DisplayMember = dt.Columns["AQL"].ToString();
            txtAQL.ValueMember = dt.Columns["AQLValue"].ToString();
        }


        string bindsamplevalue(string sampletype)
        {
            string samplevalue = "0";
            string sql = @" select  Samplevalue from  IQC_TestSampleType where  SampleType='"+ sampletype + "' ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            if (dt != null && dt.Rows.Count > 0)
            {
                samplevalue = dt.Rows[0]["Samplevalue"].ToString();
            }
            return samplevalue;
        }



        private void bindsampleType()
        {
            DataTable dt = DbAccess.SelectBySql("select  SampleType,Samplevalue from  IQC_TestSampleType").Tables[0];
            //txtsampletype.DataSource = dt;
            //txtsampletype.DisplayMember = dt.Columns["SampleType"].ToString();
            //txtsampletype.ValueMember = dt.Columns["Samplevalue"].ToString();
            txtsampletype.Properties.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
                txtsampletype.Properties.Items.Add(row["SampleType"]);
            }
           ///////////// txtsampletype.SelectedIndex = 0;

        }



        string bindTestUnitvalue(string unitname)
        {
            string unitvalue = "";
            string sql = @"  select u.unit  from ( select unit, (unit+unitname) name from IQC_TestUnit  ) u  where u.name = '"+ unitname+ "'  ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            if (dt != null && dt.Rows.Count > 0)
            {
                unitvalue = dt.Rows[0]["unit"].ToString();
            }
            return unitvalue;
        }


        string bindTestUnitname(string unitvalue)
        {
            string unitname = "";
            string sql = @"   select (unit+unitname) name from IQC_TestUnit where unit = '"+ unitvalue + "'  ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            if (dt != null && dt.Rows.Count > 0)
            {
                unitname = dt.Rows[0]["name"].ToString();
            }
            return unitname;
        }



        private void bindTestUnit()
        {
            string ssql = "select unit, (unit+unitname) name from IQC_TestUnit order by sort";
            DataTable dt = DbAccess.SelectBySql(ssql).Tables[0];
            //txtunit.DataSource = dt;
            //txtunit.DisplayMember = "name";
            //txtunit.ValueMember = "unit";

            txtunit.Properties.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
                txtunit.Properties.Items.Add(row["name"]);
            }
           //////////// txtunit.SelectedIndex = 0;

        }
        private void BindReceiver(string sType)
        {
            string sql = "select userName,userMail from MailGroup where mailType='" + sType + "'";
            DataSet ds = DbAccess.SelectBySql(sql);
            txtreceiver.DataSource = ds.Tables[0];
            txtreceiver.DisplayMember = "userName";
            txtreceiver.ValueMember = "userMail";
            txtreceiver.SelectedIndex = -1;
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
            Dictionary<string, bool> dic = GroupPermission.QMS_SelectRulesForForm(post, "程序");
            this.btnsave.Enabled = bool.Parse(dic["hasInsert"].ToString());
            this.btndel.Enabled = bool.Parse(dic["hasDelete"].ToString());
            this.btnOK.Enabled = bool.Parse(dic["hasUpdate"].ToString());
            this.btnAudit.Enabled = bool.Parse(dic["hasAudit"].ToString());
            this.btnImport.Enabled = bool.Parse(dic["hasPrint"].ToString());
        }

        public DateTime beforeTime, afterTime;
        private string IfCheck(string sType, string productcode)
        {
            string Flag = "不需审核";
            string sql = "select min(case when IFYiQi='否' then '不需审核' else ISNULL(states,'未审') end) states from  IQC_TestProgSet s inner join  IQC_TestType i on s.TestType=i.TestType where Productcode='" + productcode + "' and s.TestType='" + sType + "'";
            DataSet ds = DbAccess.SelectBySql(sql);
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
        private void bindTestItem(string TestType)
        {
            DataSet ds = ic.SelectTestItemRecord("查询", TestType, "", 0, "");
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                DataTable dt = ds.Tables[0];
                txttestitem.Properties.Items.Clear();
                foreach (DataRow row in dt.Rows)
                {
                    txttestitem.Properties.Items.Add(row["TestItem"]);
                }
                txttestitem.SelectedIndex = 0;
            }
        }
        private void txtproductcode_Leave(object sender, EventArgs e)
        {

            this.lblcheckstate.Text = "";
            gridView.Columns.Clear();
            databind.DataSource = null;
            if (txtproductcode.Text == "") return;
            txtproductcode.Leave -= txtproductcode_Leave;
            string Orasql = "select organization_id organization,segment1 materialcode,description materialname,en_des materialenname,INVENTORY_ITEM_ID,PRIMARY_UOM_CODE from cux_item_material_v where segment1='" + txtproductcode.Text.Trim() + "'";
            DataSet ds = DbAccess.SelectByOracle(Orasql);
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                lblinfo.Text = ds.Tables[0].Rows[0]["materialname"].ToString();
                lblinfo.ForeColor = Color.Blue;
                //绑定已维护过的数据
                DataSet dsprog = ic.SelectTestProgRecord("查询", txtproductcode.Text, "", "", "", lblinfo.Text, "", "", txttestdes.Text, "", "", 0, 0, 0, txtAQL.Text, (int)numericud.Value, 0, "", 0, 0);
                if (dsprog != null && dsprog.Tables.Count > 0 && dsprog.Tables[0].Rows.Count > 0)
                {
                    databind.DataSource = dsprog.Tables[0];
                    cbTestType.Text = dsprog.Tables[0].Rows[0]["测试类别"].ToString();

                    string F = IfCheck(cbTestType.Text, txtproductcode.Text);
                    if (F == "已审核")
                    {
                        this.lblcheckstate.Text = "已审核";
                        this.lblcheckstate.ForeColor = Color.Blue;
                        this.btnOK.Enabled = false;
                    }
                    else if (F == "未审")
                    {
                        this.lblcheckstate.Text = "未审核";
                        this.lblcheckstate.ForeColor = Color.Red;
                        this.btnOK.Enabled = true;
                    }
                    else if (F == "NG")
                    {
                        this.lblcheckstate.Text = "NG";
                        this.lblcheckstate.ForeColor = Color.Red;
                        this.btnOK.Enabled = true;
                    }
                    else
                    {
                        this.lblcheckstate.Text = "不需审核";
                        this.lblcheckstate.ForeColor = Color.Green;
                        btnOK.Enabled = false;
                        btnAudit.Enabled = false;
                    }
                }
                else
                    cbTestType.SelectedIndex = 0;
                cbTestType.Focus();
                bindTestItem(cbTestType.Text);
            }
            else
            {
                string ssql = "select materialcode from delivery where lotno='" + txtproductcode.Text + "'";
                DataSet dslotno = DbAccess.SelectBySql(ssql);
                if (dslotno != null && ds.Tables.Count > 0 && dslotno.Tables[0].Rows.Count > 0)
                {
                    string materialcode = dslotno.Tables[0].Rows[0]["materialcode"].ToString();
                    string Orasqlbylotno = "select organization_id organization,segment1 materialcode,description materialname,en_des materialenname,INVENTORY_ITEM_ID,PRIMARY_UOM_CODE from cux_item_material_v where segment1='" + materialcode + "'";
                    ds = DbAccess.SelectByOracle(Orasqlbylotno);
                    if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                    {
                        txtproductcode.Text = materialcode;
                        lblinfo.Text = ds.Tables[0].Rows[0]["materialname"].ToString();
                        lblinfo.ForeColor = Color.Blue;
                        //绑定已维护过的数据
                        DataSet dsprog = ic.SelectTestProgRecord("查询", txtproductcode.Text, "", "", "", lblinfo.Text, "", "", txttestdes.Text, "", "", 0, 0, 0, txtAQL.Text, (int)numericud.Value, 0, "", 0, 0);
                        if (dsprog != null && dsprog.Tables.Count > 0 && dsprog.Tables[0].Rows.Count > 0)
                        {
                            databind.DataSource = dsprog.Tables[0];
                            cbTestType.Text = dsprog.Tables[0].Rows[0]["测试类别"].ToString();

                            string F = IfCheck(cbTestType.Text, txtproductcode.Text);
                            if (F == "已审核")
                            {
                                this.lblcheckstate.Text = "已审核";
                                this.lblcheckstate.ForeColor = Color.Blue;
                                this.btnOK.Enabled = false;
                            }
                            else if (F == "未审")
                            {
                                this.lblcheckstate.Text = "未审核";
                                this.lblcheckstate.ForeColor = Color.Red;
                                this.btnOK.Enabled = true;
                            }
                            else
                            {
                                this.lblcheckstate.Text = "不需审核";
                                this.lblcheckstate.ForeColor = Color.Green;
                                btnOK.Enabled = false;
                                btnAudit.Enabled = false;
                            }
                        }
                        else
                            cbTestType.SelectedIndex = 0;
                        cbTestType.Focus();
                        bindTestItem(cbTestType.Text);
                    }
                }
                else
                {
                    lblinfo.Text = txtproductcode.Text + "该料号不存在";
                    lblinfo.ForeColor = Color.Red;
                    txtproductcode.Text = "";
                    txtproductcode.Focus();
                }
            }
            txtproductcode.Leave += txtproductcode_Leave;
        }

        private void txtIQCChecktype_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txtIQCChecktype.Text.ToString() == "加严检验" || txtIQCChecktype.Text.ToString() == "放宽检验")
            {
                DataTable dt = DbAccess.SelectBySql("select SampleType,Samplevalue from IQC_TestSampleType where SampleType = 'ISO2859-1' ").Tables[0];
                //txtsampletype.DataSource = dt;
                //txtsampletype.DisplayMember = dt.Columns["SampleType"].ToString();
                //txtsampletype.ValueMember = dt.Columns["Samplevalue"].ToString();

                txtsampletype.Properties.Items.Clear();
                foreach (DataRow row in dt.Rows)
                {
                    txtsampletype.Properties.Items.Add(row["SampleType"]);
                }

            }

            else
            {
                DataTable dt = DbAccess.SelectBySql("select SampleType,Samplevalue from IQC_TestSampleType").Tables[0];
                //txtsampletype.DataSource = dt;
                //txtsampletype.DisplayMember = dt.Columns["SampleType"].ToString();
                //txtsampletype.ValueMember = dt.Columns["Samplevalue"].ToString();

                txtsampletype.Properties.Items.Clear();
                foreach (DataRow row in dt.Rows)
                {
                    txtsampletype.Properties.Items.Add(row["SampleType"]);
                }
            }
        }

        private void cbTestType_SelectedIndexChanged(object sender, EventArgs e)
        {
            bindTestItem(cbTestType.Text);
        }
        private void bindTestsubItem(string testtype, string Testitem)
        {
            string ssql = "select  TestSubItem, Item, TestDesc, TestTool, PackType from IQC_TestSubItemSet where TestItem='" + Testitem + "' and TestType='" + testtype + "'";
            DataSet ds = DbAccess.SelectBySql(ssql);
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                DataTable dt = ds.Tables[0];
                dtsubitem = dt;
                txttestsubitem.Properties.Items.Clear();
                foreach (DataRow row in dt.Rows)
                {
                    txttestsubitem.Properties.Items.Add(row["TestSubItem"]);
                }
                txttestsubitem.SelectedIndex = 0;
                if (txttestsubitem.Properties.Items.Count == 1)
                {
                    txttestsubitem.SelectedIndex = 0;
                    DataRow[] rw = dtsubitem.Select("TestSubItem='" + this.txttestsubitem.Text + "'");
                    txttestdes.Text = rw[0]["TestDesc"].ToString();
                    txttesttools.Text = rw[0]["TestTool"].ToString();
                    txtpacktype.Text = rw[0]["PackType"].ToString();
                    if (txttestsubitem.Text.Contains("包装") || txttestsubitem.Text.Contains("核对"))
                        txtsampletype.Text = "全检";
                    else if (txttestsubitem.Text.Contains("测试") && (cbTestType.Text.Contains("电阻类") || cbTestType.Text.Contains("电容类") || cbTestType.Text.Contains("电感类")))
                    {
                        txtsampletype.Text = "5pcs/批";
                    }
                    else
                        txtsampletype.SelectedIndex = 0;
                }
                else
                {
                    txttestsubitem.Focus();
                    return;
                }
            }

        }
        private void txttestitem_SelectedIndexChanged(object sender, EventArgs e)
        {
            bindTestsubItem(cbTestType.Text, txttestitem.Text == "" ? "" : txttestitem.Text);
        }

        private void txttestsubitem_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txttestsubitem.Text == "") return;

            DataRow[] rw = dtsubitem.Select("TestSubItem='" + this.txttestsubitem.Text + "'");

            txttestdes.Text = rw[0]["TestDesc"].ToString();
            txttesttools.Text = rw[0]["TestTool"].ToString();
            txtpacktype.Text = rw[0]["PackType"].ToString();

            if (txttestsubitem.Text.Contains("包装") || txttestsubitem.Text.Contains("核对"))
                txtsampletype.Text = "全检";
            else if (txttestsubitem.Text.Contains("测试") && (cbTestType.Text.Contains("电阻类") || cbTestType.Text.Contains("电容类") || cbTestType.Text.Contains("电感类")))
            {
                txtsampletype.Text = "5pcs/批";
            }
            else
                txtsampletype.SelectedIndex = 0;
        }

        private void txtsampletype_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txtsampletype.Text.ToString() == "ISO2859-1")
            {
                DataTable dt = DbAccess.SelectBySql("select distinct AQL,AQLValue from IHPS_QUALITY_SPC_AQLIS02859 where AQL='AQL=0.40' or AQL='AQL=1.0' or AQL='AQL=0.010' or AQL='AQL=0.10' or AQL='AQL=0.65' or AQL='AQL=1.5' ").Tables[0];
                txtAQL.DataSource = null;
                txtAQL.DataSource = dt;
                txtAQL.DisplayMember = dt.Columns["AQL"].ToString();
                txtAQL.ValueMember = dt.Columns["AQLValue"].ToString();
            }
            else if (txtsampletype.Text.ToString() == "C=0")
            {
                DataTable dt = DbAccess.SelectBySql(" select distinct AQL,AQLValue from IHPS_QUALITY_SPC_AQLC0 where AQL='AQL=0.4' or AQL='AQL=1' or AQL='AQL=0.01' or AQL='AQL=0.1' or AQL='AQL=0.65' or AQL='AQL=1.5' ").Tables[0];
                txtAQL.DataSource = null;
                txtAQL.DataSource = dt;
                txtAQL.DisplayMember = dt.Columns["AQL"].ToString();
                txtAQL.ValueMember = dt.Columns["AQLValue"].ToString();
            }
            else if (txtsampletype.Text.ToString() == "计数")
            {
                txtAQL.Text = "";
            }
            else
            {
                txtAQL.DataSource = null;
                bindTestAQL();
            }

        }
        private DataTable BindTestProgRecord()
        {
            DataTable dt = ic.SelectTestProgRecord("查询", txtproductcode.Text.Trim(),"","","", lblinfo.Text,"", "", txttestdes.Text,"", "", 0, 0, 0, txtAQL.Text, (int)numericud.Value, 0, "", 0, 0).Tables[0];
            return dt;
        }
        private void btnsave_Click(object sender, EventArgs e)
        {
            if (txtproductcode.Text == "") return;
            //20150418新增加审核控制
            if (lblcheckstate.Text == "已审核" && btnAudit.Enabled == false)
            {
                MessageBox.Show(txtproductcode.Text + ",该物料已经审核,你没权再修改", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            int m = 1;
            if (!int.TryParse(txttestvalueqty.Text == "" ? "1" : txttestvalueqty.Text, out m))
            {
                MessageBox.Show("测试值个数必须为数字", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                txttestvalueqty.Text = "";
                txttestvalueqty.Focus();
                return;
            }
            else
            {
                if (m <= 0)
                {
                    MessageBox.Show("测试值个数必须大于0", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    txttestvalueqty.Text = "";
                    txttestvalueqty.Focus();
                    return;
                }
            }

            float n = 0;
            if (!float.TryParse(txtLowervalue.Text == "" ? "0" : txtLowervalue.Text, out n))
            {
                MessageBox.Show("测试值范围值不正确", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                txtLowervalue.Text = "";
                txtLowervalue.Focus();
                return;
            }
            if (!float.TryParse(txtUppervalue.Text == "" ? "0" : txtUppervalue.Text, out n))
            {
                MessageBox.Show("测试值范围值不正确", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                txtUppervalue.Text = "";
                txtUppervalue.Focus();
                return;
            }

            string unit = bindTestUnitvalue(txtunit.Text);

            try
            {
                int i = ic.AddNewTestProgRecord("新增", txtproductcode.Text, txtIQCChecktype.Text, cbTestType.Text, txtpacktype.Text, lblinfo.Text, txttestitem.Text, txttestsubitem.Text, txttestdes.Text,
                    txttesttools.Text, txtsampletype.Text, txtUppervalue.Text == "" ? "0" : txtUppervalue.Text, txtLowervalue.Text == "" ? "0" : txtLowervalue.Text, float.Parse(txtAQL.SelectedValue.ToString()), txtAQL.Text, (int)numericud.Value, m, unit == "" ? "mm" : unit, (int)txtcheckcycle.Value, (int)txtleadtime.Value);
            }
            catch
            {
                MessageBox.Show("请确认物料编码、测试项目和测试子项目不为空！","提醒",MessageBoxButtons.OK ,MessageBoxIcon.Information);
                return;
            }
            txtLowervalue.Text = "";
            txtUppervalue.Text = "";
            txttestvalueqty.Text = "";
            numericud.Value = 5;
            gridView.Columns.Clear();
            databind.DataSource = BindTestProgRecord();
        }

        private void btndel_Click(object sender, EventArgs e)
        {
            if (gridView.FocusedRowHandle < 0)
                return;
            if (gridView.GetSelectedRows().Length < 1)
                return;

            if (lblcheckstate.Text == "已审核" && btnAudit.Enabled == false)
            {
                MessageBox.Show(txtproductcode.Text + ",该物料已经审核,你没权再删除", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }


            for (int k = gridView.GetSelectedRows().Length; k > 0; k--)
            {


                DataRow dr = gridView.GetDataRow(gridView.GetSelectedRows()[k-1]);
                int i = ic.AddNewTestProgRecord("删除", dr["产品编码"].ToString(), dr["检验方式"].ToString(), dr["测试类别"].ToString(), dr["包装方式"].ToString(), "",
                                                dr["测试项目"].ToString(), dr["测试子项目"].ToString(), "", "", "", "0", "0", 0, "", 0, 0, "", 0, 0);
            }
            gridView.Columns.Clear();
            databind.DataSource = ic.SelectTestProgRecord("查询", txtproductcode.Text, txtIQCChecktype.Text, cbTestType.Text, txtpacktype.Text, lblinfo.Text, txttestitem.Text, txttestsubitem.Text, txttestdes.Text,
                                            txttesttools.Text, txtsampletype.Text, float.Parse(txtUppervalue.Text == "" ? "0" : txtUppervalue.Text), float.Parse(txtLowervalue.Text == "" ? "0" : txtLowervalue.Text), float.Parse(txtAQL.SelectedValue.ToString()), txtAQL.Text, (int)numericud.Value, 0, "", 0, 0).Tables[0];

        }

        private void btnsearch_Click(object sender, EventArgs e)
        {
            gridView.Columns.Clear();
            databind.DataSource = null;  
            DataTable dt = BindTestProgRecord();
            if (dt == null || dt.Rows.Count < 1)
            {
                MessageBox.Show("没有数据","提醒",MessageBoxButtons.OK ,MessageBoxIcon.Information );
                return;
            }
            databind.DataSource = dt;
            gridView.RefreshData();
        }
        public void Copy(string sheetPrefixName, System.Data.DataTable tb, System.Data.DataTable dt)
        {

            string finaltotalstate = "", testuser = "";
            int sheetCount = 1;
            if (sheetPrefixName == null || sheetPrefixName.Trim() == "")
                sheetPrefixName = " Sheet ";
            beforeTime = DateTime.Now;
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = true;
            afterTime = DateTime.Now;
            object missing = System.Reflection.Missing.Value;
            string templetFile = Environment.CurrentDirectory + @"\Resources\Report.xls";

            Microsoft.Office.Interop.Excel.Workbook workBook = app.Workbooks.Open(templetFile, missing, true, missing, missing, missing,
                                                          missing, missing, missing, missing, missing, missing, missing, missing, missing); 
            Microsoft.Office.Interop.Excel.Worksheet workSheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.Sheets.get_Item(1);
            for (int i = 1; i < sheetCount; i++)
            {
                ((Microsoft.Office.Interop.Excel.Worksheet)workBook.Worksheets.get_Item(i)).Copy(missing, workBook.Worksheets[i]);
            }

            Microsoft.Office.Interop.Excel.Worksheet sheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.Worksheets.get_Item(1);
            sheet.Name = sheetPrefixName.Replace("/", "");

            string lotnos = "";
            System.Data.DataTable dtlotno = null;

            for (int j = 0; j < tb.Rows.Count; j++)
            {
                string finalstate = "";
                sheet.Cells[2, 1] = "IQC检验报表(" + dt.Rows[0 + j]["TestType"].ToString() + ")";
                sheet.Cells[5, 2] = dt.Rows[0 + j]["Productcode"].ToString();
                sheet.Cells[5, 7] = dt.Rows[0 + j]["materialenname"].ToString();

                sheet.Cells[10 + j + 1, 1] = dt.Rows[0 + j]["TestItem"].ToString();
                sheet.Cells[10 + j + 1, 2] = dt.Rows[0 + j]["TestSubItem"].ToString();
                sheet.Cells[10 + j + 1, 3] = dt.Rows[0 + j]["TestDesc"].ToString();
                sheet.Cells[10 + j + 1, 9] = dt.Rows[0 + j]["TestTool"].ToString();
                sheet.Cells[10 + j + 1, 11] = dt.Rows[0 + j]["SampleType"].ToString();
                if (dt.Rows[0 + j]["AQL"].ToString().StartsWith("MI"))
                {
                    sheet.Cells[10 + j + 1, 20] = "√";

                }
                else if (dt.Rows[0 + j]["AQL"].ToString().StartsWith("MA"))
                {
                    sheet.Cells[10 + j + 1, 19] = "√";
                }
                else if (dt.Rows[0 + j]["AQL"].ToString().StartsWith("CR=0.1"))
                {
                    sheet.Cells[10 + j + 1, 18] = "√";
                }

                else if (dt.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=0.40"))
                {
                    sheet.Cells[10 + j + 1, 21] = "√";
                }
                else if (dt.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=1.0"))
                {
                    sheet.Cells[10 + j + 1, 22] = "√";
                }
                else if (dt.Rows[0 + j]["AQL"].ToString().StartsWith("CR=0.01"))
                {
                    sheet.Cells[10 + j + 1, 23] = "√";
                }


                if (dt.Rows[0 + j]["CheckType"].ToString() == "正常检验")
                {
                    sheet.Cells[7, 6] = "√正常检验";
                }
                else if (dt.Rows[0 + j]["CheckType"].ToString() == "加严检验")
                {
                    sheet.Cells[7, 2] = "√加严检验";
                }
                else if (dt.Rows[0 + j]["CheckType"].ToString() == "放宽检验")
                {
                    sheet.Cells[7, 11] = "√放宽检验";
                }
                else if (dt.Rows[0 + j]["CheckType"].ToString() == "全检")
                {
                    sheet.Cells[7, 16] = "√全检";
                }
            }
            string ssql = "select approvalman, Auditingman from  IQC_Approval";
            DataSet ds = DbAccess.SelectBySql(ssql);
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                sheet.Cells[46, 9] = ds.Tables[0].Rows[0]["approvalman"].ToString();
                sheet.Cells[46, 13] = ds.Tables[0].Rows[0]["Auditingman"].ToString();
            }
        }
        private void btntoexcel_Click(object sender, EventArgs e)
        {
            string ssql = @"SELECT Productcode, CheckType, l.TestType, l.PackType, Productname materialenname, l.SampleType, l.TestItem, l.TestSubItem,case when (cast(l.LowValue as float)<>0 or cast(l.UpValue as float)<>0) then (l.TestDesc+'(范围:【'+cast(l.LowValue as varchar(10))+'】至:【'+cast(l.UpValue as varchar(10))+'】)') else l.TestDesc end TestDesc, l.TestTool, l.UpValue, l.LowValue, l.UpScope, 
                          l.AQLValue, l.AQL, l.Item, SubItem, lastupdatedate, testvalueqty, unit, checkcycle
                          FROM IQC_TestProgSet  l
                          left join IQC_TestSubItemSet s on l.TestSubItem=s.TestSubItem and l.TestItem=s.TestItem and l.TestType=s.TestType
                          left join IQC_TestItem t on s.TestItem=t.TestItem and s.TestType=t.TestType where l.Productcode='" + txtproductcode.Text + "' order by t.Item asc,s.Item";
            DataTable dt = DbAccess.SelectBySql(ssql).Tables[0];
            if (dt.Rows.Count > 0)
            {
                Copy(dt.Rows[0]["TestType"].ToString(), dt, dt);
            }

        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            string users = "";
            for (int i = 0; i < txtreceiver.SelectedItems.Count; i++)
            {
                users = users + "," + txtreceiver.GetItemText(txtreceiver.SelectedItems[i]);
            }
            users = users.TrimStart(',').TrimEnd(',');

            if (users == "")
            {
                MessageBox.Show("请选择收件人", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            string sql = "update IQC_TestProgSet set SendState='OK',states='未审',SendUser='" + Login.username + "',Receiver='" + users + "',SendDate=getdate() where Productcode='" + txtproductcode.Text + "' and TestType='" + cbTestType.Text + "'";
            if (DbAccess.ExecuteSql(sql))
            {
                string subject = "【重要信息】" + cbTestType.Text + ",物料编码:" + txtproductcode.Text + "需要您审核处理";
                string body = "QMS系统提醒您，类别：" + cbTestType.Text + ",物料编码:" + txtproductcode.Text + "，需要您审核，谢谢!";
                ProdTest.SendMailToSQE("Quality", subject, body, "SQE审核", users);
            }
        }

        private void btnAudit_Click(object sender, EventArgs e)
        {
            string state = "";
            if (!cbok.Checked && !cbNG.Checked)
            {
                MessageBox.Show("请选择一个审核结果", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (cbNG.Checked && txtremarks.Text == "")
            {
                MessageBox.Show("请输入NG备注", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (cbok.Checked)
            {
                state = "已审核";
            }
            else if (cbNG.Checked)
            {
                state = "NG";
            }

            string sql = "update IQC_TestProgSet set states= '" + state + "',RemarksSQE='" + txtremarks.Text + "',AuditingByUser='" + Login.username + "',AuditingDate=getdate() where Productcode='" + txtproductcode.Text + "' and TestType='" + cbTestType.Text + "'";
            if (DbAccess.ExecuteSql(sql))
                MessageBox.Show(txtproductcode.Text + ",审核成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

            txtremarks.Text = "";





        }

        private void btnAuditQuery_Click(object sender, EventArgs e)
        {
            string sql = "select distinct s.TestType,Productcode,max(Productname) Productname,isnull(min(states),'未审') states,max(AuditingByUser) AuditingByUser,max(AuditingDate) AuditingDate ,max(SendState) SendState,max(SendDate) SendDate,max(SendUser) SendUser,Max(RemarksSQE) Remarks from  IQC_TestProgSet s ";
            sql += " inner join  IQC_TestType i on s.TestType=i.TestType  where SendState='OK' and IFYiQi<>'否' and s.TestType='" + cbTestType.Text + "' and isnull(Receiver," + "'" + txtreceiver.Text + "')= case " + "'" + txtreceiver.Text + "' when '' then isnull(Receiver," + "'" + txtreceiver.Text + "') else '" + txtreceiver.Text + "' end and Productcode like " + "'" + txtproductcode.Text + "%' group by s.TestType,Productcode";
            DataSet ds = DbAccess.SelectBySql(sql);
            gridView.Columns.Clear();
            databind.DataSource = null;
            databind.DataSource = ds.Tables[0];

        }
        private void UpdateByProduct(string p)
        {
            string sql = "update IQC_TestProgSet set states='已审核',AuditingByUser='" + Login.username + "',AuditingDate=getdate() where Productcode in (" + p + ") and TestType='" + cbTestType.Text + "'";
            if (DbAccess.ExecuteSql(sql))
                MessageBox.Show(p + ",审核成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void btnImport_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofdImport = new OpenFileDialog();
            ofdImport.Filter = "Excel文件(*.xls,*.xlsx)|*.xls;*.xlsx";
            ofdImport.Multiselect = true;
            DialogResult dr = ofdImport.ShowDialog();
            if (dr == DialogResult.Cancel) return;
            string[] sFilesName = ofdImport.FileNames;

            object objMissing = System.Reflection.Missing.Value;
            this.appClsExcel = new Microsoft.Office.Interop.Excel.ApplicationClass();
            for (int i = 0; i < sFilesName.Length; i++)
            {
                if (!System.IO.File.Exists(sFilesName[i]))
                {
                    MessageBox.Show("Excel文件不存在！", "错误");
                    return;
                }
                else
                {
                    Microsoft.Office.Interop.Excel.Workbook wBookExcel = appClsExcel.Workbooks.Open(sFilesName[i], objMissing, true, objMissing, objMissing, objMissing
                            , objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing);

                    Microsoft.Office.Interop.Excel.Worksheet wSheetExcel = (Microsoft.Office.Interop.Excel.Worksheet)wBookExcel.ActiveSheet;
                    Microsoft.Office.Interop.Excel.Range rCell = wSheetExcel.UsedRange;
                    object[,] objList = (object[,])rCell.Value2;

                    int excelcounts = objList.GetLength(0);

                    string s = "";
                    for (int iRow = 1; iRow < excelcounts; iRow++)
                    {

                        try
                        {
                            s = s + ",'" + objList[iRow + 1, 1].ToString().Trim().ToUpper() + "'";

                        }
                        catch (Exception ex)
                        {
                            continue;
                        }

                    }
                    s = s.TrimStart(',').TrimEnd(',');
                    if (MessageBox.Show(s + "确定批量审核吗？", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        this.UpdateByProduct(s);
                }
            }
            try
            {
                if (appClsExcel != null)
                {
                    appClsExcel.Workbooks.Close();
                    appClsExcel.Application.Quit();
                    appClsExcel.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(appClsExcel);
                    appClsExcel = null;
                }
            }
            catch
            {

                Process[] processes = Process.GetProcessesByName("EXCEL");
                foreach (Process process in processes)
                {
                    if (process.MainWindowTitle.ToString().Trim() == "")
                    {
                        process.Kill();
                    }
                }
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
            string name = "IQC首批维护信息";
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
        private void btnPrint_Click(object sender, EventArgs e)
        {
            DataTable dt = databind.DataSource as DataTable;
            if (dt == null)
                return;
            if (dt.Rows.Count <= 0) return;

           //  DataToExcel(databind);

            string fileName = ShowSaveFileDialog("Microsoft Excel 2007 Document", "Microsoft Excel|*.xlsx");
            if (fileName == string.Empty) return;
            ExportToEx(fileName, "xlsx", gridView);
            OpenFile(fileName);
        }

        private void btnUnAuditing_Click(object sender, EventArgs e)
        {
            string sql = "select distinct s.TestType,Productcode,max(Productname) Productname,isnull(min(states),'未审') states,max(Receiver) AuditingByUser,max(AuditingDate) AuditingDate ,max(SendState) SendState,max(SendDate) SendDate,max(SendUser) SendUser,Max(RemarksSQE) Remarks from  IQC_TestProgSet s ";
            sql += " inner join  IQC_TestType i on s.TestType=i.TestType  where SendState='OK' and IFYiQi<>'否' and (AuditingByUser is null or states='未审')  and isnull(Receiver," + "'" + txtreceiver.Text + "')= case " + "'" + txtreceiver.Text + "' when '' then isnull(Receiver," + "'" + txtreceiver.Text + "') else '" + txtreceiver.Text + "' end and Productcode like " + "'" + txtproductcode.Text + "%' group by s.TestType,Productcode having isnull(min(states),'未审')='未审'";

            DataSet ds = DbAccess.SelectBySql(sql);
            gridView.Columns.Clear();
            databind.DataSource = null;

            databind.DataSource = ds.Tables[0];
        }

        private void cbok_CheckedChanged(object sender, EventArgs e)
        {
            if (cbok.Checked)
            {
                cbok.Text = "OK";
                txtremarks.Enabled = false;
                txtremarks.Text = "";
                cbNG.Checked = false;
            }
        }

        private void cbNG_CheckedChanged(object sender, EventArgs e)
        {
            if (cbNG.Checked)
            {
                cbNG.Text = "NG";
                this.txtremarks.Enabled = true;
                cbok.Checked = false;
            }
        }

        private void databind_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                //txtproductcode.Text = databind.CurrentRow.Cells["产品编码"].Value.ToString();
                //txtIQCChecktype.Text = databind.CurrentRow.Cells["检验方式"].Value.ToString();
                //cbTestType.SelectedValue = databind.CurrentRow.Cells["测试类别"].Value.ToString();
                //txttestitem.SelectedValue = databind.CurrentRow.Cells["测试项目"].Value.ToString();
                //txttestsubitem.SelectedValue = databind.CurrentRow.Cells["测试子项目"].Value.ToString();
                //txttestdes.Text = databind.CurrentRow.Cells["测试内容"].Value.ToString();
                //txttesttools.SelectedValue = databind.CurrentRow.Cells["测试工具"].Value.ToString();
                //txtpacktype.Text = databind.CurrentRow.Cells["包装方式"].Value.ToString();
                //lblinfo.Text = databind.CurrentRow.Cells["产品名称"].Value.ToString();
                //txtsampletype.Text = databind.CurrentRow.Cells["抽样方式"].Value.ToString();
                //txttestvalueqty.Text = databind.CurrentRow.Cells["测试值个数"].Value.ToString();
                //txtunit.SelectedValue = databind.CurrentRow.Cells["测试单位"].Value.ToString();
                //txtcheckcycle.Value = int.Parse(databind.CurrentRow.Cells["测试周期"].Value.ToString() == "" ? "0" : databind.CurrentRow.Cells["测试周期"].Value.ToString());
                //txtleadtime.Value = int.Parse(databind.CurrentRow.Cells["提前期"].Value.ToString() == "" ? "0" : databind.CurrentRow.Cells["提前期"].Value.ToString());
                //txtLowervalue.Text = databind.CurrentRow.Cells["下限"].Value.ToString();
                //txtUppervalue.Text = databind.CurrentRow.Cells["上限"].Value.ToString();
                //txtAQL.Text = databind.CurrentRow.Cells["AQL"].Value.ToString();

                txtproductcode.Text = gridView.GetFocusedRowCellValue("产品编码").ToString();
                txtIQCChecktype.Text = gridView.GetFocusedRowCellValue("检验方式").ToString();
                cbTestType.Text = gridView.GetFocusedRowCellValue("测试类别").ToString();
                txttestitem.Text = gridView.GetFocusedRowCellValue("测试项目").ToString();
                txttestsubitem.Text = gridView.GetFocusedRowCellValue("测试子项目").ToString();
                txttestdes.Text = gridView.GetFocusedRowCellValue("测试内容").ToString();
                txttesttools.Text = gridView.GetFocusedRowCellValue("测试工具").ToString();
                txtpacktype.Text = gridView.GetFocusedRowCellValue("包装方式").ToString();
                lblinfo.Text = gridView.GetFocusedRowCellValue("产品名称").ToString();
                txtsampletype.Text = gridView.GetFocusedRowCellValue("抽样方式").ToString();
                txttestvalueqty.Text = gridView.GetFocusedRowCellValue("测试值个数").ToString();
                ////txtunit.SelectedValue = gridView.GetFocusedRowCellValue("测试单位").ToString();

                txtunit.Text = bindTestUnitname(gridView.GetFocusedRowCellValue("测试单位").ToString());

                txtcheckcycle.Value = int.Parse(gridView.GetFocusedRowCellValue("测试周期").ToString() == "" ? "0" : gridView.GetFocusedRowCellValue("测试周期").ToString());
                txtleadtime.Value = int.Parse(gridView.GetFocusedRowCellValue("提前期").ToString() == "" ? "0" : gridView.GetFocusedRowCellValue("提前期").ToString());
                txtLowervalue.Text = gridView.GetFocusedRowCellValue("下限").ToString();;
                txtUppervalue.Text = gridView.GetFocusedRowCellValue("上限").ToString();
                txtAQL.Text = gridView.GetFocusedRowCellValue("AQL").ToString();

            }
            catch
            {
            }
            try
            {
                //txtproductcode.Text = databind.CurrentRow.Cells["Productcode"].Value.ToString();
                //cbTestType.SelectedValue = databind.CurrentRow.Cells["TestType"].Value.ToString();
                txtproductcode.Text = gridView.GetFocusedRowCellValue("Productcode").ToString();
                cbTestType.Text = gridView.GetFocusedRowCellValue("TestType").ToString();
                string F = IfCheck(cbTestType.Text, txtproductcode.Text);
                if (F == "已审核")
                {
                    this.lblcheckstate.Text = "已审核";
                    this.lblcheckstate.ForeColor = Color.Blue;
                    this.btnOK.Enabled = false;
                }
                else if (F == "未审")
                {
                    this.lblcheckstate.Text = "未审核";
                    this.lblcheckstate.ForeColor = Color.Red;
                    this.btnOK.Enabled = true;
                }
                else if (F == "NG")
                {
                    this.lblcheckstate.Text = "NG";
                    this.lblcheckstate.ForeColor = Color.Red;
                    this.btnOK.Enabled = true;
                }
                else
                {
                    this.lblcheckstate.Text = "不需审核";
                    this.lblcheckstate.ForeColor = Color.Green;
                    btnOK.Enabled = false;
                    btnAudit.Enabled = false;
                }
            }
            catch { }
        }

        private void databind_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void databind_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            //DataGridViewRow dgr = databind.Rows[e.RowIndex];
            //try
            //{
            //    if (dgr.Cells["states"].Value.ToString() == "NG")
            //    {
            //        dgr.DefaultCellStyle.BackColor = Color.Yellow;
            //    }

            //    else if (dgr.Cells["states"].Value.ToString() == "已审核")
            //    {
            //        dgr.DefaultCellStyle.BackColor = Color.Gold;
            //    }

            //}
            //catch
            //{
            //}
        }




        private void TestProgSet_Load(object sender, EventArgs e)
        {
            bindDeviceType();
            bindTestTools();
            bindTestAQL();
            bindsampleType();
            bindTestUnit();
            txtIQCChecktype.SelectedIndex = 1;
            txtAQL.SelectedIndex = 1;
            BindReceiver("SQE审核");
            setRule();
            gridView.PrintExportProgress += new ProgressChangedEventHandler(gridView_PrintExportProgress);
            txtproductcode.Focus();
        }

        private void txtproductcode_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13 && txtproductcode.Text != "")
            {

                txtproductcode_Leave(sender,e);
                //this.lblcheckstate.Text = "";
                //gridView.Columns.Clear();
                //databind.DataSource = null;
                //if (txtproductcode.Text == "") return;
                //string Orasql = "select organization_id organization,segment1 materialcode,description materialname,en_des materialenname,INVENTORY_ITEM_ID,PRIMARY_UOM_CODE from cux_item_material_v where segment1='" + txtproductcode.Text.Trim() + "'";
                //DataSet ds = DbAccess.SelectByOracle(Orasql);
                //if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                //{
                //    lblinfo.Text = ds.Tables[0].Rows[0]["materialname"].ToString();
                //    lblinfo.ForeColor = Color.Blue;
                //    //绑定已维护过的数据
                //    DataSet dsprog = ic.SelectTestProgRecord("查询", txtproductcode.Text, "", "", "", lblinfo.Text, "", "", txttestdes.Text, "", "", 0, 0, 0, txtAQL.Text, (int)numericud.Value, 0, "", 0, 0);
                //    if (dsprog != null && dsprog.Tables.Count > 0 && dsprog.Tables[0].Rows.Count > 0)
                //    {
                //        gridView.Columns.Clear();
                //        databind.DataSource = dsprog.Tables[0];
                //        cbTestType.Text = dsprog.Tables[0].Rows[0]["测试类别"].ToString();

                //        string F = IfCheck(cbTestType.Text, txtproductcode.Text);
                //        if (F == "已审核")
                //        {
                //            this.lblcheckstate.Text = "已审核";
                //            this.lblcheckstate.ForeColor = Color.Blue;
                //            this.btnOK.Enabled = false;
                //        }
                //        else if (F == "未审")
                //        {
                //            this.lblcheckstate.Text = "未审核";
                //            this.lblcheckstate.ForeColor = Color.Red;
                //            this.btnOK.Enabled = true;
                //        }
                //        else if (F == "NG")
                //        {
                //            this.lblcheckstate.Text = "NG";
                //            this.lblcheckstate.ForeColor = Color.Red;
                //            this.btnOK.Enabled = true;
                //        }
                //        else
                //        {
                //            this.lblcheckstate.Text = "不需审核";
                //            this.lblcheckstate.ForeColor = Color.Green;
                //            btnOK.Enabled = false;
                //            btnAudit.Enabled = false;
                //        }
                //    }
                //    else
                //        cbTestType.SelectedIndex = 0;
                //    cbTestType.Focus();
                //    bindTestItem(cbTestType.Text);
                //}
                //else
                //{
                //    string ssql = "select materialcode from delivery where lotno='" + txtproductcode.Text + "'";
                //    DataSet dslotno = DbAccess.SelectBySql(ssql);
                //    if (dslotno != null && ds.Tables.Count > 0 && dslotno.Tables[0].Rows.Count > 0)
                //    {
                //        string materialcode = dslotno.Tables[0].Rows[0]["materialcode"].ToString();
                //        string Orasqlbylotno = "select organization_id organization,segment1 materialcode,description materialname,en_des materialenname,INVENTORY_ITEM_ID,PRIMARY_UOM_CODE from cux_item_material_v where segment1='" + materialcode + "'";
                //        ds = DbAccess.SelectByOracle(Orasqlbylotno);
                //        if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                //        {
                //            txtproductcode.Text = materialcode;
                //            lblinfo.Text = ds.Tables[0].Rows[0]["materialname"].ToString();
                //            lblinfo.ForeColor = Color.Blue;
                //            //绑定已维护过的数据
                //            DataSet dsprog = ic.SelectTestProgRecord("查询", txtproductcode.Text, "", "", "", lblinfo.Text, "", "", txttestdes.Text, "", "", 0, 0, 0, txtAQL.Text, (int)numericud.Value, 0, "", 0, 0);
                //            if (dsprog != null && dsprog.Tables.Count > 0 && dsprog.Tables[0].Rows.Count > 0)
                //            {
                //                gridView.Columns.Clear();
                //                databind.DataSource = dsprog.Tables[0];
                //                cbTestType.Text = dsprog.Tables[0].Rows[0]["测试类别"].ToString();

                //                string F = IfCheck(cbTestType.Text, txtproductcode.Text);
                //                if (F == "已审核")
                //                {
                //                    this.lblcheckstate.Text = "已审核";
                //                    this.lblcheckstate.ForeColor = Color.Blue;
                //                    this.btnOK.Enabled = false;
                //                }
                //                else if (F == "未审")
                //                {
                //                    this.lblcheckstate.Text = "未审核";
                //                    this.lblcheckstate.ForeColor = Color.Red;
                //                    this.btnOK.Enabled = true;
                //                }
                //                else
                //                {
                //                    this.lblcheckstate.Text = "不需审核";
                //                    this.lblcheckstate.ForeColor = Color.Green;
                //                    btnOK.Enabled = false;
                //                    btnAudit.Enabled = false;
                //                }
                //            }
                //            else
                //                cbTestType.SelectedIndex = 0;
                //            cbTestType.Focus();
                //            bindTestItem(cbTestType.Text);
                //        }
                //    }
                //    else
                //    {
                //        lblinfo.Text = txtproductcode.Text + "该料号不存在";
                //        lblinfo.ForeColor = Color.Red;
                //        txtproductcode.Text = "";
                //        txtproductcode.Focus();
                //    }
                //}
            }
        }

        private void gridView_Click(object sender, EventArgs e)
        {
            try
            {
                txtproductcode.Text = gridView.GetFocusedRowCellValue("产品编码").ToString();
                txtIQCChecktype.Text = gridView.GetFocusedRowCellValue("检验方式").ToString();
                cbTestType.Text = gridView.GetFocusedRowCellValue("测试类别").ToString();
                txttestitem.Text = gridView.GetFocusedRowCellValue("测试项目").ToString();
                txttestsubitem.Text = gridView.GetFocusedRowCellValue("测试子项目").ToString();
                txttestdes.Text = gridView.GetFocusedRowCellValue("测试内容").ToString();
                txttesttools.Text = gridView.GetFocusedRowCellValue("测试工具").ToString();
                txtpacktype.Text = gridView.GetFocusedRowCellValue("包装方式").ToString();
                lblinfo.Text = gridView.GetFocusedRowCellValue("产品名称").ToString();
                txtsampletype.Text = gridView.GetFocusedRowCellValue("抽样方式").ToString();
                txttestvalueqty.Text = gridView.GetFocusedRowCellValue("测试值个数").ToString();
                // txtunit.SelectedValue = gridView.GetFocusedRowCellValue("测试单位").ToString();

                txtunit.Text = bindTestUnitname(gridView.GetFocusedRowCellValue("测试单位").ToString());

                txtcheckcycle.Value = int.Parse(gridView.GetFocusedRowCellValue("测试周期").ToString() == "" ? "0" : gridView.GetFocusedRowCellValue("测试周期").ToString());
                txtleadtime.Value = int.Parse(gridView.GetFocusedRowCellValue("提前期").ToString() == "" ? "0" : gridView.GetFocusedRowCellValue("提前期").ToString());
                txtLowervalue.Text = gridView.GetFocusedRowCellValue("下限").ToString(); ;
                txtUppervalue.Text = gridView.GetFocusedRowCellValue("上限").ToString();
                txtAQL.Text = gridView.GetFocusedRowCellValue("AQL").ToString();

            }
            catch
            {
            }
            try
            {
                txtproductcode.Text = gridView.GetFocusedRowCellValue("Productcode").ToString();
                cbTestType.Text = gridView.GetFocusedRowCellValue("TestType").ToString();
                string F = IfCheck(cbTestType.Text, txtproductcode.Text);
                if (F == "已审核")
                {
                    this.lblcheckstate.Text = "已审核";
                    this.lblcheckstate.ForeColor = Color.Blue;
                    this.btnOK.Enabled = false;
                }
                else if (F == "未审")
                {
                    this.lblcheckstate.Text = "未审核";
                    this.lblcheckstate.ForeColor = Color.Red;
                    this.btnOK.Enabled = true;
                }
                else if (F == "NG")
                {
                    this.lblcheckstate.Text = "NG";
                    this.lblcheckstate.ForeColor = Color.Red;
                    this.btnOK.Enabled = true;
                }
                else
                {
                    this.lblcheckstate.Text = "不需审核";
                    this.lblcheckstate.ForeColor = Color.Green;
                    btnOK.Enabled = false;
                    btnAudit.Enabled = false;
                }
            }
            catch { }
        }

        private void gridView_CustomDrawCell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
        {

            if (e.Column.FieldName == "states")
            {
                GridCellInfo GridCell = e.Cell as GridCellInfo;
                if (GridCell.CellValue.ToString() == "NG")
                {
                    e.Appearance.BackColor = Color.Yellow;
                }
                else if (GridCell.CellValue.ToString() == "已审核")
                {
                    e.Appearance.BackColor = Color.Gold;
                }
            }


            //DataGridViewRow dgr = databind.Rows[e.RowIndex];
            //try
            //{
            //    if (dgr.Cells["states"].Value.ToString() == "NG")
            //    {
            //        dgr.DefaultCellStyle.BackColor = Color.Yellow;
            //    }

            //    else if (dgr.Cells["states"].Value.ToString() == "已审核")
            //    {
            //        dgr.DefaultCellStyle.BackColor = Color.Gold;
            //    }

            //}
            //catch
            //{
            //}

        }

        private void TestProgSet_FormClosing(object sender, FormClosingEventArgs e)
        {
            Process[] myProcesses;
            DateTime startTime;
            myProcesses = Process.GetProcessesByName("Excel");
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