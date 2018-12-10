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

namespace DX_QMS.IQCFilePosition
{
    public partial class IQC_RepeatTest : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private Microsoft.Office.Interop.Excel.ApplicationClass appClsExcel = null;
        DataTable dtsubitem = null;
         IQC ic = new IQC();
        public IQC_RepeatTest()
        {
            InitializeComponent();     
        }
        private void bindDeviceType()
        {
            DataTable dt = ic.SelectTestTypeRecord("查询", "", "测试类别", "").Tables[0];           
            cbTestType.Properties.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
                cbTestType.Properties.Items.Add(row["TestType"]);
            }
        }
        private void bindTestTools()
        {
            DataTable dt = ic.SelectTestTypeRecord("查询", "", "测试工具", "").Tables[0];
            txttesttools.Properties.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
                txttesttools.Properties.Items.Add(row["TestType"]);
            }
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
            txtsampletype.Properties.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
                txtsampletype.Properties.Items.Add(row["SampleType"]);
            }
        }
        string bindTestUnitvalue(string unitname)
        {
            string unitvalue = "";
            string sql = @"  select u.unit  from ( select unit, (unit+unitname) name from IQC_TestUnit ) u  where u.name = '"+ unitname+ "'  ";
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
            txtunit.Properties.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
                txtunit.Properties.Items.Add(row["name"]);
            }
        }
        private void BindReceiver(string sType)
        {
            string sql = "select userName,userMail from MailGroup where mailType='" + sType + "'";
            DataSet ds = DbAccess.SelectBySql(sql);
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
        }

        public DateTime beforeTime, afterTime;
        private string IfCheck(string sType, string productcode)
        {
            string Flag = "不需审核";
            string sql = "select min(case when IFYiQi='否' then '不需审核' else ISNULL(states,'未审') end) states from  IQC_RepeatTestPro s inner join  IQC_TestType i on s.TestType=i.TestType where Productcode='" + productcode + "' and s.TestType='" + sType + "'";
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
                DataSet dsprog = ic.SelectRepeatTestProg("查询", txtproductcode.Text, "", "", "", lblinfo.Text, "", "", txttestdes.Text, "", "", 0, 0, 0, txtAQL.Text, (int)numericud.Value, 0, "", 0, 0);
                if (dsprog != null && dsprog.Tables.Count > 0 && dsprog.Tables[0].Rows.Count > 0)
                {
                    databind.DataSource = dsprog.Tables[0];
                    cbTestType.Text = dsprog.Tables[0].Rows[0]["测试类别"].ToString();              
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
                        DataSet dsprog = ic.SelectRepeatTestProg("查询", txtproductcode.Text, "", "", "", lblinfo.Text, "", "", txttestdes.Text, "", "", 0, 0, 0, txtAQL.Text, (int)numericud.Value, 0, "", 0, 0);
                        if (dsprog != null && dsprog.Tables.Count > 0 && dsprog.Tables[0].Rows.Count > 0)
                        {
                            databind.DataSource = dsprog.Tables[0];
                            cbTestType.Text = dsprog.Tables[0].Rows[0]["测试类别"].ToString();                     
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
            DataTable dt = ic.SelectRepeatTestProg("查询", txtproductcode.Text.Trim(),"","","", lblinfo.Text,"", "", txttestdes.Text,"", "", 0, 0, 0, txtAQL.Text, (int)numericud.Value, 0, "", 0, 0).Tables[0];
            return dt;
        }
        private void btnsave_Click(object sender, EventArgs e)
        {
            if (txtproductcode.Text == "")
                return;
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
            if (txtproductcode.Text.Trim() == "" || txttestitem.Text.Trim() == "" || txttestsubitem.Text.Trim() == "")
            {
                MessageBox.Show("请确认物料编码、测试项目和测试子项目不为空！", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (txtRepeatTestType.Text.Trim()== "")
            {
                MessageBox.Show("请选择重检类型", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            string unit = bindTestUnitvalue(txtunit.Text);

            try
            {
                int i = ic.AddRepeatTestProg("新增", txtproductcode.Text, txtRepeatTestType.Text, cbTestType.Text, txtpacktype.Text, lblinfo.Text, txttestitem.Text, txttestsubitem.Text, txttestdes.Text,
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

            if (gridView.Columns.Count != 20)
                return;

            for (int k = gridView.GetSelectedRows().Length; k > 0; k--)
            {
                DataRow dr = gridView.GetDataRow(gridView.GetSelectedRows()[k-1]);
                int i = ic.AddRepeatTestProg("删除", dr["产品编码"].ToString(), dr["重检类型"].ToString(), dr["测试类别"].ToString(), dr["包装方式"].ToString(), "",
                                                dr["测试项目"].ToString(), dr["测试子项目"].ToString(), "", "", "", "0", "0", 0, "", 0, 0, "", 0, 0);
            }
            gridView.Columns.Clear();
            databind.DataSource = ic.SelectRepeatTestProg("查询", txtproductcode.Text, txtRepeatTestType.Text, cbTestType.Text, txtpacktype.Text, lblinfo.Text, txttestitem.Text, txttestsubitem.Text, txttestdes.Text,
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
            string templetFile = Environment.CurrentDirectory + @"\ReportFolder\MaterialExpiryDate.xls";

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
                sheet.Cells[2, 1] = "IQC重检报表(" + dt.Rows[0 + j]["TestType"].ToString() + ")";
                sheet.Cells[5, 2] = dt.Rows[0 + j]["Productcode"].ToString();
                sheet.Cells[5, 7] = dt.Rows[0 + j]["materialenname"].ToString();

                sheet.Cells[10 + j + 1, 1] = dt.Rows[0 + j]["TestItem"].ToString();
                sheet.Cells[10 + j + 1, 2] = dt.Rows[0 + j]["TestSubItem"].ToString();
                sheet.Cells[10 + j + 1, 3] = dt.Rows[0 + j]["TestDesc"].ToString();
                sheet.Cells[10 + j + 1, 9] = dt.Rows[0 + j]["TestTool"].ToString();
                sheet.Cells[10 + j + 1, 11] = dt.Rows[0 + j]["SampleType"].ToString();

                if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("MI"))
                {
                    sheet.Cells[10 + j + 1, 23] = "√";
                 
                }
                else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("MA"))
                {
                    sheet.Cells[10 + j + 1, 21] = "√";

                }
                else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("CR=0.1"))
                {
                    sheet.Cells[10 + j + 1, 19] = "√";

                }
                else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=0.40"))
                {
                    sheet.Cells[10 + j + 1, 20] = "√";

                }
                else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=1.0"))
                {
                    sheet.Cells[10 + j + 1, 22] = "√";

                }
                else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("CR=0.01"))
                {
                    sheet.Cells[10 + j + 1, 18] = "√";

                }
                else if (tb.Rows[0 + j]["SampleType"].ToString().Contains("ISO2859-1"))
                {
                    if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=1.5"))
                    {
                        sheet.Cells[10 + j + 1, 23] = "√";         
                    }
                    else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=0.65"))
                    {
                        sheet.Cells[10 + j + 1, 21] = "√";
                    }
                    else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=0.10"))
                    {
                        sheet.Cells[10 + j + 1, 19] = "√";
                    }
                    else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=0.40"))
                    {
                        sheet.Cells[10 + j + 1, 20] = "√";
                    }
                    else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=1.0"))
                    {
                        sheet.Cells[10 + j + 1, 22] = "√";
                    }
                    else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=0.010"))
                    {
                        sheet.Cells[10 + j + 1, 18] = "√";
                    }
                }
                else if (tb.Rows[0 + j]["SampleType"].ToString().Contains("C=0"))
                {
                    if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=1.5"))
                    {
                        sheet.Cells[10 + j + 1, 23] = "√";
                    }
                    else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=0.65"))
                    {
                        sheet.Cells[10 + j + 1, 21] = "√";
                    }
                    else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=0.1"))
                    {
                        sheet.Cells[10 + j + 1, 19] = "√";
                    }
                    else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=0.4"))
                    {
                        sheet.Cells[10 + j + 1, 20] = "√";
                    }
                    else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=1"))
                    {
                        sheet.Cells[10 + j + 1, 22] = "√";
                    }
                    else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=0.01"))
                    {
                        sheet.Cells[10 + j + 1, 18] = "√";
                    }
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
            if (gridView.RowCount < 1)
                return;
            string ssql = @"SELECT Productcode, RepeatTestType, l.TestType, l.PackType, Productname materialenname, l.SampleType, l.TestItem, l.TestSubItem,case when (cast(l.LowValue as float)<>0 or cast(l.UpValue as float)<>0) then (l.TestDesc+'(范围:【'+cast(l.LowValue as varchar(10))+'】至:【'+cast(l.UpValue as varchar(10))+'】)') else l.TestDesc end TestDesc, l.TestTool, l.UpValue, l.LowValue, l.UpScope, 
                          l.AQLValue, l.AQL, l.Item, SubItem, Lastupdatedate, Testvalueqty, Unit, Checkcycle
                          FROM IQC_RepeatTestPro l
                          left join IQC_TestSubItemSet s on l.TestSubItem=s.TestSubItem and l.TestItem=s.TestItem and l.TestType=s.TestType
                          left join IQC_TestItem t on s.TestItem=t.TestItem and s.TestType=t.TestType where l.Productcode='" + txtproductcode.Text + "' order by t.Item asc,s.Item";
            DataTable dt = DbAccess.SelectBySql(ssql).Tables[0];
            if (dt.Rows.Count > 0)
            {
               Copy(dt.Rows[0]["TestType"].ToString(), dt, dt);
            }

        }
        private string ShowSaveFileDialog(string title, string filter)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            string name = "IQC重检维护信息";
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
            if (dt.Rows.Count <= 0)
                return;
            string fileName = ShowSaveFileDialog("Microsoft Excel 2007 Document", "Microsoft Excel|*.xlsx");
            if (fileName == string.Empty) return;
            ExportToEx(fileName, "xlsx", gridView);
            OpenFile(fileName);
        }
        private void databind_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                txtproductcode.Text = gridView.GetFocusedRowCellValue("产品编码").ToString();
                txtRepeatTestType.Text = gridView.GetFocusedRowCellValue("检验方式").ToString();
                cbTestType.Text = gridView.GetFocusedRowCellValue("测试类别").ToString();
                txttestitem.Text = gridView.GetFocusedRowCellValue("测试项目").ToString();
                txttestsubitem.Text = gridView.GetFocusedRowCellValue("测试子项目").ToString();
                txttestdes.Text = gridView.GetFocusedRowCellValue("测试内容").ToString();
                txttesttools.Text = gridView.GetFocusedRowCellValue("测试工具").ToString();
                txtpacktype.Text = gridView.GetFocusedRowCellValue("包装方式").ToString();
                lblinfo.Text = gridView.GetFocusedRowCellValue("产品名称").ToString();
                txtsampletype.Text = gridView.GetFocusedRowCellValue("抽样方式").ToString();
                txttestvalueqty.Text = gridView.GetFocusedRowCellValue("测试值个数").ToString();

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
                txtproductcode.Text = gridView.GetFocusedRowCellValue("Productcode").ToString();
                cbTestType.Text = gridView.GetFocusedRowCellValue("TestType").ToString();        
            }
            catch { }
        }

        private void databind_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void TestProgSet_Load(object sender, EventArgs e)
        {
            bindDeviceType();
            bindTestTools();
            bindTestAQL();
            bindsampleType();
            bindTestUnit();
            txtRepeatTestType.SelectedIndex = 1;
            txtAQL.SelectedIndex = 1;          
            setRule();
            gridView.PrintExportProgress += new ProgressChangedEventHandler(gridView_PrintExportProgress);
            txtRepeatTestType.SelectedIndex = -1;
            txtproductcode.Focus();
        }

        private void txtproductcode_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13 && txtproductcode.Text != "")
            {
                txtproductcode_Leave(sender,e);
            }
        }

        private void gridView_Click(object sender, EventArgs e)
        {

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

        }

        private void txtRepeatTestType_SelectedIndexChanged(object sender, EventArgs e)
        {           
                //DataTable dt = DbAccess.SelectBySql("select SampleType,Samplevalue from IQC_TestSampleType").Tables[0]; 
                //txtsampletype.Properties.Items.Clear();
                //foreach (DataRow row in dt.Rows)
                //{
                //    txtsampletype.Properties.Items.Add(row["SampleType"]);
                //}           
        }

        private void gridView_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {
            if (gridView.RowCount<1)
            {
                return;
            }
            try
            {
                txtproductcode.Text = gridView.GetFocusedRowCellValue("产品编码").ToString();
                txtRepeatTestType.Text = gridView.GetFocusedRowCellValue("重检类型").ToString();
                cbTestType.Text = gridView.GetFocusedRowCellValue("测试类别").ToString();
                txttestitem.Text = gridView.GetFocusedRowCellValue("测试项目").ToString();
                txttestsubitem.Text = gridView.GetFocusedRowCellValue("测试子项目").ToString();
                txttestdes.Text = gridView.GetFocusedRowCellValue("测试内容").ToString();
                txttesttools.Text = gridView.GetFocusedRowCellValue("测试工具").ToString();
                txtpacktype.Text = gridView.GetFocusedRowCellValue("包装方式").ToString();
                lblinfo.Text = gridView.GetFocusedRowCellValue("产品名称").ToString();
                txtsampletype.Text = gridView.GetFocusedRowCellValue("抽样方式").ToString();
                txttestvalueqty.Text = gridView.GetFocusedRowCellValue("测试值个数").ToString();

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
            }
            catch { }
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