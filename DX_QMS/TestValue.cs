using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DX_QMS.Common;
using DevExpress.XtraBars;

namespace DX_QMS
{
    public partial class TestValue : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        IQC ic = new IQC();
        DataTable dtsubitem = null;
        string StrSql = "";
        string StrSql_id = "";
        string StrSql_cl = "";
        string StrSql_sl = "";
        string StrSqlUpdate = "";
        DataSet ds = new DataSet();
        private DataTable dt = null;
        public TestValue()
        {
            InitializeComponent();
            setRule();
            bindDeviceType();
        }

        private void setRule()
        {
            string post = "";
            if (Login.manager != null)
            {
                post = Login.manager;
            }
            else
            {
                post = Login.post;
            }
            Dictionary<string, bool> dic = GroupPermission.QMS_SelectRulesForForm(post, "公差");
            this.btnsave.Enabled = bool.Parse(dic["hasInsert"].ToString());
            this.btndel.Enabled = bool.Parse(dic["hasDelete"].ToString());
            this.btnsearch.Enabled = bool.Parse(dic["hasQuery"].ToString());
        }



        private void bindDeviceType()
        {
            DataTable dt = ic.SelectTestTypeRecord("查询", "", "测试类别", "").Tables[0];
            cbTestType.DataSource = dt;
            cbTestType.DisplayMember = dt.Columns["TestType"].ToString();
            cbTestType.ValueMember = dt.Columns["TestType"].ToString();
        }
        private void TestValue_Load(object sender, EventArgs e)
        {
            cmbVersion.Text = "A";
        }
        private void bindTestItem(string TestType)
        {
            DataSet ds = ic.SelectTestItemRecord("查询", TestType, "", 0, "");
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                DataTable dt = ds.Tables[0];
                txttestitem.DisplayMember = dt.Columns["TestItem"].ToString();
                txttestitem.ValueMember = dt.Columns["TestItem"].ToString();
                txttestitem.DataSource = dt;
                txttestitem.SelectedIndex = 0;
            }
        }
        private void txtproductcode_Leave(object sender, EventArgs e)
        {
            if (txtproductcode.Text == "") return;
            txtproductcode.Leave -= txtproductcode_Leave;
            string Orasql = "select organization_id organization,segment1 materialcode,description materialname,en_des materialenname,INVENTORY_ITEM_ID,PRIMARY_UOM_CODE from cux_item_material_v where segment1='" + txtproductcode.Text.Trim() + "'";
            DataSet ds = DbAccess.SelectByOracle(Orasql);
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                cbTestType.SelectedIndex = 0;
                cbTestType.Focus();
                bindTestItem(cbTestType.SelectedValue.ToString());
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
                        cbTestType.SelectedIndex = 0;
                        cbTestType.Focus();
                        bindTestItem(cbTestType.SelectedValue.ToString());
                    }
                }
                else
                {
                    MessageBox.Show("该料号不存在", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtproductcode.Text = "";
                    txtproductcode.Focus();
                }
            }
            txtproductcode.Leave += txtproductcode_Leave;
        }

        private void cbTestType_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txttestsubitem.SelectedValue == null) return;
            DataRow[] rw = dtsubitem.Select("TestSubItem='" + this.txttestsubitem.SelectedValue.ToString() + "'");
            txtTestValue.Text = "";
            txtUpRange.Text = "";
            txtLowRange.Text = "";
        }
        private void bindTestsubItem(string testtype, string Testitem)
        {
            string ssql = "select  TestSubItem, Item, TestDesc, TestTool, PackType from IQC_TestSubItemSet where TestItem='" + Testitem + "' and TestType='" + testtype + "'";
            DataSet ds = DbAccess.SelectBySql(ssql);
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                DataTable dt = ds.Tables[0];
                dtsubitem = dt;

                txttestsubitem.DisplayMember = dt.Columns["TestSubItem"].ToString();
                txttestsubitem.ValueMember = dt.Columns["TestSubItem"].ToString();
                txttestsubitem.DataSource = dt;
                if (txttestsubitem.Items.Count == 1)
                {
                    txttestsubitem.SelectedIndex = 0;
                    DataRow[] rw = dtsubitem.Select("TestSubItem='" + this.txttestsubitem.SelectedValue.ToString() + "'");
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
            bindTestsubItem(cbTestType.SelectedValue.ToString(), txttestitem.SelectedValue == null ? "" : txttestitem.SelectedValue.ToString());
        }

        private void txttestsubitem_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txttestsubitem.SelectedValue == null) return;
            DataRow[] rw = dtsubitem.Select("TestSubItem='" + this.txttestsubitem.SelectedValue.ToString() + "'");
            txtTestValue.Text = "";
            txtUpRange.Text = "";
            txtLowRange.Text = "";
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            txtproductcode.Text = "";
            cbTestType.Text = "";
            txttestitem.Text = "";
            txttestsubitem.Text = "";
            txtTestValue.Text = "";
            txtUpRange.Text = "";
            txtLowRange.Text = "";
            cmbVersion.Text = "";
        }

        private void btnsave_Click(object sender, EventArgs e)
        {
            if (txtproductcode.Text == "")
            {
                MessageBox.Show("请输入物料编码", "系统提醒", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtproductcode.Focus();
                return;
            }
            if (txtTestValue.Text == "")
            {
                MessageBox.Show("请输入测试标准值", "系统提醒", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtTestValue.Focus();
                return;
            }

            StrSql_id = @"select COUNT(1) from SPC..C_SPC_CONTROL_ITEM_ID_NEW WHERE  CONTROL_ITEM_UNICODE='" + txtproductcode.Text.ToString().Trim() + "' and TestSubItem='" + txttestsubitem.Text.ToString() + "'";
            StrSql_sl = @"select COUNT(1) from SPC..C_SPC_CONTROL_ITEM_SL_NEW WHERE  CONTROL_ITEM_UNICODE='" + txtproductcode.Text.ToString().Trim() + "' and TestSubItem='" + txttestsubitem.Text.ToString() + "'";
            StrSql_cl = @"select COUNT(1) from SPC..C_SPC_CONTROL_ITEM_CL_NEW WHERE  CONTROL_ITEM_UNICODE='" + txtproductcode.Text.ToString().Trim() + "' and TestSubItem='" + txttestsubitem.Text.ToString() + "'";
            //ds = DbAccess.SelectBySql(StrSql);
            if (DbAccess.Exists(StrSql_id) == false)
            {
                StrSql_id = @"insert into SPC..C_SPC_CONTROL_ITEM_ID_NEW(CONTROL_ITEM_UNICODE,ITEM_FLAG,TestType,TestItem,TestSubItem,MODE_UNICODE,TestValue,Uplimit,Downlimit,Version,EDIT_EMP,EDIT_TIME) 
                              values ('" + txtproductcode.Text.ToString() + "','1','" + cbTestType.Text.ToString() + "','" + txttestitem.Text.Trim() + "','" + txttestsubitem.Text.ToString()
                                   + "','1','" + txtTestValue.Text.Trim() + "','" + txtUpRange.Text.Trim() + "','" + txtLowRange.Text.Trim() + "','"
                                   + cmbVersion.Text.ToString() + "','" + Login.userId + "',getdate())";


                if (!string.IsNullOrEmpty(StrSql_id))
                {
                    if (DbAccess.ExecuteSql(StrSql_id))
                    {
                        MessageBox.Show("测试数据信息保存成功");
                    }
                    else
                    {
                        MessageBox.Show("测试数据信息保存失败");
                    }
                }
            }
            else
            {
                StrSql = @"update SPC..C_SPC_CONTROL_ITEM_ID_NEW set CONTROL_ITEM_UNICODE='" + txtproductcode.Text.ToString() + "',TestType='"
                           + cbTestType.Text.ToString() + "',TestItem='" + txttestitem.Text.Trim() + "',TestSubItem='"
                           + txttestsubitem.Text.ToString() + "',TestValue='" + txtTestValue.Text.Trim() + "',Uplimit='"
                           + txtUpRange.Text.Trim() + "',Downlimit='" + txtLowRange.Text.Trim() + "',Version='"
                           + cmbVersion.Text.ToString() + "',EDIT_EMP='" + Login.userId + "',EDIT_TIME=getdate()";
                if (!string.IsNullOrEmpty(StrSql))
                {
                    if (DbAccess.ExecuteSql(StrSql))
                    {
                        MessageBox.Show("测试数据信息更新成功");
                    }
                    else
                    {
                        MessageBox.Show("测试数据信息更新失败");
                    }
                }
            }
            if (DbAccess.Exists(StrSql_sl) == false)
            {
                StrSql_sl = @"insert into SPC..C_SPC_CONTROL_ITEM_SL_NEW(CONTROL_ITEM_UNICODE,TestType,TestItem,TestSubItem,USL,TARGET,LSL,CHART_FLAG,VERSION,START_TIME,END_TIME,EDIT_EMP,EDIT_TIME) 
                              values  ('" + txtproductcode.Text.ToString() + "','" + cbTestType.Text.ToString() + "','" + txttestitem.Text.Trim() + "','" + txttestsubitem.Text.ToString()
                                          + "','" + (float.Parse(txtTestValue.Text.Trim()) + float.Parse(txtUpRange.Text.Trim())) + "','"
                                          + txtTestValue.Text.Trim() + "','" + (float.Parse(txtTestValue.Text.Trim()) - float.Parse(txtLowRange.Text.Trim())) + "','0',convert(varchar(10),GETDATE(),120),'','','"
                                          + Login.userId + "',getdate())";


                if (!string.IsNullOrEmpty(StrSql_sl))
                {
                    if (DbAccess.ExecuteSql(StrSql_sl))
                    {
                        MessageBox.Show("规格数据保存成功");
                    }
                    else
                    {
                        MessageBox.Show("规格数据保存失败");
                    }
                }
            }
            else
            {
                StrSql = @"update SPC..C_SPC_CONTROL_ITEM_ID_NEW set CONTROL_ITEM_UNICODE='" + txtproductcode.Text.ToString() + "',TestType='"
                           + cbTestType.Text.ToString() + "',TestItem='" + txttestitem.Text.Trim() + "',TestSubItem='"
                           + txttestsubitem.Text.ToString() + "',TARGET='" + txtTestValue.Text.Trim() + "',USL='"
                           + (float.Parse(txtTestValue.Text.Trim()) + float.Parse(txtUpRange.Text.Trim())) + "',LSL='"
                           + (float.Parse(txtTestValue.Text.Trim()) - float.Parse(txtLowRange.Text.Trim())) + "',Version=convert(varchar(10),GETDATE(),120),EDIT_EMP='"
                           + Login.userId + "',EDIT_TIME=getdate()";
                if (!string.IsNullOrEmpty(StrSql))
                {
                    if (DbAccess.ExecuteSql(StrSql))
                    {
                        MessageBox.Show("规格数据更新成功");
                    }
                    else
                    {
                        MessageBox.Show("规格数据更新失败");
                    }
                }
            }
            if (DbAccess.Exists(StrSql_cl) == false)
            {
                StrSql_cl = @"insert into SPC..C_SPC_CONTROL_ITEM_CL_NEW(CONTROL_ITEM_UNICODE,TestType,TestItem,TestSubItem,UCL,CL,LCL,XUCL,XCL,XLCL,VERSION,ITEM_DATA_TYPE,CONTROL_CHART,SAMPLE_SIZE,EDIT_EMP,EDIT_TIME) 
                              values ('" + txtproductcode.Text.ToString() + "','" + cbTestType.Text.ToString() + "','" + txttestitem.Text.Trim() + "','" + txttestsubitem.Text.ToString()
                                          + "','" + Math.Round((float.Parse(txtTestValue.Text.Trim()) + float.Parse(txtUpRange.Text.Trim()) * 0.8), 2) + "','"
                                          + txtTestValue.Text.Trim() + "','" + Math.Round((float.Parse(txtTestValue.Text.Trim()) - float.Parse(txtLowRange.Text.Trim()) * 0.8), 2) + "','0.1','0.05','0',convert(varchar(10),GETDATE(),120),'0','101','5','"
                                          + Login.userId + "',getdate())";


                if (!string.IsNullOrEmpty(StrSql_cl))
                {
                    if (DbAccess.ExecuteSql(StrSql_cl))
                    {
                        MessageBox.Show("管制数据保存成功");
                    }
                    else
                    {
                        MessageBox.Show("管制数据保存失败");
                    }
                }
            }
            else
            {
                StrSql = @"update SPC..C_SPC_CONTROL_ITEM_CL_NEW set CONTROL_ITEM_UNICODE='" + txtproductcode.Text.ToString() + "',TestType='"
                           + cbTestType.Text.ToString() + "',TestItem='" + txttestitem.Text.Trim() + "',TestSubItem='"
                           + txttestsubitem.Text.ToString() + "',CL='" + txtTestValue.Text.Trim() + "',UCL='"
                           + Math.Round((float.Parse(txtTestValue.Text.Trim()) + float.Parse(txtUpRange.Text.Trim()) * 0.8), 2) + "',LCL='"
                           + Math.Round((float.Parse(txtTestValue.Text.Trim()) - float.Parse(txtLowRange.Text.Trim()) * 0.8), 2) + "',Version=convert(varchar(10),GETDATE(),120),EDIT_EMP='"
                           + Login.userId + "',EDIT_TIME=getdate()";
                if (!string.IsNullOrEmpty(StrSql))
                {
                    if (DbAccess.ExecuteSql(StrSql))
                    {
                        MessageBox.Show("管制数据更新成功");
                    }
                    else
                    {
                        MessageBox.Show("管制数据更新失败");
                    }
                }
            }
        }

        private void btndel_Click(object sender, EventArgs e)
        {
            if (txtproductcode.Text.Trim() == "")
            {
                MessageBox.Show("物料编码为空,请确认");
                return;
            }
            StrSql_id = @"delete from SPC..C_SPC_CONTROL_ITEM_ID_NEW where CONTROL_ITEM_UNICODE='" + txtproductcode.Text.ToString().Trim() + "' and TestType='"
                       + cbTestType.Text.ToString() + "' and TestItem='" + txttestitem.Text.Trim() + "' and TestSubItem='"
                       + txttestsubitem.Text.ToString() + "' and Version='" + cmbVersion.Text.ToString() + "'";
            if (!string.IsNullOrEmpty(StrSql_id))
            {
                if (DbAccess.ExecuteSql(StrSql_id))
                {
                    MessageBox.Show("测试数据信息删除成功");
                }
                else
                {
                    MessageBox.Show("测试数据信息删除失败");
                }
            }
            StrSql_sl = @"delete from SPC..C_SPC_CONTROL_ITEM_SL_NEW where CONTROL_ITEM_UNICODE='" + txtproductcode.Text.ToString().Trim() + "' and TestType='"
           + cbTestType.Text.ToString() + "' and TestItem='" + txttestitem.Text.Trim() + "' and TestSubItem='"
           + txttestsubitem.Text.ToString() + "'";
            if (!string.IsNullOrEmpty(StrSql_sl))
            {
                if (DbAccess.ExecuteSql(StrSql_sl))
                {
                    MessageBox.Show("规格数据删除成功");
                }
                else
                {
                    MessageBox.Show("规格数据删除失败");
                }
            }
            StrSql_cl = @"delete from SPC..C_SPC_CONTROL_ITEM_CL_NEW where CONTROL_ITEM_UNICODE='" + txtproductcode.Text.ToString().Trim() + "' and TestType='"
           + cbTestType.Text.ToString() + "' and TestItem='" + txttestitem.Text.Trim() + "' and TestSubItem='"
           + txttestsubitem.Text.ToString() + "'";
            if (!string.IsNullOrEmpty(StrSql_cl))
            {
                if (DbAccess.ExecuteSql(StrSql_cl))
                {
                    MessageBox.Show("管制数据删除成功");
                }
                else
                {
                    MessageBox.Show("管制数据删除失败");
                }
            }
        }
        public DataSet dsSelect()
        {
            //this.databind.Rows.Clear();
            databind.DataSource = null;
            string StrWhere = "1=1";
            if (txtproductcode.Text.Trim() != "")
            {
                StrWhere += " and id.CONTROL_ITEM_UNICODE like '" + "%" + txtproductcode.Text.Trim() + "%" + "'";
            }
            StrSql = @"   select id.CONTROL_ITEM_UNICODE,id.TestType,id.TestItem,id.TestSubItem,iqc.UpValue,iqc.LowValue,
                          id.TestValue TestStandardValue,id.Uplimit,id.Downlimit,sl.USL,sl.LSL,id.Version fileversion,id.EDIT_EMP,id.EDIT_TIME 
                          from SPC..C_SPC_CONTROL_ITEM_ID_NEW id,SPC..C_SPC_CONTROL_ITEM_SL_NEW sl,IQC_TestList iqc where 
                          id.CONTROL_ITEM_UNICODE=sl.CONTROL_ITEM_UNICODE and id.TestType=sl.TestType and id.TestItem=sl.TestItem 
                          and id.TestSubItem=sl.TestSubItem and id.CONTROL_ITEM_UNICODE=iqc.Productcode and id.TestType=iqc.TestType and id.TestItem=iqc.TestItem 
                          and id.TestSubItem=iqc.TestSubItem
                          and " + StrWhere + "  ";
            if (!string.IsNullOrEmpty(StrSql))
            {
                ds = DbAccess.SelectBySql(StrSql);
            }
            return ds;
        }
        private void btnsearch_Click(object sender, EventArgs e)
        {
            //this.databind.Rows.Clear();
            //dsSelect();
            //if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            //{
            //    dt = ds.Tables[0];
            //    for (int i = 0; i < dt.Rows.Count; i++)
            //    {
            //        this.databind.Rows.Add(1);
            //        this.databind.Rows[this.databind.Rows.Count - 1].Cells["CONTROL_ITEM_UNICODE"].Value = dt.Rows[i]["CONTROL_ITEM_UNICODE"].ToString();
            //        this.databind.Rows[this.databind.Rows.Count - 1].Cells["TestType"].Value = dt.Rows[i]["TestType"].ToString();
            //        this.databind.Rows[this.databind.Rows.Count - 1].Cells["TestItem"].Value = dt.Rows[i]["TestItem"].ToString();
            //        this.databind.Rows[this.databind.Rows.Count - 1].Cells["TestSubItem"].Value = dt.Rows[i]["TestSubItem"].ToString();
            //        this.databind.Rows[this.databind.Rows.Count - 1].Cells["UpValue"].Value = dt.Rows[i]["UpValue"].ToString();
            //        this.databind.Rows[this.databind.Rows.Count - 1].Cells["LowValue"].Value = dt.Rows[i]["LowValue"].ToString();
            //        this.databind.Rows[this.databind.Rows.Count - 1].Cells["TestStandardValue"].Value = dt.Rows[i]["TestStandardValue"].ToString();
            //        this.databind.Rows[this.databind.Rows.Count - 1].Cells["Uplimit"].Value = dt.Rows[i]["Uplimit"].ToString();
            //        this.databind.Rows[this.databind.Rows.Count - 1].Cells["Downlimit"].Value = dt.Rows[i]["Downlimit"].ToString();
            //        this.databind.Rows[this.databind.Rows.Count - 1].Cells["USL"].Value = dt.Rows[i]["USL"].ToString();
            //        this.databind.Rows[this.databind.Rows.Count - 1].Cells["LSL"].Value = dt.Rows[i]["LSL"].ToString();
            //        this.databind.Rows[this.databind.Rows.Count - 1].Cells["fileversion"].Value = dt.Rows[i]["fileversion"].ToString();
            //        this.databind.Rows[this.databind.Rows.Count - 1].Cells["EDIT_EMP"].Value = dt.Rows[i]["EDIT_EMP"].ToString();
            //        this.databind.Rows[this.databind.Rows.Count - 1].Cells["EDIT_TIME"].Value = dt.Rows[i]["EDIT_TIME"].ToString();

            //    }

            //}
            //else
            //{
            //    MessageBox.Show("没有符合条件的记录", "系统提醒", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}

            databind.DataSource = null;
            dsSelect();
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                dt = ds.Tables[0];
                databind.DataSource = dt;
            }
            else
            {
                MessageBox.Show("没有符合条件的记录", "系统提醒", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }























        }
    }
}