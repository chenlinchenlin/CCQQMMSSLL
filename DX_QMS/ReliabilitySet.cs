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
using DX_QMS.Common;

namespace DX_QMS
{
    public partial class ReliabilitySet : DevExpress.XtraEditors.XtraForm
    {
        IQC ic = new IQC();
        public ReliabilitySet()
        {
            InitializeComponent();
        }

        private void bindReliabilityType()
        {
            DataTable dt = ic.SelectTestTypeRecord("查询", "", "可靠性", "").Tables[0];
            DataTable dtt = ic.SelectTestTypeRecord("查询", "", "真实性", "").Tables[0];
            for (int i = 0; i < dtt.Rows.Count; i++)
            {
                dt.Rows.Add(dtt.Rows[i].ItemArray);
            }
            cbTestType.DataSource = dt;
            cbTestType.DisplayMember = dt.Columns["TestType"].ToString();
            cbTestType.ValueMember = dt.Columns["TestType"].ToString();
        }

        private DataSet BindReliabilitySet(string productcode)
        {
            string ssql = @"select Productcode 产品编码, rs.TestType 可靠性测试项,IFYiQi 测试内容, Productname 产品名称, lastupdatedate 最后更新日期, OperUser 操作员, isnull(TestCycle,90) 测试周期,isnull(LeadTime,10) 提前期 from  IQC_ReliabilitySet
                         rs left join (select * from IQC_TestType i where i.TTypes in('可靠性','真实性')) t on rs.TestType=t.TestType 
                         where Productcode like '" + productcode + "%'";
            DataSet ds = Common.DbAccess.SelectBySql(ssql);
            return ds;
        }
        private void txtproductcode_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13)
            {
                databind.DataSource = null;
                if (txtproductcode.Text == "") return;
                string Orasql = "select organization_id organization,segment1 materialcode,description materialname,en_des materialenname,INVENTORY_ITEM_ID,PRIMARY_UOM_CODE from cux_item_material_v where segment1='" + txtproductcode.Text.Trim() + "'";
                DataSet ds = Common.DbAccess.SelectByOracle(Orasql);
                if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                {
                    lblinfo.Text = ds.Tables[0].Rows[0]["materialname"].ToString();
                    lblinfo.ForeColor = Color.Blue;
                    //绑定已维护过的数据
                    DataSet dsprog = BindReliabilitySet(txtproductcode.Text);
                    if (dsprog != null && dsprog.Tables.Count > 0 && dsprog.Tables[0].Rows.Count > 0)
                    {
                        databind.DataSource = dsprog.Tables[0];

                    }
                    else
                        cbTestType.SelectedIndex = 0;
                    cbTestType.Focus();

                }
                else
                {
                    string ssql = "select materialcode from delivery where lotno='" + txtproductcode.Text + "'";
                    DataSet dslotno = Common.DbAccess.SelectBySql(ssql);
                    if (dslotno != null && ds.Tables.Count > 0 && dslotno.Tables[0].Rows.Count > 0)
                    {
                        string materialcode = dslotno.Tables[0].Rows[0]["materialcode"].ToString();
                        string Orasqlbylotno = "select organization_id organization,segment1 materialcode,description materialname,en_des materialenname,INVENTORY_ITEM_ID,PRIMARY_UOM_CODE from cux_item_material_v where segment1='" + materialcode + "'";
                        ds = Common.DbAccess.SelectByOracle(Orasqlbylotno);
                        if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                        {
                            txtproductcode.Text = materialcode;
                            lblinfo.Text = ds.Tables[0].Rows[0]["materialname"].ToString();
                            lblinfo.ForeColor = Color.Blue;
                            DataSet dsprog = BindReliabilitySet(txtproductcode.Text);
                            if (dsprog != null && dsprog.Tables.Count > 0 && dsprog.Tables[0].Rows.Count > 0)
                            {
                                databind.DataSource = dsprog.Tables[0];
                            }
                            else
                                cbTestType.SelectedIndex = 0;
                            cbTestType.Focus();

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
            }
        }

        ReliabilityTypeSet RTS;
        private void btnReliabilitytypeAdd_Click(object sender, EventArgs e)
        {
            if (RTS == null || RTS.IsDisposed)
            {
                RTS = new ReliabilityTypeSet();
                RTS.ShowDialog();
            }
            else
            {
                RTS.Activate();
                RTS.ShowDialog();
            }
            bindReliabilityType();
        }

        private void btnsave_Click(object sender, EventArgs e)
        {
            if (txtproductcode.Text == "") return;
            int m = 1, n = 1;
            if (!int.TryParse(txtcheckcycle.Text == "" ? "1" : txtcheckcycle.Text, out m))
            {
                MessageBox.Show("测试周期必须为数字", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                txtcheckcycle.Value = 0;
                txtcheckcycle.Focus();
                return;
            }
            else
            {
                if (m <= 0)
                {
                    MessageBox.Show("测试周期必须大于0", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    txtcheckcycle.Value = 0;
                    txtcheckcycle.Focus();
                    return;
                }
            }

            if (!int.TryParse(txtleadtime.Text == "" ? "1" : txtleadtime.Text, out m))
            {
                MessageBox.Show("提前期必须为数字", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                txtleadtime.Value = 0;
                txtleadtime.Focus();
                return;
            }
            else
            {
                if (m <= 0)
                {
                    MessageBox.Show("提前期必须大于0", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    txtleadtime.Value = 0;
                    txtleadtime.Focus();
                    return;
                }
            }

          int i = ic.IQC_ReliabilityOper("新增", txtproductcode.Text, lblinfo.Text, cbTestType.SelectedValue.ToString(), (int)txtcheckcycle.Value, (int)txtleadtime.Value, Login.username);

            databind.DataSource = BindReliabilitySet(txtproductcode.Text).Tables[0];
        }

        private void btnsearch_Click(object sender, EventArgs e)
        {
            databind.DataSource = BindReliabilitySet(txtproductcode.Text).Tables[0];
        }

        private void btndel_Click(object sender, EventArgs e)
        {
            DataTable de = databind.DataSource as DataTable;
            if (de == null || de.Rows.Count < 0)
                return;
            if (gridView.FocusedRowHandle < 0)
                return;
            if (gridView.GetSelectedRows().Length < 0)
                return;
            for (int k = gridView.GetSelectedRows().Length; k > 0; k--)
            {

            int i = ic.IQC_ReliabilityOper("删除", gridView.GetDataRow(gridView.GetSelectedRows()[k - 1])["产品编码"].ToString(), "",gridView.GetDataRow(gridView.GetSelectedRows()[k - 1])["可靠性测试项"].ToString(), 0, 0,Login.username);

            }
            databind.DataSource = BindReliabilitySet(txtproductcode.Text).Tables[0];
        }

        private void databind_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            //txtproductcode.Text = databind.CurrentRow.Cells["产品编码"].Value.ToString();
            //cbTestType.SelectedValue = databind.CurrentRow.Cells["可靠性测试项"].Value.ToString();
            //txtcheckcycle.Value = decimal.Parse(databind.CurrentRow.Cells["测试周期"].Value.ToString());
            //txtleadtime.Value = decimal.Parse(databind.CurrentRow.Cells["提前期"].Value.ToString());
        }

        private void gridView_Click(object sender, EventArgs e)
        {
            //txtproductcode.Text = databind.CurrentRow.Cells["产品编码"].Value.ToString();
            //cbTestType.SelectedValue = databind.CurrentRow.Cells["可靠性测试项"].Value.ToString();
            //txtcheckcycle.Value = decimal.Parse(databind.CurrentRow.Cells["测试周期"].Value.ToString());
            //txtleadtime.Value = decimal.Parse(databind.CurrentRow.Cells["提前期"].Value.ToString());
            DataTable de = databind.DataSource as DataTable;
            if (de == null || de.Rows.Count < 0)
                return;
            try
            {
                txtproductcode.Text = gridView.GetFocusedRowCellValue("产品编码").ToString();
                cbTestType.SelectedValue = gridView.GetFocusedRowCellValue("可靠性测试项").ToString();
                txtcheckcycle.Value = decimal.Parse(gridView.GetFocusedRowCellValue("测试周期").ToString());
                txtleadtime.Value = decimal.Parse(gridView.GetFocusedRowCellValue("提前期").ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show("没有数据！");
            }
        }
    }
}