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
    public partial class TestSubItemInfo : DevExpress.XtraEditors.XtraForm
    {
        public string oldtestsubitem = "";
        private IQC ic = new IQC();
        public TestSubItemInfo()
        {
            InitializeComponent();
        }

        private void TestSubItemInfo_Load(object sender, EventArgs e)
        {
            bindDeviceType();
            bindTestTools();
            txtTestsubItem.Select();
           // dgvDefect.AutoGenerateColumns = false;
            dgvDefect.DataSource = null;
            DataTable de = dgvDefect.DataSource as DataTable;
            if (de!=null && de.Rows.Count > 0)
                cbTestType.SelectedIndex = 0;
            if (cbTestType.SelectedValue == null) return;
            bindTestItem(cbTestType.SelectedValue.ToString());
        }
        private void bindDeviceType()
        {
            DataTable dt = ic.SelectTestTypeRecord("查询", "", "测试类别", "").Tables[0];
            cbTestType.DataSource = dt;
            cbTestType.DisplayMember = dt.Columns["TestType"].ToString();
            cbTestType.ValueMember = dt.Columns["TestType"].ToString();
            btnSelect_Click(null, null);
        }
        private void bindTestTools()
        {
            DataTable dt = ic.SelectTestTypeRecord("查询", "", "测试工具", "").Tables[0];
            txttesttools.DataSource = dt;
            txttesttools.DisplayMember = dt.Columns["TestType"].ToString();
            txttesttools.ValueMember = dt.Columns["TestType"].ToString();
        }
        private void bindTestItem(string TestType)
        {
            DataSet ds = ic.SelectTestItemRecord("查询", TestType, "", 0, "");
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                DataTable dt = ds.Tables[0];
                txttestitem.DataSource = dt;
                txttestitem.DisplayMember = dt.Columns["TestItem"].ToString();
                txttestitem.ValueMember = dt.Columns["TestItem"].ToString();

                btnSelect_Click(null, null);
            }
            else
                txttestitem.DataSource = null;
        }

        private void btnSelect_Click(object sender, EventArgs e)
        {
            if (cbTestType.SelectedValue == null) return;
            DataSet ds = ic.SelectTestSubItemRecord("查询", cbTestType.SelectedValue.ToString(), txttestitem.SelectedValue == null ? "" : txttestitem.SelectedValue.ToString(), txtTestsubItem.Text.Trim(), txtTestDes.Text,
                txttesttools.SelectedValue == null ? "" : txttesttools.SelectedValue.ToString(), txtPacktype.Text, "", 0, 0, 0, "", 0, 0, "");
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                dgvDefect.DataSource = ds.Tables[0];
            }
            else
                dgvDefect.DataSource = null;
        }

        private void cbTestType_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbTestType.SelectedValue == null) return;
            bindTestItem(cbTestType.SelectedValue.ToString());
        }

        private void txttestitem_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnSelect_Click(sender, e);
        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            Common.DbAccess.SetControlEmpty(this);
            dgvDefect.DataSource = null;
            if (cbTestType.Enabled)
                cbTestType.SelectedIndex = -1;

            txtTestsubItem.BackColor = Color.White;
            txtTestsubItem.Text = "";
            oldtestsubitem = "";
            this.btnAdd.Text = "新 增";
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            int x = 0;
            if (txtPacktype.Text.Trim() == "")
            {
                MessageBox.Show("请选择包装方式");
                return;
            }
            if (txtTestsubItem.Text.Trim() == "")
            {
                MessageBox.Show("请输入子项目");
                return;
            }
            if (int.TryParse(txtTestsubItem.Text, out x))
            {
                MessageBox.Show("子项目不能输入纯数字");
                return;
            }
            if (txttesttools.SelectedValue == null)
            {
                MessageBox.Show("请选择检测工具");
                return;
            }
            int m = 0;
            if (!int.TryParse(txtid.Text == "" ? "0" : txtid.Text, out m))
            {
                MessageBox.Show("请输入数字顺序");
                return;
            }
            int i = ic.AddNewTestSubItemRecord("新增", cbTestType.SelectedValue.ToString(), txttestitem.SelectedValue.ToString(), txtTestsubItem.Text.Trim(), txtTestDes.Text,
                txttesttools.SelectedValue.ToString(), txtPacktype.Text, "", 0, 0, 0, "", int.Parse(txtid.Text == "" ? "0" : txtid.Text), 0, oldtestsubitem);
            if (i > 0)
                dgvDefect.DataSource = ic.SelectTestSubItemRecord("查询", cbTestType.SelectedValue.ToString(), txttestitem.SelectedValue == null ? "" : txttestitem.SelectedValue.ToString(), txtTestsubItem.Text.Trim(), txtTestDes.Text,
                txttesttools.SelectedValue.ToString(), txtPacktype.Text, "", 0, 0, 0, "", 0, 0, "").Tables[0];
            oldtestsubitem = "";
            this.txtTestsubItem.BackColor = Color.White;
            btnAdd.Text = "新 增";
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (gridView.FocusedRowHandle < 0)
                return;
            if (gridView.GetSelectedRows().Length > 0)
            {
                if (MessageBox.Show("确定删除选中的测试项目？", "删除提示！", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    int temp = ic.AddNewTestSubItemRecord("删除", gridView.GetFocusedRowCellValue("TestType").ToString(), gridView.GetFocusedRowCellValue("TestItem").ToString(), gridView.GetFocusedRowCellValue("TestSubItem").ToString(), "", "", gridView.GetFocusedRowCellValue("PackType").ToString(), "", 0, 0, 0, "", 0, 0, "");
                    if (temp > 0)
                        dgvDefect.DataSource = ic.SelectTestSubItemRecord("查询", cbTestType.SelectedValue.ToString(), txttestitem.SelectedValue == null ? "" : txttestitem.SelectedValue.ToString(), txtTestsubItem.Text.Trim(), txtTestDes.Text,
                txttesttools.SelectedValue.ToString(), txtPacktype.Text, "", 0, 0, 0, "", 0, 0, "").Tables[0];
                }
            }
        }

        private void btnupdate_Click(object sender, EventArgs e)
        {
            int i = ic.AddNewTestSubItemRecord("更新", cbTestType.SelectedValue.ToString(), txttestitem.SelectedValue.ToString(), txtTestsubItem.Text.Trim(), txtTestDes.Text,
           txttesttools.SelectedValue.ToString(), txtPacktype.Text, "", 0, 0, 0, "", int.Parse(txtid.Text == "" ? "0" : txtid.Text), 0, oldtestsubitem);
            if (i > 0)
                dgvDefect.DataSource = ic.SelectTestSubItemRecord("查询", cbTestType.SelectedValue.ToString(), txttestitem.SelectedValue == null ? "" : txttestitem.SelectedValue.ToString(), txtTestsubItem.Text.Trim(), txtTestDes.Text,
                txttesttools.SelectedValue.ToString(), txtPacktype.Text, "", 0, 0, 0, "", 0, 0, "").Tables[0];
            oldtestsubitem = "";
            this.txtTestsubItem.BackColor = Color.White;
            this.btnupdate.Enabled = false;
        }

        private void dgvDefect_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //cbTestType.SelectedValue = dgvDefect.CurrentRow.Cells["TestType"].Value.ToString();
            //txttestitem.SelectedValue = dgvDefect.CurrentRow.Cells["TestItem"].Value.ToString();
            //txtTestsubItem.Text = dgvDefect.CurrentRow.Cells["TestSubItem"].Value.ToString();
            //txtTestDes.Text = dgvDefect.CurrentRow.Cells["TestDesc"].Value.ToString();
            //txttesttools.SelectedValue = dgvDefect.CurrentRow.Cells["TestTool"].Value.ToString();
            //txtPacktype.Text = dgvDefect.CurrentRow.Cells["PackType"].Value.ToString();
            //txtid.Text = dgvDefect.CurrentRow.Cells["Item"].Value.ToString();
            //oldtestsubitem = dgvDefect.CurrentRow.Cells["TestSubItem"].Value.ToString();
            //this.txtTestsubItem.BackColor = Color.Yellow;
            //this.btnupdate.Enabled = true;
        }

        private void dgvDefect_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //cbTestType.SelectedValue = dgvDefect.CurrentRow.Cells["TestType"].Value.ToString();
            //txttestitem.SelectedValue = dgvDefect.CurrentRow.Cells["TestItem"].Value.ToString();
            //txtTestsubItem.Text = dgvDefect.CurrentRow.Cells["TestSubItem"].Value.ToString();
            //txtTestDes.Text = dgvDefect.CurrentRow.Cells["TestDesc"].Value.ToString();
            //txttesttools.SelectedValue = dgvDefect.CurrentRow.Cells["TestTool"].Value.ToString();
            //txtPacktype.Text = dgvDefect.CurrentRow.Cells["PackType"].Value.ToString();
            //txtid.Text = dgvDefect.CurrentRow.Cells["Item"].Value.ToString();
            //oldtestsubitem = dgvDefect.CurrentRow.Cells["TestSubItem"].Value.ToString();
            //this.txtTestsubItem.BackColor = Color.Yellow;
        }

        private void gridView_Click(object sender, EventArgs e)
        {
            //cbTestType.SelectedValue = dgvDefect.CurrentRow.Cells["TestType"].Value.ToString();
            //txttestitem.SelectedValue = dgvDefect.CurrentRow.Cells["TestItem"].Value.ToString();
            //txtTestsubItem.Text = dgvDefect.CurrentRow.Cells["TestSubItem"].Value.ToString();
            //txtTestDes.Text = dgvDefect.CurrentRow.Cells["TestDesc"].Value.ToString();
            //txttesttools.SelectedValue = dgvDefect.CurrentRow.Cells["TestTool"].Value.ToString();
            //txtPacktype.Text = dgvDefect.CurrentRow.Cells["PackType"].Value.ToString();
            //txtid.Text = dgvDefect.CurrentRow.Cells["Item"].Value.ToString();
            //oldtestsubitem = dgvDefect.CurrentRow.Cells["TestSubItem"].Value.ToString();
            //this.txtTestsubItem.BackColor = Color.Yellow;
            //this.btnupdate.Enabled = true;
            DataTable de = dgvDefect.DataSource as DataTable;
            if (de == null || de.Rows.Count < 1)
                return;
            if (gridView.FocusedRowHandle < 0)
                return;
            cbTestType.SelectedValue = gridView.GetFocusedRowCellValue("TestType").ToString();
            txttestitem.SelectedValue = gridView.GetFocusedRowCellValue("TestItem").ToString();
            txtTestsubItem.Text = gridView.GetFocusedRowCellValue("TestSubItem").ToString();
            txtTestDes.Text = gridView.GetFocusedRowCellValue("TestDesc").ToString();
            txttesttools.SelectedValue = gridView.GetFocusedRowCellValue("TestTool").ToString();
            txtPacktype.Text = gridView.GetFocusedRowCellValue("PackType").ToString();
            txtid.Text = gridView.GetFocusedRowCellValue("Item").ToString();
            oldtestsubitem = gridView.GetFocusedRowCellValue("TestSubItem").ToString();
            this.txtTestsubItem.BackColor = Color.Yellow;
            this.btnupdate.Enabled = true;
        }

        private void gridView_DoubleClick(object sender, EventArgs e)
        {
            //cbTestType.SelectedValue = dgvDefect.CurrentRow.Cells["TestType"].Value.ToString();
            //txttestitem.SelectedValue = dgvDefect.CurrentRow.Cells["TestItem"].Value.ToString();
            //txtTestsubItem.Text = dgvDefect.CurrentRow.Cells["TestSubItem"].Value.ToString();
            //txtTestDes.Text = dgvDefect.CurrentRow.Cells["TestDesc"].Value.ToString();
            //txttesttools.SelectedValue = dgvDefect.CurrentRow.Cells["TestTool"].Value.ToString();
            //txtPacktype.Text = dgvDefect.CurrentRow.Cells["PackType"].Value.ToString();
            //txtid.Text = dgvDefect.CurrentRow.Cells["Item"].Value.ToString();
            //oldtestsubitem = dgvDefect.CurrentRow.Cells["TestSubItem"].Value.ToString();
            //this.txtTestsubItem.BackColor = Color.Yellow;
            DataTable de = dgvDefect.DataSource as DataTable;
            if (de == null || de.Rows.Count < 1)
                return;
            if (gridView.FocusedRowHandle < 0)
                return;
            cbTestType.SelectedValue = gridView.GetFocusedRowCellValue("TestType").ToString();
            txttestitem.SelectedValue = gridView.GetFocusedRowCellValue("TestItem").ToString();
            txtTestsubItem.Text = gridView.GetFocusedRowCellValue("TestSubItem").ToString();
            txtTestDes.Text = gridView.GetFocusedRowCellValue("TestDesc").ToString();
            txttesttools.SelectedValue = gridView.GetFocusedRowCellValue("TestTool").ToString();
            txtPacktype.Text = gridView.GetFocusedRowCellValue("PackType").ToString();
            txtid.Text = gridView.GetFocusedRowCellValue("Item").ToString();
            oldtestsubitem = gridView.GetFocusedRowCellValue("TestSubItem").ToString();
            this.txtTestsubItem.BackColor = Color.Yellow;
        }
    }
}