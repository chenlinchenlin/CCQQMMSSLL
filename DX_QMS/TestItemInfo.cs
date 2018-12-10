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
    public partial class TestItemInfo : DevExpress.XtraEditors.XtraForm
    {
        public string oldtestitem = "";
        private IQC ic = new IQC();
        public TestItemInfo()
        {
            InitializeComponent();
        }

        private void setRule()
        {
           // Dictionary<string, bool> dic = GroupPermission.SelectRulesForForm(this.Name, Login.groupId);


        }

        private void TestItemInfo_Load(object sender, EventArgs e)
        {
            bindDeviceType();
            txtTestItem.Select();
            setRule();
            //dgvDefect.AutoGenerateColumns = false;
            dgvDefect.DataSource = null;
            DataTable de = dgvDefect.DataSource as DataTable;
            if ( de!=null && de.Rows.Count > 0)
            cbTestType.SelectedIndex = 0;
        }
        private void bindDeviceType()
        {
            DataTable dt = ic.SelectTestTypeRecord("查询", "", "测试类别", "").Tables[0];
            cbTestType.DataSource = dt;
            cbTestType.DisplayMember = dt.Columns["TestType"].ToString();
            cbTestType.ValueMember = dt.Columns["TestType"].ToString();
            this.btnSelect_Click(null, null);
        }

        private void btnSelect_Click(object sender, EventArgs e)
        {
            if (cbTestType.SelectedValue == null) return;
            DataSet ds = ic.SelectTestItemRecord("查询", cbTestType.SelectedValue.ToString(), txtTestItem.Text.Trim(), 0, "");
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                dgvDefect.DataSource = ds.Tables[0];
            else
                dgvDefect.DataSource = null;
        }

        private void cbTestType_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnSelect_Click(sender, e);
        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            dgvDefect.DataSource = null;
            if (cbTestType.Enabled)
                cbTestType.SelectedIndex = -1;
            txtTestItem.Text = "";
            txtid.Text = "";
            oldtestitem = "";
            this.txtTestItem.BackColor = Color.White;
            this.btnAdd.Text = "新 增";
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            if (txtTestItem.Text.Trim() == "") return;
            int m = 0;
            if (!int.TryParse(txtid.Text, out m))
            {
                MessageBox.Show("顺序请输入数字");
                return;
            }
            int i = ic.AddNewTestItemRecord("新增", cbTestType.SelectedValue.ToString(), txtTestItem.Text.Trim(), int.Parse(txtid.Text), oldtestitem);
            if (i > 0)
                dgvDefect.DataSource = ic.SelectTestItemRecord("查询", cbTestType.SelectedValue.ToString(), txtTestItem.Text.Trim(), 0, "").Tables[0];
            oldtestitem = "";
            txtTestItem.BackColor = Color.White;
            btnAdd.Text = "新 增";
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            DataTable de = dgvDefect.DataSource as DataTable;
            if (de == null || de.Rows.Count < 0)
                return;
            if (gridView.FocusedRowHandle < 0)
                return;
            if (gridView.GetSelectedRows().Length > 0)        ////gridView.GetFocusedRowCellValue("TestType").ToString();
            {
                if (MessageBox.Show("确定删除选中的测试项目？", "删除提示！", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    int temp = ic.AddNewTestItemRecord("删除", gridView.GetFocusedRowCellValue("TestType").ToString(), gridView.GetFocusedRowCellValue("TestItem").ToString(), 0, "");
                    if (temp > 0)
                        dgvDefect.DataSource = ic.SelectTestItemRecord("查询", cbTestType.SelectedValue.ToString(), txtTestItem.Text.Trim(), 0, "").Tables[0];
                }
            }
        }

        private void btnupdate_Click(object sender, EventArgs e)
        {
            if (txtTestItem.Text.Trim() == "") return;
            int m = 0;
            if (!int.TryParse(txtid.Text, out m))
            {
                MessageBox.Show("顺序请输入数字");
                return;
            }
            int i = ic.AddNewTestItemRecord("更新", cbTestType.SelectedValue.ToString(), txtTestItem.Text.Trim(), int.Parse(txtid.Text), oldtestitem);
            if (i > 0)
                dgvDefect.DataSource = ic.SelectTestItemRecord("查询", cbTestType.SelectedValue.ToString(), txtTestItem.Text.Trim(), 0, "").Tables[0];
            oldtestitem = "";
            txtTestItem.BackColor = Color.White;
            this.btnupdate.Enabled = false;
        }
        private void dgvDefect_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //cbTestType.SelectedValue = dgvDefect.CurrentRow.Cells["TestType"].Value.ToString();
            //txtTestItem.Text = dgvDefect.CurrentRow.Cells["TestItem"].Value.ToString();
            //txtid.Text = dgvDefect.CurrentRow.Cells["Item"].Value.ToString();
            //oldtestitem = dgvDefect.CurrentRow.Cells["TestItem"].Value.ToString();
            //this.txtTestItem.BackColor = Color.Yellow;
            //this.btnupdate.Enabled = true;
        }

        private void gridView_Click(object sender, EventArgs e)
        {
            //cbTestType.SelectedValue = dgvDefect.CurrentRow.Cells["TestType"].Value.ToString();
            //txtTestItem.Text = dgvDefect.CurrentRow.Cells["TestItem"].Value.ToString();
            //txtid.Text = dgvDefect.CurrentRow.Cells["Item"].Value.ToString();
            //oldtestitem = dgvDefect.CurrentRow.Cells["TestItem"].Value.ToString();
            //this.txtTestItem.BackColor = Color.Yellow;
            //this.btnupdate.Enabled = true;
            DataTable de = dgvDefect.DataSource as DataTable;
            if (de == null || de.Rows.Count < 0)
                return;
            cbTestType.SelectedValue = gridView.GetFocusedRowCellValue("TestType").ToString();
            txtTestItem.Text = gridView.GetFocusedRowCellValue("TestItem").ToString();
            txtid.Text = gridView.GetFocusedRowCellValue("Item").ToString();
            oldtestitem = gridView.GetFocusedRowCellValue("TestItem").ToString();
            this.txtTestItem.BackColor = Color.Yellow;
            this.btnupdate.Enabled = true;
        }
    }
}