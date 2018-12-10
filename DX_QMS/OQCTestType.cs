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
    public partial class OQCTestType : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public string oldtestitem = "";
        private IQC ic = new IQC();
        public OQCTestType()
        {
            InitializeComponent();
        }
        private void setRule()
        {
            Dictionary<string, bool> dic = GroupPermission.SelectRulesForForm(this.Name, Login.groupId);
        }
        private void bindDeviceType()
        {
            string sql = "select checkType from OQC_CheckType";
            DataTable dt = Common.DbAccess.SelectBySql(sql).Tables[0];
            cbTestType.DataSource = dt;
            cbTestType.DisplayMember = dt.Columns["checkType"].ToString();
            cbTestType.ValueMember = dt.Columns["checkType"].ToString();

            this.btnSelect_Click(null, null);
        }
        private void OQCTestType_Load(object sender, EventArgs e)
        {
            bindDeviceType();
            txtTestItem.Select();
            setRule();
            dgvDefect.AutoGenerateColumns = false;
            dgvDefect.DataSource = null;
            if (dgvDefect.Rows.Count > 0)
            cbTestType.SelectedIndex = 0;
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
            txtconment.Text = "";
            this.txtTestItem.BackColor = Color.White;
            this.btnAdd.Text = "新 增";
        }
        private DataTable GetType(string checktype, string stype)
        {
            string sql = "select Definetype,Definevalue,sort,code from OQC_TypeDefine where Definetype='" + checktype + "' and Definevalue like '%" + stype + "%'";
            DataTable dt = Common.DbAccess.SelectBySql(sql).Tables[0];
            return dt;
        }
        private void btnSelect_Click(object sender, EventArgs e)
        {
            if (cbTestType.SelectedValue == null) return;
            DataTable dt = GetType(cbTestType.SelectedValue.ToString(), txtTestItem.Text.Trim());
            if (dt.Rows.Count > 0)
                dgvDefect.DataSource = dt;
            else
                dgvDefect.DataSource = null;
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
            int i = ic.AddOQCCheckItem("新增", cbTestType.SelectedValue.ToString(), txtTestItem.Text.Trim(), int.Parse(txtid.Text), oldtestitem, txtconment.Text);
            if (i > 0)
                dgvDefect.DataSource = GetType(cbTestType.SelectedValue.ToString(), txtTestItem.Text.Trim());
            oldtestitem = "";
            txtTestItem.BackColor = Color.White;
            btnAdd.Text = "新 增";
        }
        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (dgvDefect.SelectedRows.Count > 0)
            {
                if (MessageBox.Show("确定删除选中的测试项目？", "删除提示！", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    int temp = ic.AddOQCCheckItem("删除", dgvDefect.CurrentRow.Cells[0].Value.ToString(), dgvDefect.CurrentRow.Cells[1].Value.ToString(), 0, "", "");
                    if (temp > 0)
                        dgvDefect.DataSource = GetType(cbTestType.SelectedValue.ToString(), txtTestItem.Text.Trim());
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
            int i = ic.AddOQCCheckItem("更新", cbTestType.SelectedValue.ToString(), txtTestItem.Text.Trim(), int.Parse(txtid.Text), oldtestitem, txtconment.Text);
            if (i > 0)
                dgvDefect.DataSource = GetType(cbTestType.SelectedValue.ToString(), txtTestItem.Text.Trim());
            oldtestitem = "";
            txtTestItem.BackColor = Color.White;
            this.btnupdate.Enabled = false;
        }
        private void dgvDefect_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            cbTestType.SelectedValue = dgvDefect.CurrentRow.Cells["Definetype"].Value.ToString();
            txtTestItem.Text = dgvDefect.CurrentRow.Cells["Definevalue"].Value.ToString();
            txtid.Text = dgvDefect.CurrentRow.Cells["sort"].Value.ToString();
            oldtestitem = dgvDefect.CurrentRow.Cells["Definevalue"].Value.ToString();
            this.txtconment.Text = dgvDefect.CurrentRow.Cells["code"].Value.ToString();
            this.txtTestItem.BackColor = Color.Yellow;
            this.btnupdate.Enabled = true;
        }
    }
}