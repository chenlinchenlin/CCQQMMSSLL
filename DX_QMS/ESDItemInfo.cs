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
    public partial class ESDItemInfo : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public string oldtestitem = "";
        private IQC ic = new IQC();
        public ESDItemInfo()
        {
            InitializeComponent();
            setRule();
        }

        private void ESDItemInfo_Load(object sender, EventArgs e)
        {
            bindESDType();
            txtESDItem.Select();
            setRule();            
            dgvDefect.DataSource = null;
            if (gridView.RowCount > 0)
            cbESDType.SelectedIndex = 0;
        }
        private void setRule()
        {
            string post = "";
            if (Login.manager != null && Login.manager != "")
            {
                post = Login.manager;
            }
            else
            {
                post = Login.post;
            }
            Dictionary<string, bool> dic = GroupPermission.QMS_SelectRulesForForm(post, "项目维护");
            btnAdd.Enabled = bool.Parse(dic["hasInsert"].ToString());
            btnupdate.Enabled = bool.Parse(dic["hasUpdate"].ToString());     //btnupdate
            btnDelete.Enabled = bool.Parse(dic["hasDelete"].ToString());
        }
        private void bindESDType()
        {
            string sql = "select Dtype from ESD_TypeItem order by item ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            cbESDType.Properties.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
                cbESDType.Properties.Items.Add(row["Dtype"]);
            }
            cbESDType.SelectedIndex = 0;
            this.btnSelect_Click(null, null);
        }

        private void cbESDType_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnSelect_Click(sender, e);
        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            dgvDefect.DataSource = null;
            if (cbESDType.Enabled)
                cbESDType.SelectedIndex = -1;
            txtESDItem.Text = "";
            txtid.Text = "";
            oldtestitem = "";
            this.txtESDItem.BackColor = Color.White;
            this.btnAdd.Text = "新 增";
        }

        private void btnSelect_Click(object sender, EventArgs e)
        {
            if (cbESDType.Text == "")
                return;
            DataSet ds = ic.SelectESDItemRecord("查询", cbESDType.Text, txtESDItem.Text.Trim(), 0, "");
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                dgvDefect.DataSource = ds.Tables[0];
            else
                dgvDefect.DataSource = null;
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            if (txtESDItem.Text.Trim() == "") return;
            int m = 0;
            if (!int.TryParse(txtid.Text, out m))
            {
                MessageBox.Show("顺序请输入数字");
                return;
            }
            int i = ic.AddNewESDItemRecord("新增", cbESDType.Text, txtESDItem.Text.Trim(), int.Parse(txtid.Text), oldtestitem);
            if (i > 0)
                dgvDefect.DataSource = ic.SelectESDItemRecord("查询", cbESDType.Text, txtESDItem.Text.Trim(), 0, "").Tables[0];
            oldtestitem = "";
            txtESDItem.BackColor = Color.White;
            btnAdd.Text = "新 增";
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            //if (dgvDefect.SelectedRows.Count > 0)
            //{
            //    if (MessageBox.Show("确定删除选中的测试项目？", "删除提示！", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            //    {
            //        int temp = ic.AddNewESDItemRecord("删除", dgvDefect.CurrentRow.Cells[0].Value.ToString(), dgvDefect.CurrentRow.Cells[1].Value.ToString(), 0, "");
            //        if (temp > 0)
            //            dgvDefect.DataSource = ic.SelectESDItemRecord("查询", cbESDType.SelectedValue.ToString(), txtESDItem.Text.Trim(), 0, "").Tables[0];
            //    }
            //}
            if (gridView.RowCount < 1)
                return;
            if (gridView.GetSelectedRows().Length > 0)
            {
                if (MessageBox.Show("确定删除选中的测试项目？", "删除提示！", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    int temp = ic.AddNewESDItemRecord("删除", gridView.GetFocusedRowCellValue("Dtype").ToString(), gridView.GetFocusedRowCellValue("Dvalue").ToString(), 0, "");
                    if (temp > 0)
                        dgvDefect.DataSource = ic.SelectESDItemRecord("查询", cbESDType.Text, txtESDItem.Text.Trim(), 0, "").Tables[0];
                }
            }
        }

        private void btnupdate_Click(object sender, EventArgs e)
        {
            if (txtESDItem.Text.Trim() == "") return;
            int m = 0;
            if (!int.TryParse(txtid.Text, out m))
            {
                MessageBox.Show("顺序请输入数字");
                return;
            }
            int i = ic.AddNewESDItemRecord("更新", cbESDType.Text, txtESDItem.Text.Trim(), int.Parse(txtid.Text), oldtestitem);
            if (i > 0)
                dgvDefect.DataSource = ic.SelectESDItemRecord("查询", cbESDType.Text, txtESDItem.Text.Trim(), 0, "").Tables[0];
            oldtestitem = "";
            txtESDItem.BackColor = Color.White;
            this.btnupdate.Enabled = false;
        }

        private void dgvDefect_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //cbESDType.SelectedValue = dgvDefect.CurrentRow.Cells["Dtype"].Value.ToString();
            //txtESDItem.Text = dgvDefect.CurrentRow.Cells["Dvalue"].Value.ToString();
            //txtid.Text = dgvDefect.CurrentRow.Cells["Item"].Value.ToString();
            //oldtestitem = dgvDefect.CurrentRow.Cells["Dvalue"].Value.ToString();
            //this.txtESDItem.BackColor = Color.Yellow;
            //this.btnupdate.Enabled = true;

        }

        private void gridView_Click(object sender, EventArgs e)
        {
            //cbESDType.Text = gridView.GetFocusedRowCellValue("Dtype").ToString();
            //txtESDItem.Text = gridView.GetFocusedRowCellValue("Dvalue").ToString();
            //txtid.Text = gridView.GetFocusedRowCellValue("Item").ToString();
            //oldtestitem = gridView.GetFocusedRowCellValue("Dvalue").ToString();
            //this.txtESDItem.BackColor = Color.Yellow;
            //this.btnupdate.Enabled = true;
        }

        private void gridView_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {
            if (gridView.RowCount < 1)
                return;
            try
            {
                cbESDType.Text = gridView.GetFocusedRowCellValue("Dtype").ToString();
                txtESDItem.Text = gridView.GetFocusedRowCellValue("Dvalue").ToString();
                txtid.Text = gridView.GetFocusedRowCellValue("Item").ToString();
                oldtestitem = gridView.GetFocusedRowCellValue("Dvalue").ToString();
                this.txtESDItem.BackColor = Color.Yellow;
                this.btnupdate.Enabled = true;
            }
            catch
            {

            }
        }
    }
}