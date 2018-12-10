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

namespace DX_QMS.SystemConfig
{
    public partial class UsersInfo : DevExpress.XtraBars.Ribbon.RibbonForm
    {
         SetUser setU;

        public UsersInfo()
        {
            InitializeComponent();
        }

        private void UsersInfo_Load(object sender, EventArgs e)
        {
            txtId.Select();
            setRule();
            dgvUsers.AutoGenerateColumns = false;
            dgvUsers.DataSource = null;
        }
        private void setRule()
        {
            Dictionary<string, bool> dic = GroupPermission.SelectRulesForForm("UsersInfo", Login.groupId);
            btnAdd.Enabled = dic["hasInsert"];
            btnUpdate.Enabled = btnSetPwd.Enabled = dic["hasUpdate"];
            btnDelete.Enabled = dic["hasDelete"];
        }
        public static void SetControlEmpty(Control contrs)
        {
            foreach (Control con in contrs.Controls)
            {
                if (con is DataGridView)
                    ((DataGridView)con).DataSource = null;
                else if (con.HasChildren)
                {
                    SetControlEmpty(con);
                }
                else
                {
                    if (con.GetType().Name == "TextBox")
                        con.Text = "";
                }
            }
        }
        private void btnReset_Click(object sender, EventArgs e)
        {
            SetControlEmpty(this);
        }

        private void btnSelect_Click(object sender, EventArgs e)
        {
            DataSet ds = Users.SelectByConditon(txtId.Text.Trim(), txtName.Text.Trim(), Login.deptId);
            dgvUsers.DataSource = ds.Tables[0];
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            if (setU == null || setU.IsDisposed)
                setU = new SetUser();
                setU.ShowDialog();

            dgvUsers.DataSource = Users.SelectByConditon(txtId.Text.Trim(), txtName.Text.Trim(), Login.deptId).Tables[0];
            if (dgvUsers.Rows.Count > 0)
                dgvUsers.CurrentCell = dgvUsers.Rows[DbAccess.SelectStringInDGV(setU.UserID, dgvUsers, 0)].Cells[0];
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            if (!btnUpdate.Enabled)
                return;
            if (dgvUsers.SelectedRows.Count > 0)
            {
                DataGridViewCellCollection dgvCells = dgvUsers.CurrentRow.Cells;
                if (setU == null || setU.IsDisposed)
                    setU = new SetUser(dgvCells[0].Value.ToString(), dgvCells[1].Value.ToString(), dgvCells[2].Value.ToString(), dgvCells[6].Value.ToString(), dgvCells[7].Value.ToString(), dgvCells[4].Value.ToString(), dgvCells[8].Value.ToString());
                setU.ShowDialog();

                dgvUsers.DataSource = Users.SelectByConditon(txtId.Text.Trim(), txtName.Text.Trim(), Login.deptId).Tables[0];
                if (dgvUsers.Rows.Count > 0)
                    dgvUsers.CurrentCell = dgvUsers.Rows[DbAccess.SelectStringInDGV(setU.UserID, dgvUsers, 0)].Cells[0];
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (dgvUsers.SelectedRows.Count > 0)
            {
                DialogResult result = MessageBox.Show("确定删除< " + dgvUsers.CurrentRow.Cells[0].Value.ToString() + " >用户吗？", "删除提示！", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.Yes)
                {
                    int temp = Users.DeleteByKey(dgvUsers.CurrentRow.Cells[0].Value.ToString().Trim());
                    if (temp > 0)
                        dgvUsers.DataSource = Users.SelectByConditon(txtId.Text.Trim(), txtName.Text.Trim(), Login.deptId).Tables[0];
                }
            }
        }

        private void btnSetPwd_Click(object sender, EventArgs e)
        {
            if (dgvUsers.SelectedRows.Count > 0)
            {
                if (MessageBox.Show("确定初始化< " + dgvUsers.CurrentRow.Cells[0].Value.ToString() + " >的密码吗？", "密码初始化提示！", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    int temp = Users.ClearPassword(dgvUsers.CurrentRow.Cells[0].Value.ToString().Trim(), DbAccess.Encrypt("123456"));
                    if (temp > 0)
                        MessageBox.Show("密码初始化成功！");
                }
            }
        }
    }
}