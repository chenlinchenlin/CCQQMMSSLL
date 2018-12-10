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
    public partial class SetUser : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private string sUserID = "";
        public string UserID
        {
            get { return sUserID; }
            set { sUserID = value; }
        }
        public SetUser()
        {
            InitializeComponent();
            gbUser.Text = "新增用户";
            bingDept(Login.groupId, Login.deptId);
            bindDept2();
            bingGroup("", Login.deptId);
        }
        public SetUser(string userId, string userName, string tel, string deptId, string groupId, string deptid2, string groupBSid)
        {
            InitializeComponent();
            gbUser.Text = "修改用户";
            label5.Visible = label6.Visible = false;

            txtUserId.Text = userId;
            txtUserId.ReadOnly = true;
            txtUserName.Text = userName;
            txtTel.Text = tel;
            bingDept(Login.groupId, deptId);
            bindDept2();
            bingGroup(Login.groupId, Login.deptId);

            cbDept.SelectedValue = deptId;
            cbGroup.SelectedValue = groupId;
            cbBsGroup.SelectedValue = groupBSid;
            cbDept2.SelectedValue = deptid2;
        }
        protected void bingDept(string groupid, string deptid)
        {
            cbDept.SelectedIndexChanged -= cbDept_SelectedIndexChanged;
            DataSet ds = null;
            if (Login.deptId == "IT")
                ds = Dept.SelectAll(groupid, "IT");
            else
                ds = Dept.SelectAll(groupid, deptid);
            if (ds.Tables.Count > 0)
            {
                cbDept.DataSource = ds.Tables[0];
                cbDept.DisplayMember = ds.Tables[0].Columns["deptname"].ToString();
                cbDept.ValueMember = ds.Tables[0].Columns["deptid"].ToString();

                cbDept.SelectedIndex = -1;
            }
            cbDept.SelectedIndexChanged += cbDept_SelectedIndexChanged;
        }
        protected void bindDept2()
        {
            DataTable dt = Dept.SelectInfoByCondition("", "", "").Tables[0];
            dt.Rows.Add("", "空", "");

            cbDept2.DataSource = dt;
            cbDept2.DisplayMember = dt.Columns["deptname"].ToString();
            cbDept2.ValueMember = dt.Columns["deptid"].ToString();

            cbDept2.SelectedIndex = cbDept2.Items.Count - 1;
        }
        protected void bingGroup(string groupid, string deptid)
        {
            DataSet ds = Groups.SelectAll(groupid, deptid, "N");
            if (ds.Tables[0].Rows.Count > 0)
            {
                cbGroup.DataSource = ds.Tables[0];
                cbGroup.DisplayMember = ds.Tables[0].Columns["groupname"].ToString();
                cbGroup.ValueMember = ds.Tables[0].Columns["groupid"].ToString();

                cbGroup.SelectedIndex = -1;
            }

            ds = null;
            ds = Groups.SelectAll(groupid, deptid, "Y");
            if (ds.Tables[0].Rows.Count > 0)
            {
                cbBsGroup.DataSource = ds.Tables[0];
                cbBsGroup.DisplayMember = ds.Tables[0].Columns["groupname"].ToString();
                cbBsGroup.ValueMember = ds.Tables[0].Columns["groupid"].ToString();

                cbBsGroup.SelectedIndex = -1;
            }
        }
        private void cbDept_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbDept.SelectedIndex != -1)
            {
                cbBsGroup.DataSource = cbGroup.DataSource = null;
                bingGroup("", cbDept.SelectedValue.ToString());
            }
        }

        private void SetUser_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
            this.Close();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (txtUserId.Text.Trim() == "" || txtUserName.Text.Trim() == "" || cbDept.Text == "") return;
            if (cbGroup.SelectedIndex < 0 && cbBsGroup.SelectedIndex < 0) return;

            if (txtUserId.ReadOnly == false)
            {
                int insert = Users.AddRecordByKey(txtUserId.Text.Trim(), txtUserName.Text.Trim(), DbAccess.Encrypt("123456"), cbGroup.SelectedValue.ToString(), cbDept.SelectedValue.ToString(), txtTel.Text.Trim(), cbDept2.SelectedValue.ToString(), cbBsGroup.SelectedValue == null ? "" : cbBsGroup.SelectedValue.ToString());
                if (insert > 0)
                    MessageBox.Show("新增< " + txtUserId.Text.Trim() + " >用户成功！", "新增提示！", MessageBoxButtons.OK, MessageBoxIcon.Information);
                else
                {
                    MessageBox.Show("新增< " + txtUserId.Text.Trim() + " >用户失败,账号已存在！", "新增提示！", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txtUserId.Text = txtUserName.Text = txtTel.Text = "";
                    cbDept.SelectedIndex = cbGroup.SelectedIndex = cbDept2.SelectedIndex = cbBsGroup.SelectedIndex = -1;
                    return;
                }
            }
            else
            {
                int update = Users.UpdateByKey(txtUserId.Text.Trim(), txtUserName.Text.Trim(), cbGroup.SelectedValue.ToString(), cbDept.SelectedValue.ToString(), txtTel.Text.Trim(), cbDept2.SelectedValue.ToString(), cbBsGroup.SelectedValue == null ? "" : cbBsGroup.SelectedValue.ToString());
                if (update > 0)
                {
                    MessageBox.Show("修改成功！", "修改提示！", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            sUserID = txtUserId.Text.Trim();
            this.Dispose();
            this.Close();
        }
    }
}