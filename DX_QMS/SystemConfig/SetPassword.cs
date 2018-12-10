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
using System.Data;
using System.Data.SqlClient;
using DX_QMS.Common;
using System.Text.RegularExpressions;

namespace DX_QMS.SystemConfig
{
    public partial class SetPassword : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public SetPassword()
        {
            InitializeComponent();
            txtUserId.Text = Login.userId;
        }

        private void SetPassword_Load(object sender, EventArgs e)
        {
            txtUserId.Text = Login.userId;
            txtOldPwd.Select();
        }


        public int UpdatePasswordByUserId(string userId, string oldPwd, string newPwd)
        {
            SqlParameter[] para = new SqlParameter[3];
            para[0] = new SqlParameter("@userid", userId);
            para[1] = new SqlParameter("@oldpassword", oldPwd);
            para[2] = new SqlParameter("@newpassword", newPwd);

            return DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "QMS_UpdatePwd", para);
        }


        private void btnSave_Click(object sender, EventArgs e)
        {
            if (txtNewPwd1.Text.Trim() == "" || txtNewPwd2.Text.Trim() == "" || txtOldPwd.Text.Trim() == "")
                return;
            if (txtNewPwd1.Text.Trim() != txtNewPwd2.Text.Trim())
            {
                txtNewPwd2.Text = txtNewPwd1.Text = "";
                MessageBox.Show("两次输入的新密码不一致！");
                return;
            }
            int temp = UpdatePasswordByUserId(txtUserId.Text.Trim(), DbAccess.Encrypt(txtOldPwd.Text.Trim()), DbAccess.Encrypt(txtNewPwd1.Text.Trim()));
            if (temp > 0)
            {
                DialogResult result = MessageBox.Show("密码更改成功，是否关闭窗体？", "提示！", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (result == DialogResult.Yes)
                    this.Dispose();
                else
                    txtOldPwd.Text = txtNewPwd1.Text = txtNewPwd2.Text = "";
            }
            else
            {
                txtOldPwd.Text = "";
                MessageBox.Show("密码输入错误！");
            }
        }

        private void sBtn_reset_Click(object sender, EventArgs e)
        {
            txtOldPwd.Text = "";
            txtNewPwd1.Text = "";
            txtNewPwd2.Text = "";
            txtOldPwd.Focus();
        }
        private bool Test(string input)
        {
            return Regex.IsMatch(input, @"^[A-Za-z0-9_]*$");
        }





    }
}