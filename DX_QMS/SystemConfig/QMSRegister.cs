using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using DX_QMS.Common;
using System.Text.RegularExpressions;

namespace DX_QMS.SystemConfig
{
    public partial class QMSRegister : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public QMSRegister()
        {
            InitializeComponent();
        }

        private void QMSRegister_Load(object sender, EventArgs e)
        {
            bindData();
        }

        private void bindData()
        {
            comgroupname.Properties.Items.Clear();
            string sql = "";
            if (Login.manager == "IT管理员")
            {
                sql = @"select distinct groupName from QMS_groupMaintain ";
            }
            else
            {
              string manager = string.IsNullOrEmpty(Login.manager) ? "" : Login.post;
                sql = @"select distinct groupName from QMS_groupMaintain  where groupName = '"+manager+ "'";
            }
         
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            if (dt != null && dt.Rows.Count > 0)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    comgroupname.Properties.Items.Add(dr["groupName"]);
                }
            }        
        }
        private bool queryRegistere()
        {
            bool flag = false;
            string sql = @"select *  from QMS_UserInfo where userId='" + txtuserid.Text + "'";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            if (dt != null && dt.Rows.Count > 0)
            {
                flag = true;
            }

            return flag;
        }

        private bool Test(string input)
        {
            return Regex.IsMatch(input, @"^[A-Za-z0-9_]*$");
        }

        private void sBtnsave_Click(object sender, EventArgs e)
        {

            if (Test(txtuserid.Text) == false || Test(txtpassword.Text) == false)
            {
                MessageBox.Show("用户名和密码只能输入数字、字母、下划线");
                return;
            }

            if (txtuserid.Text == "" || txtpassword.Text == "" || txtusername.Text == "" || comgroupname.Text == "")
            {
                MessageBox.Show("信息填写不完整");
                return;
            }
            //先做判断，不能重复注册
            if (queryRegistere() == true)
            {
                MessageBox.Show("用户已注册，不能重复注册");
                return;
            }

            else   /////DbAccess.Encrypt(txtpassword.Text.Trim())
            {
                string sql = @"insert into QMS_UserInfo(userId,passWord,userName,tel,mail,post,updateUser,updateTime)"
                            + "values('" + txtuserid.Text + "','" + DbAccess.Encrypt(txtpassword.Text.Trim()) + "','" + txtusername.Text + "','" + txttelcontact.Text + "','" + txtmail.Text + "','" + comgroupname.Text + "','" + Login.username + "','" + DateTime.Now + "')";  //Login.username
                if (DbAccess.ExecuteSql(sql) == true)
                {
                    MessageBox.Show("保存成功");
                }

            }
            string ssql = @"select userId 用户工号,userName 用户名称,post 岗位名称,tel 用户电话,mail 用户邮箱,updateUser 更新人,updateTime 更新时间  from QMS_UserInfo where userId='" + txtuserid.Text + "'";
            DataTable dt = DbAccess.SelectBySql(ssql).Tables[0];
            gridControl.DataSource = dt;


        }

        private void sBtndelete_Click(object sender, EventArgs e)
        {
            if (txtuserid.Text == "")
            {
                return;
            }
            string sql = @"delete from QMS_UserInfo where userId='" + txtuserid.Text + "'";
            if (DbAccess.ExecuteSql(sql) == true)
            {
                MessageBox.Show("删除成功");
            }

        }

        private void sBtnselect_Click(object sender, EventArgs e)
        {
            string sqlstr = "", where = " where updateUser='"+ Login.username + "'";
            string userid = txtuserid.Text.Trim(), username = txtusername.Text.Trim();
           ///////string updateman = Login.username;

         
            if (!string.IsNullOrEmpty(userid))
            {
                where += " and userId = '" + userid + "' ";
            }
            if (!string.IsNullOrEmpty(username))
            {
                where += " and userName like '%" + username + "%' ";
            }
            string sql = @" select userId 用户工号,userName 用户名称,post 岗位名称,tel 用户电话,mail 用户邮箱,updateUser 更新人,updateTime 更新时间  from QMS_UserInfo  ";
            sql += where + " order by updateTime desc ";

            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];

            if (dt != null && dt.Rows.Count > 0)
            {
                gridControl.DataSource = dt;
            }
            else
            {
                MessageBox.Show("没有符合条件的记录");
                gridControl.DataSource = null;
            }

        }
    }
}