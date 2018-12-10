using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using DX_QMS.Common;

namespace DX_QMS.SystemConfig
{
    public partial class PostMaintain : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public PostMaintain()
        {
            InitializeComponent();
            bindData();
            txtpost.Focus();
        }
        private void bindData()
        {
            string sql = @"select groupId 岗位ID,groupName 岗位名称,groupDescribe 岗位描述,updateTime 更新时间,updateUser 更新人 from QMS_groupMaintain";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            gridControl.DataSource = dt;
        }

        private void sBtnreset_Click(object sender, EventArgs e)
        {
            txtpost.Text="";
            txtpostdescribe.Text ="";
            gridControl.DataSource = null;
        }

        private void sBtnAdd_Click(object sender, EventArgs e)
        {
            if (txtpost.Text != null)
            {
                string sql = @"select groupName from QMS_groupMaintain where groupName='" + txtpost.Text + "'";
                DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
                if (dt != null && dt.Rows.Count > 0)
                {
                    MessageBox.Show("该岗位已存在，不能重复增加");
                    return;
                }
                else
                {
                    string sql_save = @"insert into QMS_groupMaintain(groupName,groupDescribe,updateTime,updateUser)"
                                    + "values('" + txtpost.Text + "','" + txtpostdescribe.Text + "','" + DateTime.Now + "','" + Login.username + "')";  // Login.username 


                    DbAccess.ExecuteSql(sql_save);

                }
                bindData();

            }
        }

        private void sBtndelete_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("确认删除？", "Confirm Message", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                string groupName  = gridView.GetFocusedRowCellValue("岗位名称").ToString();
                string sql = "delete from QMS_groupMaintain where groupName ='" +groupName+"'";
                bool falg =   DbAccess.ExecuteSql(sql);
                if (falg)
                    MessageBox.Show("删除成功！","提醒",MessageBoxButtons.OK,MessageBoxIcon.Information);
                else
                    MessageBox.Show("删除失败！", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                bindData();
            }
        }

        private void gridView_Click(object sender, EventArgs e)
        {
            int i = gridView.FocusedRowHandle;
            if (i < 0)
            {
                MessageBox.Show("没有数据", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null && dt.Rows.Count < 0)
            {
                MessageBox.Show("没有数据", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            txtpost.Text = gridView.GetFocusedRowCellValue("岗位名称").ToString();
            txtpostdescribe.Text = gridView.GetFocusedRowCellValue("岗位描述").ToString();

        }
    }
}