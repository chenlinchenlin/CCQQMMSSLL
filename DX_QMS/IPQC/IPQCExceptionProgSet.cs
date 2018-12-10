using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Windows.Forms;
using DevExpress.XtraBars;
using DX_QMS.Common;

namespace DX_QMS.IPQC
{
    public partial class IPQCExceptionProgSet : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public IPQCExceptionProgSet()
        {
            InitializeComponent();
        }


        private void bindProgsettype()
        {
            string sql = "select distinct Progsettype from IPQCProgset  ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            if (dt==null || dt.Rows .Count <1 )
            {
                return;
            }
            txtProgsettype.Properties.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
                txtProgsettype.Properties.Items.Add(row["Progsettype"]);
            }
            txtProgsettype.SelectedIndex = 0;

        }

        private void setRule()
        {
            string post = "";
            if (!string.IsNullOrEmpty(Login.manager))
            {
                post = Login.manager;
            }
            else
            {
                post = Login.post;
            }
            Dictionary<string, bool> dic = GroupPermission.QMS_SelectRulesForForm(post, "制程信息");
            this.sBtnsave.Enabled = bool.Parse(dic["hasInsert"].ToString());
            this.sBtndelete.Enabled = bool.Parse(dic["hasDelete"].ToString());          
        }

        private void IPQCExceptionProgSet_Load(object sender, EventArgs e)
        {
            bindProgsettype();
            setRule();
        }

        private void txtProgsettype_SelectedIndexChanged(object sender, EventArgs e)
        {
            sBtnselect_Click(sender,e);
        }

        private void sBtnselect_Click(object sender, EventArgs e)
        {
            string where = " where 1=1 ";
            string Progsettype = txtProgsettype.Text.Trim();
            string Progsetvalue = txtProgsetvalue.Text.Trim();

            if (!string.IsNullOrEmpty(Progsettype))
            {
                where += " and Progsettype = '" + Progsettype + "' ";
            }
            if (!string.IsNullOrEmpty(Progsetvalue))
            {
                where += " and Progsetvalue = '" + Progsetvalue + "' ";
            }

            string sql = @" select Progsettype 类别,Progsetvalue 内容,remarks 备注,updateuser 更新人,updatetime 更新时间 from IPQCProgset  ";
            sql += where + " order by updatetime desc  ";

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

        private void sBtnsave_Click(object sender, EventArgs e)
        {
            if (txtProgsettype.Text.Trim() == "" || txtProgsetvalue.Text.Trim() == "")
            {
                MessageBox.Show("请输入类别、内容！", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            string sql = @" select 1 from IPQCProgset where Progsettype = '"+txtProgsettype.Text+ "' and  Progsetvalue = '"+ txtProgsetvalue.Text+ "' order by updatetime desc  ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            if (dt != null && dt.Rows.Count > 0)
            {
                MessageBox.Show("该类别内容已经存在！", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            string savasql = @"  insert into IPQCProgset (Progsettype,Progsetvalue,remarks,updateuser,updatetime)
			                        values ( '"+ txtProgsettype.Text + "','"+ txtProgsetvalue.Text + "','"+ txtremarks.Text+ "','"+Login.username+"',GETDATE())  ";

            bool flag = DbAccess.ExecuteSql(savasql);

            if (flag)
            {
                MessageBox.Show("保存成功", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("保存失败", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }

        private void sBtnupdate_Click(object sender, EventArgs e)
        {

            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
                return;
            if (gridView.FocusedRowHandle < 0)
                return;

        }

        private void sBtndelete_Click(object sender, EventArgs e)
        {
            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
                return;
            if (gridView.FocusedRowHandle < 0)
                return;
            string Progsettype = gridView.GetFocusedRowCellValue("类别").ToString();
            string Progsetvalue = gridView.GetFocusedRowCellValue("内容").ToString();
            string sql = @" delete IPQCProgset where Progsettype = '"+ Progsettype + "' and  Progsetvalue = '"+ Progsetvalue + "' ";

            bool flag = DbAccess.ExecuteSql(sql);

            if (flag)
            {
                MessageBox.Show("删除成功", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("删除失败", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }

        private void sBtnreset_Click(object sender, EventArgs e)
        {
             txtProgsettype.Text ="";
             txtProgsetvalue.Text = "";
             txtremarks.Text = "";
             gridControl.DataSource = null;
        }
    }
}