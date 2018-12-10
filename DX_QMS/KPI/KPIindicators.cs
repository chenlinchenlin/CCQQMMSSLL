using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Windows.Forms;
using DevExpress.XtraBars;
using DX_QMS.Common;

namespace DX_QMS.KPI
{
    public partial class KPIindicators : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public KPIindicators()
        {
            InitializeComponent();

            this.sBtnadd.Enabled = false ;
            this.sBtndelete.Enabled =false;
            this.sBtnupdate.Enabled = false;

            setRule();
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
            Dictionary<string, bool> dic = GroupPermission.QMS_SelectRulesForForm(post, "测试记录");
            this.sBtnadd.Enabled = bool.Parse(dic["hasInsert"].ToString());
            this.sBtndelete.Enabled = bool.Parse(dic["hasDelete"].ToString());
            this.sBtnupdate.Enabled = bool.Parse(dic["hasUpdate"].ToString());
        }


        private void sBtnselect_Click(object sender, EventArgs e)
        {

            string where = " where 1=1 ";
            string businessType = txtbusinessType.Text.Trim();
            string items = txtitems.Text.Trim();
            string indicatorsName = txtindicatorsName.Text.Trim();

            if (!string.IsNullOrEmpty(businessType))
            {
                where += " and businessType = '" + businessType + "' ";
            }
            //if (!string.IsNullOrEmpty(items))
            //{
            //    where += " and items = '" + items + "' ";
            //}
            if (!string.IsNullOrEmpty(indicatorsName))
            {
                where += " and indicatorsName like '%" + indicatorsName + "%' ";
            }


            string sql = @"  select businessType 业务分类,items 顺序,indicatorsName 指标名称,remarks 备注,updateMan 操作人,updateTime 时间  from QMS_KPIindicators  ";
            sql += where + " order by items asc ";

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

            gridView.ExpandAllGroups();
            gridView.ExpandGroupLevel(0);

        }


        public string AddNewTestROHS(string opertype, string businessType, string items, string indicatorsName, string remarks, string updateMan)
        {
            SqlParameter[] para = new SqlParameter[7];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@businessType", businessType);
            para[2] = new SqlParameter("@items", items);
            para[3] = new SqlParameter("@indicatorsName", indicatorsName);
            para[4] = new SqlParameter("@remarks", remarks);
            para[5] = new SqlParameter("@updateMan", updateMan);
            para[6] = new SqlParameter("@msg", SqlDbType.VarChar, 50);
            para[6].Direction = ParameterDirection.Output;
            DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "QMS_OperKPIindicators", para);
            return para[6].Value.ToString();
        }


        private void sBtnadd_Click(object sender, EventArgs e)
        {
            if (txtbusinessType.Text.Trim() == "" || txtitems.Text.Trim() == "" || txtindicatorsName.Text.Trim() == "")
            {
                MessageBox.Show("业务分类、顺序和指标名称不能为空", "提醒",MessageBoxButtons.OK ,MessageBoxIcon.Information);
                return;
            }

            string flag = AddNewTestROHS("新增", txtbusinessType.Text, txtitems.Text, txtindicatorsName.Text, txtRemark.Text, Login.username);
            if (flag.Contains("成功"))
            {
                MessageBox.Show(flag, "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("新增失败", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            sBtnselect_Click(null , null);
        }

        private void sBtndelete_Click(object sender, EventArgs e)
        {

            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
            {
                return;
            }
            int i = gridView.FocusedRowHandle;
            if (i < 0)
            {
                MessageBox.Show("请选中要删除的项目", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            string  businessType = gridView.GetFocusedRowCellValue("业务分类").ToString();
            string  txtitems = gridView.GetFocusedRowCellValue("顺序").ToString();

            string flag = AddNewTestROHS("删除", businessType, txtitems, "", "", "");

            if (flag.Contains("成功"))
            {
                MessageBox.Show(flag, "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            else
            {
                MessageBox.Show("删除失败", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }

        private void sBtnupdate_Click(object sender, EventArgs e)
        {

            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
            {
                return;
            }
            int i = gridView.FocusedRowHandle;
            if (i < 0)
            {
                MessageBox.Show("请选中要修改的项目", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            string businessType = gridView.GetFocusedRowCellValue("业务分类").ToString();
            string txtitems = gridView.GetFocusedRowCellValue("顺序").ToString();

            string flag = AddNewTestROHS("修改", businessType, txtitems, txtindicatorsName.Text, txtRemark.Text, Login.username);

            if (flag.Contains("成功"))
            {
                MessageBox.Show(flag, "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("修改失败", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            sBtnselect_Click(null, null);
        }

        private void sBtnreset_Click(object sender, EventArgs e)
        {
            txtbusinessType.Text = "";
            txtitems.Text = "";
            txtindicatorsName.Text = "";
            txtRemark.Text = "";
            gridControl.DataSource = null;
        }

        private void gridView_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {
            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
            {
                return;
            }
            if (gridView.RowCount < 1)
                return;
            try
            {
                txtbusinessType.Text = gridView.GetFocusedRowCellValue("业务分类").ToString();
                txtitems.Text = gridView.GetFocusedRowCellValue("顺序").ToString();
                txtindicatorsName.Text = gridView.GetFocusedRowCellValue("指标名称").ToString();
                txtRemark.Text = gridView.GetFocusedRowCellValue("备注").ToString();
            }
            catch
            {


            }
        }
    }
}